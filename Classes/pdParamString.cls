VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdParamString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Specialized String Parameter Handler
'Copyright ©2012-2013 by Tanner Helland
'Created: 25/March/13
'Last updated: 25/March/13
'Last update: initial build
'
'PhotoDemon has unique needs regarding parameter passing.  Because it allows the user to record all actions as part of
' macros, the program needs a way to not only trade around parameters, but also store them in a file.
'
'The problem occurs when realizing how many different parameters a macro can handle - some actions, such as "Invert Image",
' require no additional parameters.  Others, like "Image curves", require a huge amount of data to be passed from the
' tool form to the actual function.  Functions utilize every type of value (bool, int, float, etc), and if a macro is
' being recorded, all those various parameters need to be written to file.  Rather than attempt to deal with a convoluted
' system for tracking those, as was previously used, PhotoDemon now sticks everything inside a string.  This allows any
' amount - and type - of data to be stored.
'
'Individual functions therefore rely on this class to parse parameter strings for them.
'
'***************************************************************************

Option Explicit

Dim m_ParamString As String
Dim m_Params() As Variant
Dim m_numOfParams As Long

'Assign the class a parameter string, and parse it as necessary into the
Public Sub setParamString(ByVal pString As String)
    m_ParamString = pString
    If InStr(1, m_ParamString, "|", vbBinaryCompare) > 0 Then
        m_Params = Split(m_ParamString, "|")
        m_numOfParams = UBound(m_Params)
    Else
        ReDim m_Params(0) As Variant
        m_Params(0) = m_ParamString
        m_numOfParams = 1
    End If
End Sub

'Retrieve a bool-type parameter
Public Function GetBool(ByVal pIndex As Long) As Boolean

    If m_numOfParams = 1 Then
        GetBool = CBool(m_ParamString)
    Else
        If pIndex <= m_numOfParams Then
            GetBool = CBool(m_Params(pIndex - 1))
        Else
            GetBool = False
        End If
    End If

End Function

'Retrieve a byte-type parameter
Public Function GetByte(ByVal pIndex As Long) As Byte

    If m_numOfParams = 1 Then
        GetByte = CByte(m_ParamString)
    Else
        If pIndex <= m_numOfParams Then
            GetByte = CByte(m_Params(pIndex - 1))
        Else
            GetByte = 0
        End If
    End If

End Function

'Retrieve an int-type parameter
Public Function GetInteger(ByVal pIndex As Long) As Integer

    If m_numOfParams = 1 Then
        GetInteger = CInt(m_ParamString)
    Else
        If pIndex <= m_numOfParams Then
            GetInteger = CInt(m_Params(pIndex - 1))
        Else
            GetInteger = 0
        End If
    End If

End Function

'Retrieve a long-type parameter
Public Function GetLong(ByVal pIndex As Long) As Long

    If m_numOfParams = 1 Then
        GetLong = CLng(m_ParamString)
    Else
        If pIndex <= m_numOfParams Then
            GetLong = CLng(m_Params(pIndex - 1))
        Else
            GetLong = 0
        End If
    End If

End Function

'Retrieve a single-type parameter
Public Function GetSingle(ByVal pIndex As Long) As Single

    If m_numOfParams = 1 Then
        GetSingle = CSng(m_ParamString)
    Else
        If pIndex <= m_numOfParams Then
            GetSingle = CSng(m_Params(pIndex - 1))
        Else
            GetSingle = 0#
        End If
    End If

End Function

'Retrieve a double-type parameter
Public Function GetDouble(ByVal pIndex As Long) As Double

    If m_numOfParams = 1 Then
        GetDouble = CDbl(m_ParamString)
    Else
        If pIndex <= m_numOfParams Then
            GetDouble = CDbl(m_Params(pIndex - 1))
        Else
            GetDouble = 0#
        End If
    End If

End Function
