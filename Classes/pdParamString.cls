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
'Copyright ©2013-2014 by Tanner Helland
'Created: 25/March/13
'Last updated: 07/June/13
'Last update: added support for String-type params
'
'PhotoDemon has unique needs regarding parameter passing.  Because it allows the user to record all actions as part of
' macros, the program needs a way to not only trade around parameters, but also store them in a file.
'
'The problem occurs when realizing how many different parameters a macro can handle - some actions, such as "Invert Image",
' require no additional parameters.  Others, like "Image curves", require a huge amount of data to be passed from the
' tool form to the actual function.  Functions utilize every type of value (bool, int, float, etc), and if a macro is
' being recorded, all those various parameters need to be written to file.  Rather than attempt to deal with a convoluted
' system for tracking each unique data type, as we previously did, the latest versions of PhotoDemon now stick all
' parameters inside a string.  This allows any amount - and type - of data to be stored.
'
'Individual functions therefore rely on this class to parse parameter strings for them.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

Private m_ParamString As String
Private m_Params() As String
Private m_numOfParams As Long

'Return the number of parameters in the current param string.  This function has no effect if a param string
' has not yet been assigned to the class.
Public Function getParamCount() As Long
    getParamCount = m_numOfParams
End Function

Public Function getParamString() As String
    getParamString = m_ParamString
End Function

'Assign the class a parameter string, and parse it as necessary into its individual components
Public Sub setParamString(ByVal pString As String)
    m_ParamString = pString
    If InStr(1, m_ParamString, "|", vbBinaryCompare) > 0 Then
        m_Params() = Split(m_ParamString, "|")
        m_numOfParams = UBound(m_Params) + 1
    Else
        If Len(pString) > 0 Then
            ReDim m_Params(0) As String
            m_Params(0) = m_ParamString
            m_numOfParams = 1
        Else
            m_numOfParams = 0
        End If
    End If
End Sub

'Does the parameter at location X exist?
Public Function doesParamExist(ByVal pIndex As Long) As Boolean
    If m_numOfParams = 0 Then
        doesParamExist = False
    Else
        If pIndex <= m_numOfParams And pIndex > 0 Then
            doesParamExist = True
        Else
            doesParamExist = False
        End If
    End If
End Function

'Change the value at a given location.  (This assumes setParamString has already been called, obviously.)
Public Sub setNewValueAtLocation(ByVal pIndex As Long, ByVal newValue As String)

    If pIndex <= m_numOfParams Then
        m_Params(pIndex - 1) = newValue
    End If

End Sub

'If setNewValueAtLocation has been used, this sub needs to be called to reassemble the modified parameter string
Public Sub reassembleModifiedParamString()

    'Build a new param string by concatenating all values into a single "|"-delimited string
    Dim tmpString As String
    tmpString = ""
    
    Dim i As Long
    For i = 0 To m_numOfParams - 1
        tmpString = tmpString & m_Params(i)
        If i < (m_numOfParams - 1) Then tmpString = tmpString & "|"
    Next i
    
    'Re-initialize this object with that string
    setParamString tmpString

End Sub

'Retrieve a bool-type parameter
Public Function GetBool(ByVal pIndex As Long, Optional ByVal defaultReturn As Boolean = False) As Boolean

    If m_numOfParams = 1 Then
        GetBool = CBool(m_ParamString)
    Else
        If doesParamExist(pIndex) Then
            GetBool = CBool(m_Params(pIndex - 1))
        Else
            GetBool = defaultReturn
        End If
    End If

End Function

'Retrieve a byte-type parameter
Public Function GetByte(ByVal pIndex As Long, Optional ByVal defaultReturn As Byte = 0) As Byte

    If m_numOfParams = 1 Then
        GetByte = CByte(m_ParamString)
    Else
        If doesParamExist(pIndex) Then
            GetByte = CByte(m_Params(pIndex - 1))
        Else
            GetByte = defaultReturn
        End If
    End If

End Function

'Retrieve an int-type parameter
Public Function GetInteger(ByVal pIndex As Long, Optional ByVal defaultReturn As Integer = 0) As Integer

    If m_numOfParams = 1 Then
        GetInteger = CInt(m_ParamString)
    Else
        If doesParamExist(pIndex) Then
            GetInteger = CInt(m_Params(pIndex - 1))
        Else
            GetInteger = defaultReturn
        End If
    End If

End Function

'Retrieve a long-type parameter
Public Function GetLong(ByVal pIndex As Long, Optional ByVal defaultReturn As Long = 0) As Long

    If m_numOfParams = 1 Then
        GetLong = CLng(m_ParamString)
    Else
        If doesParamExist(pIndex) Then
            GetLong = CLng(m_Params(pIndex - 1))
        Else
            GetLong = defaultReturn
        End If
    End If

End Function

'Retrieve a single-type parameter
Public Function GetSingle(ByVal pIndex As Long, Optional ByVal defaultReturn As Single = 0#) As Single

    If m_numOfParams = 1 Then
        GetSingle = CSng(m_ParamString)
    Else
        If doesParamExist(pIndex) Then
            GetSingle = CSng(m_Params(pIndex - 1))
        Else
            GetSingle = defaultReturn
        End If
    End If

End Function

'Retrieve a double-type parameter
Public Function GetDouble(ByVal pIndex As Long, Optional ByVal defaultReturn As Double = 0#) As Double

    If m_numOfParams = 1 Then
        GetDouble = CDblCustom(m_ParamString)
    Else
        If doesParamExist(pIndex) Then
            GetDouble = CDblCustom(m_Params(pIndex - 1))
        Else
            GetDouble = defaultReturn
        End If
    End If

End Function

'Retrieve a string-type parameter
Public Function GetString(ByVal pIndex As Long, Optional ByVal defaultReturn As String = "") As String

    If m_numOfParams = 1 Then
        GetString = m_ParamString
    Else
        If doesParamExist(pIndex) Then
            GetString = m_Params(pIndex - 1)
        Else
            GetString = defaultReturn
        End If
    End If

End Function

Private Sub Class_Initialize()
    m_ParamString = ""
    m_numOfParams = 0
End Sub
