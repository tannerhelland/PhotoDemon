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
'Copyright 2013-2015 by Tanner Helland
'Created: 25/March/13
'Last updated: 03/July/14
'Last update: added failsafe conversions for variable decimal characters in values retrieved as Longs.  (Just because a
'              function requests a Long does not necessarily mean the stored value is a Long; this fix ensures that
'              regardless of locale, the subsequent CLng() function won't throw a type mismatch error.)
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
    
    'If multiple parameters are present, parse them out now; this improves performance when a function comes
    ' asking for specific params
    If InStr(1, m_ParamString, "|", vbBinaryCompare) > 0 Then
        m_Params() = Split(m_ParamString, "|")
        m_numOfParams = UBound(m_Params) + 1
    
    'If only one param is present, handle it specially
    Else
        If Len(pString) <> 0 Then
            ReDim m_Params(0) As String
            m_Params(0) = m_ParamString
            m_numOfParams = 1
        Else
            m_numOfParams = 0
        End If
    End If
    
    'If any characters have been escaped, fix them now
    If m_numOfParams > 0 Then
        
        Dim i As Long
        For i = 0 To UBound(m_Params)
            m_Params(i) = unEscapeParamCharacters(m_Params(i))
        Next i
        
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
    
    'Replace the entry at pIndex with the new string value
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
        tmpString = tmpString & escapeParamCharacters(m_Params(i))
        If i < (m_numOfParams - 1) Then tmpString = tmpString & "|"
    Next i
    
    'Re-initialize this object with that string
    setParamString tmpString

End Sub

'Remove a parameter from the current param string.  This is typically used when a function needs to preface a param string
' with some kind of target ID (e.g. layer ID for actions that affect layers), but then it needs to forward the rest of
' the param string onto that target for additional handling.
Public Sub removeParamAtPosition(ByVal pIndex As Long)

    'If there's only one parameter in the list, we can't remove it!
    If m_numOfParams = 1 Then Exit Sub
    
    'Loop through the current param array, shifting array entries down as necessary
    Dim i As Long
    For i = pIndex - 1 To m_numOfParams - 2
        m_Params(i) = m_Params(i + 1)
    Next i
    
    'Reassemble the modified param string
    reassembleModifiedParamString

End Sub

'Retrieve a bool-type parameter
Public Function GetBool(ByVal pIndex As Long, Optional ByVal defaultReturn As Boolean = False) As Boolean

    If m_numOfParams = 1 Then
        GetBool = CBool(m_ParamString)
    Else
        If doesParamExist(pIndex) Then
            GetBool = CBool(Trim$(m_Params(pIndex - 1)))
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
        GetInteger = CInt(CDblCustom(m_ParamString))
    Else
        If doesParamExist(pIndex) Then
            GetInteger = CInt(CDblCustom(m_Params(pIndex - 1)))
        Else
            GetInteger = defaultReturn
        End If
    End If

End Function

'Retrieve a long-type parameter
Public Function GetLong(ByVal pIndex As Long, Optional ByVal defaultReturn As Long = 0) As Long

    If m_numOfParams = 1 Then
        GetLong = CLng(CDblCustom(m_ParamString))
    Else
        If doesParamExist(pIndex) Then
            GetLong = CLng(CDblCustom(m_Params(pIndex - 1)))
        Else
            GetLong = defaultReturn
        End If
    End If

End Function

'Retrieve a single-type parameter
Public Function GetSingle(ByVal pIndex As Long, Optional ByVal defaultReturn As Single = 0#) As Single

    If m_numOfParams = 1 Then
        GetSingle = CDblCustom(m_ParamString)
    Else
        If doesParamExist(pIndex) Then
            GetSingle = CDblCustom(m_Params(pIndex - 1))
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

'Retrieve a variant-type parameter
Public Function GetVariant(ByVal pIndex As Long, Optional ByVal defaultReturn As Variant = Empty) As Variant

    If m_numOfParams = 1 Then
        GetVariant = m_ParamString
    Else
        If doesParamExist(pIndex) Then
            GetVariant = CVar(m_Params(pIndex - 1))
        Else
            GetVariant = defaultReturn
        End If
    End If

End Function

Private Sub Class_Initialize()
    m_ParamString = ""
    m_numOfParams = 0
End Sub
