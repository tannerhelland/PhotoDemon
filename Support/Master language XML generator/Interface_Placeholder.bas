Attribute VB_Name = "Interface"
'Placeholder module to allow me to use PhotoDemon code in non-PhotoDemon projects

Option Explicit

Public Sub Message(ByVal mString As String, ParamArray ExtraText() As Variant)
    Debug.Print mString
End Sub

Public Sub NotifySystemDialogState(ByVal dialogIsActive As Boolean)
    'Do nothing
End Sub

Public Function PDMsgBox(ByVal pMessage As String, ByVal pButtons As VbMsgBoxStyle, ByVal pTitle As String, ParamArray ExtraText() As Variant) As VbMsgBoxResult
    PDMsgBox = MsgBox(pMessage, pButtons, pTitle)
End Function

'Version retrieval (from PD Updates module)


'Given an arbitrary version string (e.g. "6.0.04 stability patch" or 6.0.04" or just plain "6.0"), return a canonical major/minor string, e.g. "6.0"
Public Function RetrieveVersionMajorMinorAsString(ByVal srcVersionString As String) As String

    'To avoid locale issues, replace any "," with "."
    If InStr(1, srcVersionString, ",") Then srcVersionString = Replace$(srcVersionString, ",", ".")
    
    'For this function to work, the major/minor data has to exist somewhere in the string.  Look for at least one "." occurrence.
    Dim tmpArray() As String
    tmpArray = Split(srcVersionString, ".")
    
    If (UBound(tmpArray) >= 1) Then
        RetrieveVersionMajorMinorAsString = Trim$(tmpArray(0)) & "." & Trim$(tmpArray(1))
    Else
        RetrieveVersionMajorMinorAsString = vbNullString
    End If

End Function

'Given an arbitrary version string (e.g. "6.0.04 stability patch" or 6.0.04" or just plain "6.0"), return the revision number
' as a string, e.g. 4 for "6.0.04".  If no revision is found, return 0.
Public Function RetrieveVersionRevisionAsLong(ByVal srcVersionString As String) As Long
    
    'An improperly formatted version number can cause failure; if this happens, we'll assume a revision of 0, which should
    ' force a re-download of the problematic file.
    On Error GoTo CantFormatRevisionAsLong
    
    'To avoid locale issues, replace any "," with "."
    If InStr(1, srcVersionString, ",") Then srcVersionString = Replace$(srcVersionString, ",", ".")
    
    'For this function to work, the revision has to exist somewhere in the string.  Look for at least two "." occurrences.
    Dim tmpArray() As String
    tmpArray = Split(srcVersionString, ".")
    
    If (UBound(tmpArray) >= 2) Then
        RetrieveVersionRevisionAsLong = CLng(Trim$(tmpArray(2)))
    
    'If one or less "." chars are found, assume a revision of 0
    Else
        RetrieveVersionRevisionAsLong = 0
    End If
    
    Exit Function
    
CantFormatRevisionAsLong:
    RetrieveVersionRevisionAsLong = 0

End Function

