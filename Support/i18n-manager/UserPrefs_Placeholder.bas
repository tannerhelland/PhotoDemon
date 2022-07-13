Attribute VB_Name = "UserPrefs"
'Placeholder module to allow me to use PhotoDemon code in non-PhotoDemon projects

Option Explicit

Public Function IsReady() As Boolean
    IsReady = True
End Function

Public Function GetTempPath() As String
    GetTempPath = OS.SystemTempPath()
End Function
