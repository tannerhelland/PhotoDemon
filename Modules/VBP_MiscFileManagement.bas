Attribute VB_Name = "Misc_FileInteractions"
'***************************************************************************
'Miscellaneous Functions Related to File and Folder Interactions
'Copyright ©2000-2012 by Tanner Helland
'Created: 6/12/01
'Last updated: 03/October/12
'Last update: First build
'***************************************************************************

Option Explicit

'Returns a boolean as to whether or not a given file exists
Public Function FileExist(ByRef fName As String) As Boolean
    On Error Resume Next
    Dim Temp As Long
    Temp = GetAttr(fName)
    FileExist = Not CBool(Err)
End Function

'Returns a boolean as to whether or not a given directory exists
Public Function DirectoryExist(ByRef dName As String) As Boolean
    On Error Resume Next
    Dim Temp As Long
    Temp = GetAttr(dName) And vbDirectory
    DirectoryExist = Not CBool(Err)
End Function

'Straight from MSDN - generate a "browse for folder" dialog
Public Function BrowseForFolder(ByVal srcHwnd As Long) As String
    
    Dim objShell As Shell
    Dim objFolder As Folder
    Dim returnString As String
        
    Set objShell = New Shell
    Set objFolder = objShell.BrowseForFolder(srcHwnd, "Please select a folder:", 0)
            
    If (Not objFolder Is Nothing) Then returnString = objFolder.Items.Item.Path Else returnString = ""
    
    Set objFolder = Nothing
    Set objShell = Nothing
    
    BrowseForFolder = returnString
    
End Function
