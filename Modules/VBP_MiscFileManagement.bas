Attribute VB_Name = "Misc_FileInteractions"
'***************************************************************************
'Miscellaneous Functions Related to File and Folder Interactions
'Copyright ©2000-2012 by Tanner Helland
'Created: 6/12/01
'Last updated: 13/November/12
'Last update: Updated DirectoryExist to not only check for a directory's existence, but also make sure we have access rights.
'
'***************************************************************************

Option Explicit

'Used to quickly check if a file (or folder) exists
Private Const ERROR_SHARING_VIOLATION As Long = 32
Private Declare Function GetFileAttributesW Lib "kernel32" (ByVal lpFileName As Long) As Long

'Returns a boolean as to whether or not a given file exists
Public Function FileExist(ByRef fName As String) As Boolean
    Select Case (GetFileAttributesW(StrPtr(fName)) And vbDirectory) = 0
        Case True: FileExist = True
        Case Else: FileExist = (Err.LastDllError = ERROR_SHARING_VIOLATION)
    End Select
End Function

'Returns a boolean as to whether or not a given directory exists AND whether we have write access to it or not.
' (If we do not have write access, the function will return "False".)
Public Function DirectoryExist(ByRef dName As String) As Boolean
    
    'First, make sure the directory exists
    Dim chkExistence As Boolean
    chkExistence = Abs(GetFileAttributesW(StrPtr(dName))) And vbDirectory
        
    'Next, make sure we have write access
    On Error GoTo noWriteAccess
    
    If chkExistence Then
        
        Dim tmpFilename As String
        tmpFilename = FixPath(dName) & "tmp.tmp"
        
        Dim fileNum As Integer
        fileNum = FreeFile
    
        'Attempt to create a file within this directory.  If we succeed, delete the file and return "true".
        ' If we fail, we do not have access rights.
        Open tmpFilename For Binary As #fileNum
            Put #fileNum, 1, "0"
        Close #fileNum
        
        If FileExist(tmpFilename) Then Kill tmpFilename
        
        DirectoryExist = True
        Exit Function
        
    End If
    
noWriteAccess:

    DirectoryExist = False
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

'Open a string as a hyperlink in the user's default browser
Public Sub OpenURL(ByVal targetURL As String)
    ShellExecute FormMain.hWnd, "Open", targetURL, "", 0, SW_SHOWNORMAL
End Sub
