Attribute VB_Name = "File_And_Path_Handling"
'***************************************************************************
'Miscellaneous Functions Related to File and Folder Interactions
'Copyright ©2000-2013 by Tanner Helland
'Created: 6/12/01
'Last updated: 13/November/12
'Last update: Updated DirectoryExist to not only check for a directory's existence, but also make sure we have access rights.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://www.tannerhelland.com/photodemon/#license
'
'***************************************************************************

Option Explicit

'Used to quickly check if a file (or folder) exists
Private Const ERROR_SHARING_VIOLATION As Long = 32
Private Declare Function GetFileAttributesW Lib "kernel32" (ByVal lpFileName As Long) As Long

'Used to shell an external program, then wait until it completes
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Const SYNCHRONIZE = &H100000
Private Const WAIT_INFINITE = -1&

'If a file exists, this function can be used to intelligently increment the file name (e.g. "filename (n+1).ext")
' Note that the function returns the filename WITHOUT an extension so that it can be passed to a common dialog without
' further parsing. However, an initial extension is required, because this function should only be used if a file name
' with that extension exists (as it is perfectly fine to have the same filename with DIFFERENT extensions in the
' same directory).
Public Function incrementFilename(ByRef dstDirectory As String, ByRef fName As String, ByRef desiredExtension As String) As String

    'First, check to see if a file with that name and extension appears in the destination directory.
    ' If it does, just return the filename we were passed.
    If Not FileExist(dstDirectory & fName & "." & desiredExtension) Then
        incrementFilename = fName
        Exit Function
    End If

    'If we made it to this line of code, a file with that name and extension appears in the destination directory.
    
    'Start by figuring out if the file is already in the format: "filename (#).ext"
    Dim tmpFilename As String
    tmpFilename = Trim(fName)
    
    Dim numToAppend As Long
    
    'Check the trailing character.  If it is a closing parentheses ")", we need to analyze more
    If Right(tmpFilename, 1) = ")" Then
    
        Dim I As Long
        For I = Len(tmpFilename) - 2 To 1 Step -1
            
            ' If it isn't a number, see if it's an initial parentheses: "("
            If Not (IsNumeric(Mid(tmpFilename, I, 1))) Then
                
                'If it is a parentheses, then this file already has a "( #)" appended to it.  Figure out what the
                ' number inside the parentheses is, and strip that entire block from the filename.
                If Mid(tmpFilename, I, 1) = "(" Then
                
                    numToAppend = CLng(Val(Mid(tmpFilename, I + 1, Len(tmpFilename) - I - 1)))
                    tmpFilename = Left(tmpFilename, I - 2)
                    Exit For
                
                'If this character is non-numeric and NOT an initial parentheses, this filename is not in the format we want.
                ' Treat it like any other filename and start by appending " (2)" to it
                Else
                    numToAppend = 2
                    Exit For
                End If
                
            End If
        
        'If this character IS a number, keep scanning.
        Next I
    
    'If this is not already a copy of the format "filename (#).ext", start scanning at # = 2
    Else
        numToAppend = 2
    End If
            
    'Loop through
    Do While FileExist(dstDirectory & tmpFilename & " (" & CStr(numToAppend) & ")" & "." & desiredExtension)
        numToAppend = numToAppend + 1
    Loop
        
    'If the loop has terminated, a unique filename has been found.  Make that the recommended filename.
    incrementFilename = tmpFilename & " (" & CStr(numToAppend) & ")"

End Function

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

'Returns a boolean as to whether or not we have write access to a given directory.
' (If we do not have write access, the function will return "False".)
Public Function DirectoryHasWriteAccess(ByRef dName As String) As Boolean
    
    'Before checking write access, make sure the directory exists
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
        
        DirectoryHasWriteAccess = True
        Exit Function
        
    Else
        DirectoryHasWriteAccess = True
    End If
    
noWriteAccess:

    DirectoryHasWriteAccess = False
End Function

'Straight from MSDN - generate a "browse for folder" dialog
Public Function BrowseForFolder(ByVal srcHwnd As Long) As String
    
    Dim objShell As Shell
    Dim objFolder As Folder
    Dim returnString As String
        
    Set objShell = New Shell
    Set objFolder = objShell.BrowseForFolder(srcHwnd, g_Language.TranslateMessage("Please select a folder:"), 0)
            
    If (Not objFolder Is Nothing) Then returnString = objFolder.Items.Item.Path Else returnString = ""
    
    Set objFolder = Nothing
    Set objShell = Nothing
    
    BrowseForFolder = returnString
    
End Function

'Open a string as a hyperlink in the user's default browser
Public Sub OpenURL(ByVal targetURL As String)
    ShellExecute FormMain.hWnd, "Open", targetURL, "", 0, SW_SHOWNORMAL
End Sub

'Execute another program (in PhotoDemon's case, a plugin), then wait for it to finish running.
Public Function ShellAndWait(ByVal sPath As String, ByVal winStyle As VbAppWinStyle) As Boolean

    Dim procID As Long
    Dim procHandle As Long

    ' Start the program.
    On Error GoTo ShellError
    procID = Shell(sPath, winStyle)
    On Error GoTo 0

    ' Wait for the program to finish.
    ' Get the process handle.
    procHandle = OpenProcess(SYNCHRONIZE, 0, procID)
    If procHandle <> 0 Then
        WaitForSingleObject procHandle, WAIT_INFINITE
        CloseHandle procHandle
    End If

    ' Reappear.
    ShellAndWait = True
    Exit Function

ShellError:
    ShellAndWait = False
End Function
