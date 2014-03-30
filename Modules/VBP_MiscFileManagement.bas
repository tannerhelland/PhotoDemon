Attribute VB_Name = "File_And_Path_Handling"
'***************************************************************************
'Miscellaneous Functions Related to File and Folder Interactions
'Copyright ©2001-2014 by Tanner Helland
'Created: 6/12/01
'Last updated: 30/March/14
'Last update: add $ qualifier to various string functions (e.g. Mid$())
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Used to quickly check if a file (or folder) exists.  Thanks to Bonnie West's "Optimum FileExists Function"
' for this technique: http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=74264&lngWId=1
Private Const ERROR_SHARING_VIOLATION As Long = 32
Private Declare Function GetFileAttributesW Lib "kernel32" (ByVal lpFileName As Long) As Long

'Used to shell an external program, then wait until it completes
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessID As Long) As Long
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
    
        Dim i As Long
        For i = Len(tmpFilename) - 2 To 1 Step -1
            
            ' If it isn't a number, see if it's an initial parentheses: "("
            If Not (IsNumeric(Mid(tmpFilename, i, 1))) Then
                
                'If it is a parentheses, then this file already has a "( #)" appended to it.  Figure out what the
                ' number inside the parentheses is, and strip that entire block from the filename.
                If Mid(tmpFilename, i, 1) = "(" Then
                
                    numToAppend = CLng(Val(Mid(tmpFilename, i + 1, Len(tmpFilename) - i - 1)))
                    tmpFilename = Left(tmpFilename, i - 2)
                    Exit For
                
                'If this character is non-numeric and NOT an initial parentheses, this filename is not in the format we want.
                ' Treat it like any other filename and start by appending " (2)" to it
                Else
                    numToAppend = 2
                    Exit For
                End If
                
            End If
        
        'If this character IS a number, keep scanning.
        Next i
    
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

'Make sure the right backslash of a path is existant
Public Function FixPath(ByVal tempString As String) As String
    If Right$(tempString, 1) <> "\" Then
        FixPath = tempString & "\"
    Else
        FixPath = tempString
    End If
End Function

'Given a full file path (path + name + extension), remove everything but the directory structure
Public Sub StripDirectory(ByRef sString As String)
    
    Dim x As Long
    
    For x = Len(sString) - 1 To 1 Step -1
        If (Mid$(sString, x, 1) = "/") Or (Mid$(sString, x, 1) = "\") Then
            sString = Left$(sString, x)
            Exit Sub
        End If
    Next x
    
End Sub

'Given a full file path (path + name + extension), return the directory structure
Public Function getDirectory(ByRef sString As String) As String
    
    Dim x As Long
    
    For x = Len(sString) - 1 To 1 Step -1
        If (Mid$(sString, x, 1) = "/") Or (Mid$(sString, x, 1) = "\") Then
            getDirectory = Left(sString, x)
            Exit Function
        End If
    Next x
    
End Function

'Pull the filename ONLY (no directory) off a path
Public Sub StripFilename(ByRef sString As String)
    
    Dim x As Long
    
    For x = Len(sString) - 1 To 1 Step -1
        If (Mid$(sString, x, 1) = "/") Or (Mid$(sString, x, 1) = "\") Then
            sString = Right(sString, Len(sString) - x)
            Exit Sub
        End If
    Next x
    
End Sub

'Return the filename chunk of a path
Public Function getFilename(ByVal sString As String) As String

    Dim i As Long
    
    For i = Len(sString) - 1 To 1 Step -1
        If (Mid$(sString, i, 1) = "/") Or (Mid$(sString, i, 1) = "\") Then
            getFilename = Right$(sString, Len(sString) - i)
            Exit Function
        End If
    Next i
    
End Function

'Return a filename without an extension
Public Function getFilenameWithoutExtension(ByVal sString As String) As String

    Dim tmpFilename As String

    Dim i As Long
    
    For i = Len(sString) - 1 To 1 Step -1
        If (Mid$(sString, i, 1) = "/") Or (Mid$(sString, i, 1) = "\") Then
            tmpFilename = Right$(sString, Len(sString) - i)
            Exit For
        End If
    Next i
    
    StripOffExtension tmpFilename
    
    getFilenameWithoutExtension = tmpFilename
    
End Function

'Pull the filename & directory out WITHOUT any extension (but with the ".")
Public Sub StripOffExtension(ByRef sString As String)

    Dim x As Long

    For x = Len(sString) - 1 To 1 Step -1
        If (Mid$(sString, x, 1) = ".") Then
            sString = Left$(sString, x - 1)
            Exit Sub
        End If
    Next x
    
End Sub

'Function to return the extension from a filename
Public Function GetExtension(sFile As String) As String
    
    Dim i As Long
    For i = Len(sFile) To 1 Step -1
    
        'If we find a path before we find an extension, return a blank string
        If (Mid(sFile, i, 1) = "\") Or (Mid(sFile, i, 1) = "/") Then
            GetExtension = ""
            Exit Function
        End If
        
        If Mid(sFile, i, 1) = "." Then
            GetExtension = Right$(sFile, Len(sFile) - i)
            Exit Function
        End If
    Next i
    
    'If we reach this point, no extension was found
    GetExtension = ""
            
End Function

'Take a string and replace any invalid characters with "_"
Public Sub makeValidWindowsFilename(ByRef FileName As String)

    Dim strInvalidChars As String
    strInvalidChars = "/*?""<>|"
    
    Dim invLoc As Long
    
    Dim x As Long
    For x = 1 To Len(strInvalidChars)
        invLoc = InStr(FileName, Mid$(strInvalidChars, x, 1))
        If invLoc <> 0 Then
            FileName = Left(FileName, invLoc - 1) & "_" & Right(FileName, Len(FileName) - invLoc)
        End If
    Next x

End Sub

'This lovely function comes from "penagate"; it was downloaded from http://www.vbforums.com/showthread.php?t=342995 on 08 June '12
Public Function GetDomainName(ByVal Address As String) As String
        
    Dim strOutput As String, strTemp As String
    Dim lngLoopCount As Long
    Dim lngBCount As Long, lngCharCount As Long
    
    strOutput$ = Replace(Address, "\", "/")
        
    lngCharCount = Len(strOutput)
    
    If (InStrB(1, strOutput, "/")) Then
        
        Do Until ((strTemp = "/") Or (lngLoopCount = lngCharCount))
            lngLoopCount = lngLoopCount + 1
            strTemp = Mid$(strOutput, lngBCount + 1, 1)
            lngBCount = lngBCount + 1
        Loop
        
    End If
        
    strOutput = Right$(strOutput, Len(strOutput) - lngBCount)
    lngBCount = 0
    strTemp = "/"
    
    If (InStrB(1, strOutput, "/")) Then
        
        Do Until strTemp <> "/"
            strTemp = Mid$(strOutput, lngBCount + 1, 1)
            If strTemp = "/" Then lngBCount = lngBCount + 1
        Loop
    
    End If
        
    strOutput = Right$(strOutput, Len(strOutput) - lngBCount)
    strOutput = Left$(strOutput, InStr(1, strOutput, "/", vbTextCompare) - 1)
    GetDomainName = strOutput

End Function

'When passing file and path strings among API calls, they often have to be pre-initialized to some arbitrary buffer length
' (typically MAX_PATH).  When finished, the string needs to be resized to remove any null chars.  Use this function.
Public Function TrimNull(ByVal origString As String) As String

    Dim nullPosition As Long
   
   'double check that there is a chr$(0) in the string
    nullPosition = InStr(origString, Chr$(0))
    If nullPosition Then
       TrimNull = Left$(origString, nullPosition - 1)
    Else
       TrimNull = origString
    End If
  
End Function
