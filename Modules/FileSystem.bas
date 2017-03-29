Attribute VB_Name = "FileSystem"
'***************************************************************************
'Miscellaneous Functions Related to File and Folder Interactions
'Copyright 2001-2017 by Tanner Helland
'Created: 6/12/01
'Last updated: 28/May/16
'Last update: update the BrowseForFolder function to use the new IFileDialog interfaces on Vista+.  Note that this
'             introduces several new dependencies for this module.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'API calls for retrieving detailed date time for a given file
Private Const MAX_PATH = 260
Private Const INVALID_HANDLE_VALUE = -1

Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As Currency
    ftLastAccessTime As Currency
    ftLastWriteTime As Currency
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type

Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function FileTimeToLocalFileTime Lib "kernel32" (ByRef lpFileTime As Currency, ByRef lpLocalFileTime As Currency) As Long

' Difference between day zero for VB dates and Win32 dates (or #12-30-1899# - #01-01-1601#)
Private Const rDayZeroBias As Double = 109205#   ' Abs(CDbl(#01-01-1601#))

' 10000000 nanoseconds * 60 seconds * 60 minutes * 24 hours / 10000 comes to 86400000 (the 10000 adjusts for fixed point in Currency)
Private Const rMillisecondPerDay As Double = 10000000# * 60# * 60# * 24# / 10000#

'Min/max date values
Private Const datMin As Date = #1/1/100#
Private Const datMax As Date = #12/31/9999 11:59:59 PM#

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

'PD creates a lot of temp files.  To prevent hard drive thrashing, this module tracks created temp files,
' and deletes them en masse when PD terminates.  (You can also delete temp files prematurely by calling the
' associated function, but because some PD tasks run asynchronously, this behavior is not generally advised.)
Private m_NumOfTempFiles As Long
Private m_ListOfTempFiles() As String
Private Const INIT_TEMP_FILE_CACHE As Long = 16

'If a file exists, this function can be used to intelligently increment the file name (e.g. "filename (n+1).ext")
' Note that the function returns the auto-incremented filename WITHOUT an extension and WITHOUT a prepended folder,
' by design, so that the result can be passed to a common dialog without further parsing.
'
'That said, an initial extension is still required, because this function should only be used if a file name with
' a matching extension exists (e.g. it is perfectly fine to have the same filename with DIFFERENT extensions in the
' target directory).
Public Function IncrementFilename(ByRef dstDirectory As String, ByRef fName As String, ByRef desiredExtension As String) As String
    
    'First, check to see if a file with that name and extension appears in the destination directory.
    ' If it does, just return the filename we were passed.
    Dim cFile As pdFSO
    Set cFile = New pdFSO
    
    If Not cFile.FileExist(dstDirectory & fName & "." & desiredExtension) Then
        IncrementFilename = fName
    Else

        'If we made it to this line of code, a file with that name and extension appears in the destination directory.
        
        'Start by figuring out if the file is already in the format: "filename (#).ext"
        Dim tmpFilename As String
        tmpFilename = Trim$(fName)
        
        Dim numToAppend As Long
        
        'Check the trailing character.  If it is a closing parentheses ")", we need to analyze more
        If (StrComp(Right$(tmpFilename, 1), ")", vbBinaryCompare) = 0) Then
        
            Dim i As Long
            For i = Len(tmpFilename) - 1 To 1 Step -1
                
                'If this character is a number, continue scanning leftward until we find a character that is *not* a number
                If (Not IsNumeric(Mid$(tmpFilename, i, 1))) Then
                    
                    'This character is non-numeric.  See if it's an opening parentheses.
                    If StrComp(Mid$(tmpFilename, i, 1), "(", vbBinaryCompare) = 0 Then
                        
                        'This filename already adheres to the pattern "Filename (###).ext".  To spare us from auto-scanning
                        ' numbers that are likely taken, use this number as our starting point for auto-incrementing.
                        numToAppend = CLng(Mid$(tmpFilename, i + 1, Len(tmpFilename) - i - 1))
                        tmpFilename = Left$(tmpFilename, i - 2)
                        Exit For
                    
                    'If this character is non-numeric and NOT an initial parentheses, this filename is not in the format we want.
                    ' Treat it like any other filename and start by appending " (2)" to it
                    Else
                        numToAppend = 2
                        Exit For
                    End If
                    
                End If
            
            Next i
        
        'If this is not already a copy of the format "Filename (###).ext", start scanning at ### = 2
        Else
            numToAppend = 2
        End If
        
        'Loop through the folder, looking for the first "Filename (###).ext" variant that is not already taken.
        Do While cFile.FileExist(dstDirectory & tmpFilename & " (" & CStr(numToAppend) & ")" & "." & desiredExtension)
            numToAppend = numToAppend + 1
        Loop
            
        'If the loop has terminated, a unique filename has been found.  Make that the recommended filename.
        IncrementFilename = tmpFilename & " (" & CStr(numToAppend) & ")"
        
    End If

End Function

'Generate a "browse for folder" dialog.  Vista+ users get a fancy new interface; XP users get the shitty old interface.
' Many thanks to vbForums user "LaVolpe", who wrote both the interfaces used here.  For implementation details and
' links to the original, unmodified classes, please read the header comments in the referenced classes.
Public Function BrowseForFolder(ByVal srcHwnd As Long, Optional ByVal initFolder As String = vbNullString, Optional ByVal dialogTitleText As String = vbNullString) As String
    
    If (Len(dialogTitleText) = 0) Then dialogTitleText = g_Language.TranslateMessage("Please select a folder")
    
    Dim cSysInfo As pdSystemInfo
    Set cSysInfo = New pdSystemInfo
    
    'Vista+ users get the fancy new "browse for folder" interface
    If cSysInfo.IsOSVistaOrLater Then
    
        Dim cBrowseNew As cFileDialogVista
        Set cBrowseNew = New cFileDialogVista
        
        With cBrowseNew
            .propFlags = FOS__BrowseFoldersDefaults
            If (Len(initFolder) <> 0) Then .propStartupFolder_Set initFolder, ppType_AsString
            
            'For reasons I don't understand, a strange magic number is used to report cancellation; see the
            ' function documentation for additional details.
            If (.DialogShow(srcHwnd, FDLG_BROWSEFOLDERS, dialogTitleText) <> -2147023673) Then
                BrowseForFolder = .IShellItem_GetDisplayName(ObjPtr(.ResultsItem(1)), SIGDN_FILESYSPATH, False)
            Else
                BrowseForFolder = vbNullString
            End If
        End With
        
    'XP users get the crappy old browse interface.
    Else
        
        Dim cBrowse As cUnicodeBrowseFolders
        Set cBrowse = New cUnicodeBrowseFolders
        
        With cBrowse
        
            If (Len(initFolder) <> 0) Then .InitialDirectory = initFolder
            .dialogTitle = dialogTitleText
            .Flags = BIF_RETURNONLYFSDIRS Or BIF_NEWDIALOGSTYLE
                
            If .ShowBrowseForFolder(srcHwnd) Then
                BrowseForFolder = cBrowse.SelectedFolder
            Else
                BrowseForFolder = vbNullString
            End If
            
        End With
        
    End If
    
    'PD folder functions enforce a trailing slash, to simplify subsequent concatenations
    If (Len(BrowseForFolder) <> 0) Then
        Dim cFSO As pdFSO
        Set cFSO = New pdFSO
        BrowseForFolder = cFSO.EnforcePathSlash(BrowseForFolder)
    End If
    
End Function

'Open a string as a hyperlink in the user's default browser
Public Sub OpenURL(ByVal targetURL As String)
    Dim targetAction As String: targetAction = "Open"
    ShellExecute FormMain.hWnd, StrPtr(targetAction), StrPtr(targetURL), 0&, 0&, SW_SHOWNORMAL
End Sub

'Execute another program (in PhotoDemon's case, a plugin), then wait for it to finish running.
Public Function ShellAndWait(ByVal sPath As String, ByVal winStyle As VbAppWinStyle) As Boolean

    Dim procID As Long
    Dim procHandle As Long

    'Start the program
    On Error GoTo ShellError
    procID = Shell(sPath, winStyle)
    On Error GoTo 0

    'Use the external program's handle to initiate a wait process
    procHandle = OpenProcess(SYNCHRONIZE, 0, procID)
    If procHandle <> 0 Then
        WaitForSingleObject procHandle, WAIT_INFINITE
        CloseHandle procHandle
    End If
    
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
    
    For x = Len(sString) To 1 Step -1
        If (Mid$(sString, x, 1) = "/") Or (Mid$(sString, x, 1) = "\") Then
            sString = Left$(sString, x)
            Exit Sub
        End If
    Next x
    
End Sub

'Given a full file path (path + name + extension), return the directory structure
Public Function GetDirectory(ByRef sString As String) As String
    
    Dim x As Long
    
    For x = Len(sString) - 1 To 1 Step -1
        If (Mid$(sString, x, 1) = "/") Or (Mid$(sString, x, 1) = "\") Then
            GetDirectory = Left$(sString, x)
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
Public Function GetFilename(ByVal sString As String) As String

    Dim i As Long
    
    For i = Len(sString) - 1 To 1 Step -1
        If (Mid$(sString, i, 1) = "/") Or (Mid$(sString, i, 1) = "\") Then
            GetFilename = Right$(sString, Len(sString) - i)
            Exit Function
        End If
    Next i
    
End Function

'Return a filename without an extension
Public Function GetFilenameWithoutExtension(ByVal sString As String) As String

    Dim tmpFilename As String

    Dim i As Long
    
    For i = Len(sString) - 1 To 1 Step -1
        If (Mid$(sString, i, 1) = "/") Or (Mid$(sString, i, 1) = "\") Then
            tmpFilename = Right$(sString, Len(sString) - i)
            Exit For
        End If
    Next i
    
    'If we were only passed a filename (without the rest of the path), restore the original entry now
    If Len(tmpFilename) = 0 Then tmpFilename = sString
    
    'Remove the extension, if any
    StripOffExtension tmpFilename
    
    GetFilenameWithoutExtension = tmpFilename
    
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
    strOutput = Left$(strOutput, InStr(1, strOutput, "/", vbBinaryCompare) - 1)
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

'Retrieve the requested date type (creation, access, or last-modified time) of a file.
' Thank you to http://vb.mvps.org/hardcore/html/filedatestimes.htm for this function.
Public Function FileAnyDateTime(ByRef sPath As String, Optional ByRef datCreation As Date = datMin, Optional ByRef datAccess As Date = datMin) As Date
    
    ' Take the easy way if no optional arguments
    If datCreation = datMin And datAccess = datMin Then
        FileAnyDateTime = VBA.FileDateTime(sPath)
        Exit Function
    End If

    Dim fnd As WIN32_FIND_DATA
    Dim hFind As Long
    
    ' Get all three times in UDT
    hFind = FindFirstFile(sPath, fnd)
    If hFind = INVALID_HANDLE_VALUE Then Debug.Print "Requested file " & sPath & " was not found!"
    FindClose hFind
    
    ' Convert them to Visual Basic format
    datCreation = Win32ToVbTime(fnd.ftCreationTime)
    datAccess = Win32ToVbTime(fnd.ftLastAccessTime)
    FileAnyDateTime = Win32ToVbTime(fnd.ftLastWriteTime)
    
End Function

'Sub function for FileAnyDateTime, above.  Once again, thank you to
' http://vb.mvps.org/hardcore/html/filedatestimes.htm for the code.
Private Function Win32ToVbTime(ft As Currency) As Date
    
    Dim ftl As Currency
    
    ' Call API to convert from UTC time to local time
    If FileTimeToLocalFileTime(ft, ftl) Then
        ' Local time is nanoseconds since 01-01-1601
        ' In Currency that comes out as milliseconds
        ' Divide by milliseconds per day to get days since 1601
        ' Subtract days from 1601 to 1899 to get VB Date equivalent
        Win32ToVbTime = CDate((ftl / rMillisecondPerDay) - rDayZeroBias)
    Else
        Debug.Print "FileTimeToLocalFileTime failed!"
    End If
    
End Function

'Request a temporary filename.  The filename will automatically be added to PD's internal cache, and deleted when
' PD exits.  The caller *can* delete the file if they want, but it is not necessary.
Public Function RequestTempFile() As String

    If (m_NumOfTempFiles = 0) Then
        ReDim m_ListOfTempFiles(0 To INIT_TEMP_FILE_CACHE - 1) As String
    Else
        If m_NumOfTempFiles > UBound(m_ListOfTempFiles) Then ReDim Preserve m_ListOfTempFiles(0 To m_NumOfTempFiles * 2 - 1) As String
    End If
        
    Dim cFile As pdSystemInfo
    Set cFile = New pdSystemInfo
    
    Dim tmpFile As String
    tmpFile = cFile.GetUniqueTempFilename()
    
    m_ListOfTempFiles(m_NumOfTempFiles) = tmpFile
    m_NumOfTempFiles = m_NumOfTempFiles + 1
    RequestTempFile = tmpFile

End Function

Public Sub DeleteTempFiles()

    If (m_NumOfTempFiles > 0) Then
    
        Dim cFile As pdFSO
        Set cFile = New pdFSO
        
        Dim i As Long
        For i = 0 To m_NumOfTempFiles - 1
            If Len(m_ListOfTempFiles(i)) <> 0 Then
                If cFile.FileExist(m_ListOfTempFiles(i)) Then cFile.KillFile m_ListOfTempFiles(i)
            End If
        Next i
        
        m_NumOfTempFiles = 0
    
    End If
    
End Sub
