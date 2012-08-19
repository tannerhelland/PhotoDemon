Attribute VB_Name = "SoftwareUpdater"
'***************************************************************************
'Automatic Software Updater (note: at present this doesn't techincally DO the updating (e.g. overwriting program files), it just CHECKS for updates)
'Copyright �2000-2012 by Tanner Helland
'Created: 19/August/12
'Last updated: 19/August/12
'Last update: initial build
'
'Interface for checking if a new version of PhotoDemon is available for download.  This code is a stripped-down
' version of PhotoDemon's "download image from Internet" code.
'
'A number of features have been added to the original version of this code, particularly the many checks I've added
' to protect against Internet and download errors.  Technically this code is fairly simply - it simply downloads a text
' file from the tannerhelland.com server, and compares the version numbers it provides against the ones supplied by this
' build  If the numbers don't match, it spawns the related form and recommends a download.
'
'Additionally, this code interfaces with the .INI file so the user can opt to not check for updates and never be
' notified again. (FYI - this option can be enabled/disabled from the 'Edit' -> 'Program Preferences' menu.)
'
'***************************************************************************

Option Explicit

'Because the update form needs access to the update version numbers, they are made publicly available
Public updateMajor As Long, updateMinor As Long

'Same goes for the update announcement path
Public updateAnnouncement As String

'Check for a software update; the update info will be contained in a text file at http://tannerhelland.com/photodemon_files/updates.txt
Public Function CheckForSoftwareUpdate() As Boolean

    'First things first - set up our target URL
    Dim URL As String
    URL = "http://tannerhelland.com/photodemon_files/updates.txt"
       
    'Open an Internet session and assign it a handle
    Dim hInternetSession As Long
    hInternetSession = InternetOpen(App.EXEName, INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)
    
    'If a connection couldn't be established, exit out
    If hInternetSession = 0 Then
        CheckForSoftwareUpdate = False
        Exit Function
    End If
    
    'Using the new Internet session, attempt to find the URL; if found, assign it a handle
    Dim hUrl As Long
    hUrl = InternetOpenUrl(hInternetSession, URL, vbNullString, 0, INTERNET_FLAG_EXISTING_CONNECT, 0)

    'If the URL couldn't be found, the server may be down.  Close out this connection and exit out
    If hUrl = 0 Then
        If hInternetSession Then InternetCloseHandle hInternetSession
        CheckForSoftwareUpdate = False
        Exit Function
    End If
        
    'We need a temporary file to house the update information; generate it automatically
    Dim tmpFile As String
    tmpFile = TempPath & "updates.txt"
    
    'Open the temporary file and begin downloading the update information to it
    Dim fileNum As Integer
    fileNum = FreeFile
    Open tmpFile For Binary As fileNum
    
        'Prepare a receiving buffer (this will be used to hold chunks of the file)
        Dim Buffer As String
        Buffer = Space(4096)
   
        'We will need to verify each chunk as its downloaded
        Dim chunkOK As Boolean
   
        'This will track the size of each chunk
        Dim numOfBytesRead As Long
   
        'This will track of how many bytes we've downloaded so far
        Dim totalBytesRead As Long
        totalBytesRead = 0
   
        Do
   
            'Read the next chunk of the image
            chunkOK = InternetReadFile(hUrl, Buffer, Len(Buffer), numOfBytesRead)
   
            'If something went wrong - like the connection dropping mid-download - terminate the function
            If chunkOK = False Then
                
                'Remove the temporary file
                If FileExist(tmpFile) Then
                    Close #fileNum
                    Kill tmpFile
                End If
                
                'Close out the Internet connection
                If hUrl Then InternetCloseHandle hUrl
                If hInternetSession Then InternetCloseHandle hInternetSession
                
                CheckForSoftwareUpdate = False
                Exit Function
            End If
   
            'If the file has downloaded completely, exit this loop
            If numOfBytesRead = 0 Then
                Exit Do
            End If
   
            'If we've made it this far, assume we've received legitimate data.  Place that data into the temporary file.
            Put #fileNum, , Left$(Buffer, numOfBytesRead)
            
        'Carry on
        Loop
        
    'Close the temporary file
    Close #fileNum
    
    'Close this URL and Internet session
    If hUrl Then InternetCloseHandle hUrl
    If hInternetSession Then InternetCloseHandle hInternetSession
    
    'Fix the line endings (which will be in UNIX format)
    fixLineEndings tmpFile
    
    'If we've made it this far, it's probably safe to assume that download worked.  Attempt to load the update information.
    Dim iniCategory As String
    iniCategory = "PhotoDemon Update Information"
    
    Dim tmpIniRead As String
    
    'Attempt to retrieve the major version number
    tmpIniRead = GetFromArbitraryIni(tmpFile, iniCategory, "Major")
    
    'Verify the major version number
    If tmpIniRead <> "" Then
        updateMajor = CLng(tmpIniRead)
    
    'If it returns a blank string, something went wrong.  Exit the function
    Else
        If FileExist(tmpFile) Then Kill tmpFile
        CheckForSoftwareUpdate = False
        Exit Function
    End If
    
    'Attempt to retrieve the minor version number
    tmpIniRead = GetFromArbitraryIni(tmpFile, iniCategory, "Minor")
    
    'Verify the minor version number
    If tmpIniRead <> "" Then
        updateMinor = CLng(tmpIniRead)
    
    'If it returns a blank string, something went wrong.  Exit the function
    Else
        If FileExist(tmpFile) Then Kill tmpFile
        CheckForSoftwareUpdate = False
        Exit Function
    End If
    
    'Finally, attempt to grab the update announcement URL.  This may or may not be blank; it depends on whether I've
    ' written an announcement yet, heh.
    tmpIniRead = GetFromArbitraryIni(tmpFile, iniCategory, "AnnouncementURL")
    updateAnnouncement = tmpIniRead
    
    'We have what we need from the temporary file, so delete it
    If FileExist(tmpFile) Then Kill tmpFile
        
    'If we made it all the way here, we can assume the update check was successful.  The last thing we need to do is compare
    ' the updated software version numbers with the current software version numbers.  If THAT yields results, we can finally
    ' return "TRUE" for this function
    If (updateMajor > App.Major) Or (updateMinor > App.Minor) Then
        CheckForSoftwareUpdate = True
    
    '...otherwise, we went to all that work for nothing.  Oh well.  An update check occurred, but this version is up-to-date.
    Else
        CheckForSoftwareUpdate = False
    End If
    
End Function

'Downloaded files will have UNIX file endings, so they need to be converted to Windows ones (which VB requires)
Private Sub fixLineEndings(fileToFix As String)

    Dim fileContents As String

    'Open the file and pull out the text (which VB thinks is all on one line)
    Dim fileNum As Integer
    fileNum = FreeFile
    
    Open fileToFix For Input As #fileNum
        Line Input #fileNum, fileContents
    Close #fileNum
    
    'Kill the original file
    If FileExist(fileToFix) Then Kill fileToFix
    
    'Open the file again, but this time we're going to write out the proper text
    fileContents = Replace(fileContents, Chr(10), vbCrLf)
    Open fileToFix For Binary As #fileNum
        Put #fileNum, , fileContents
    Close #fileNum

End Sub
