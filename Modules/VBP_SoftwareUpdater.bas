Attribute VB_Name = "Software_Updater"
'***************************************************************************
'Automatic Software Updater (note: at present this doesn't techincally DO the updating (e.g. overwriting program files), it just CHECKS for updates)
'Copyright ©2012-2014 by Tanner Helland
'Created: 19/August/12
'Last updated: 14/August/13
'Last update: rewrote all update code against XML instead of INI.  This was the last INI fix needed, so now PD is 100% free of INI files.  Yay!
'              Also, the software update function now returns custom type UpdateCheck, which is more descriptive than arbitrary ints.
'
'Interface for checking if a new version of PhotoDemon is available for download.  This code is a stripped-down
' version of PhotoDemon's "download image from Internet" code.
'
'The code should be extremely robust against Internet and other miscellaneous errors.  Technically an update check is
' very simple - simply download an XML file from the photodemon.org server, and compare the version numbers in the
' file against the ones supplied by this build.  If the numbers don't match, recommend an update.
'
'Note that this code interfaces with the user preferences file so the user can opt to not check for updates and never
' be notified again. (FYI - this option can be enabled/disabled from the 'Tools' -> 'Options' menu.)
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Because the update form needs access to the update version numbers, they are made publicly available
Public updateMajor As Long, updateMinor As Long, updateBuild As Long

Public Enum UpdateCheck
    UPDATE_ERROR = 0
    UPDATE_NOT_NEEDED = 1
    UPDATE_AVAILABLE = 2
    UPDATE_UNAVAILABLE = 3
End Enum

#If False Then
    Const UPDATE_ERROR = 0
    Const UPDATE_NOT_NEEDED = 1
    Const UPDATE_AVAILABLE = 2
    Const UPDATE_UNAVAILABLE = 3
#End If

'Same goes for the update announcement path
Public updateAnnouncement As String

'Check for a software update; it's assumed the update file has already been downloaded, if available, from its standard
' location at http://photodemon.org/downloads/updates.xml.  If an update file has not been downloaded, this function
' will exit with status code UPDATE_UNAVAILABLE.

'This function will return one of four values:
' UPDATE_ERROR - something went wrong
' UPDATE_NOT_NEEDED - an update file was found, but the current software version is already up-to-date
' UPDATE_AVAILABLE - an update file was found, and an updated PD copy is available
' UPDATE_UNAVAILABLE - no update file was found (this is the most common occurrence, as updates are only checked every 10 days)
Public Function CheckForSoftwareUpdate(Optional ByVal downloadUpdateManually As Boolean = False) As UpdateCheck

    'If the user has requested a forcible update check (as can be done from the Help menu), manually
    ' download a new copy of the update file now.
    If downloadUpdateManually Then
    
        'First things first - set up our target URL
        Dim URL As String
        URL = "http://photodemon.org/downloads/updates.xml"
        'URL = "http://tannerhelland.com/photodemon_files/updates_testing.txt"
           
        'Open an Internet session and assign it a handle
        Dim hInternetSession As Long
        hInternetSession = InternetOpen(App.EXEName, INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)
        
        'If a connection couldn't be established, exit out
        If hInternetSession = 0 Then
            CheckForSoftwareUpdate = UPDATE_ERROR
            Exit Function
        End If
        
        'Using the new Internet session, attempt to find the URL; if found, assign it a handle
        Dim hUrl As Long
        hUrl = InternetOpenUrl(hInternetSession, URL, vbNullString, 0, INTERNET_FLAG_RELOAD, 0)
    
        'If the URL couldn't be found, my server may be down. Close out this connection and exit.
        If hUrl = 0 Then
            If hInternetSession Then InternetCloseHandle hInternetSession
            CheckForSoftwareUpdate = UPDATE_ERROR
            Exit Function
        End If
            
        'We need a temporary file to house the update information; generate it automatically
        Dim tmpFile As String
        tmpFile = g_UserPreferences.getDataPath & "updates.xml"
        
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
       
                'If something went wrong - like the connection dropping mid-download - delete the temp file and terminate the update function
                If Not chunkOK Then
                    
                    'Remove the temporary file
                    If FileExist(tmpFile) Then
                        Close #fileNum
                        Kill tmpFile
                    End If
                    
                    'Close the Internet connection
                    If hUrl Then InternetCloseHandle hUrl
                    If hInternetSession Then InternetCloseHandle hInternetSession
                    
                    CheckForSoftwareUpdate = UPDATE_ERROR
                    Exit Function
                    
                End If
       
                'If the file has downloaded completely, exit this loop
                If numOfBytesRead = 0 Then Exit Do
                
                'If we've made it this far, assume we've received legitimate data. Place that data into the temporary file.
                Put #fileNum, , Left$(Buffer, numOfBytesRead)
                
            'Carry on
            Loop
            
        'Close the temporary file
        Close #fileNum
        
        'With the update file completely downloaded, we can close this URL and Internet session
        If hUrl Then InternetCloseHandle hUrl
        If hInternetSession Then InternetCloseHandle hInternetSession
        
    End If

    
    'Check for the presence of an update file
    Dim updateFile As String
    updateFile = g_UserPreferences.getDataPath & "updates.xml"
    
    If FileExist(updateFile) Then
    
        'Update information file found!  Investigate its contents.
        
        'Note that the update information file is in XML format, so we need an XML parser to read it.
        Dim xmlEngine As pdXML
        Set xmlEngine = New pdXML
        
        'Load the XML file into memory
        xmlEngine.loadXMLFile updateFile
        
        'Check for a few necessary tags, just to make sure this is a valid PhotoDemon update file
        If xmlEngine.isPDDataType("Update report") And xmlEngine.validateLoadedXMLData("updateMajor", "updateMinor", "updateBuild") Then
        
            'Retrieve the version numbers
            updateMajor = xmlEngine.getUniqueTag_Long("updateMajor", -1)
            updateMinor = xmlEngine.getUniqueTag_Long("updateMinor", -1)
            updateBuild = xmlEngine.getUniqueTag_Long("updateBuild", -1)
            
            'If any of the version numbers weren't found, report an error and exit
            If (updateMajor = -1) Or (updateMinor = -1) Or (updateBuild = -1) Then
                If FileExist(updateFile) Then Kill updateFile
                CheckForSoftwareUpdate = UPDATE_ERROR
                Exit Function
            End If
            
            'Finally, check for an update announcement article URL.  This may or may not be blank; it depends on whether I've written an
            ' announcement article yet... :)
            updateAnnouncement = xmlEngine.getUniqueTag_String("updateAnnouncementURL")
            
            'We have what we need from the temporary file, so delete it
            If FileExist(updateFile) Then Kill updateFile
                
            'If we made it all the way here, we can assume the update check was successful.  The last thing we need to do is compare
            ' the updated software version numbers with the current software version numbers.  If THAT yields results, we can finally
            ' return "UPDATE_NEEDED" for this function
            If (updateMajor > App.Major) Or ((updateMinor > App.Minor) And (updateMajor = App.Major)) Or ((updateBuild > App.Revision) And (updateMinor = App.Minor) And (updateMajor = App.Major)) Then
                CheckForSoftwareUpdate = UPDATE_AVAILABLE
            
            '...otherwise, we went to all that work for nothing.  Oh well.  An update check occurred, but this version is up-to-date.
            Else
                CheckForSoftwareUpdate = UPDATE_NOT_NEEDED
            End If
            
        Else
            CheckForSoftwareUpdate = UPDATE_ERROR
        End If
    
    'No update information found.  Return the proper code and exit.
    Else
        CheckForSoftwareUpdate = UPDATE_UNAVAILABLE
    End If
    
    
End Function

'Determine if the program should check online for update information.  This will return true IFF the following
' criteria are met:
' 1) User preferences allow us to check for updates (e.g. the user has not forcibly disabled such checks)
' 2) At least 10 days have passed since the last update check...
' 3) ...or 10 days haven't passed, but we have never checked for updates before, and this is NOT the first time the user
'    is running the program
Public Function isItTimeForAnUpdate() As Boolean

    'Locale settings can sometimes screw with the DateDiff function in unpredictable ways.  If something goes
    ' wrong, disable update checks for this session
    On Error GoTo noUpdates

    Dim allowedToUpdate As Boolean
    allowedToUpdate = g_UserPreferences.GetPref_Boolean("Updates", "Check For Updates", True)
        
    'If updates ARE allowed, see when we last checked.  To be polite, only check once every 10 days.
    If allowedToUpdate Then
    
        Dim lastCheckDate As String
        lastCheckDate = g_UserPreferences.GetPref_String("Updates", "Last Update Check", "")
        
        'If the last update check date was not found, request an update check now
        If Len(lastCheckDate) = 0 Then
        
            allowedToUpdate = True
        
        'If a last update check date was found, check to see how much time has elapsed since that check
        Else
        
            Dim currentDate As Date
            currentDate = Format$(Now, "Medium Date")
            
            'If 10 days have elapsed, allow an update check
            If CLng(DateDiff("d", CDate(lastCheckDate), currentDate)) >= 10 Then
                allowedToUpdate = True
            
            'If 10 days haven't passed, prevent an update
            Else
                Message "Update check postponed (a check has been performed in the last 10 days)"
                allowedToUpdate = False
            End If
                    
        End If
    
    End If
    
    isItTimeForAnUpdate = allowedToUpdate
    
    Exit Function
    
noUpdates:

    isItTimeForAnUpdate = False
    
End Function
