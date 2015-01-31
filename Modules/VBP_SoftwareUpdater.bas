Attribute VB_Name = "Software_Updater"
'***************************************************************************
'Automatic Software Updater (note: at present this doesn't techincally DO the updating (e.g. overwriting program files), it just CHECKS for updates)
'Copyright 2012-2015 by Tanner Helland
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

    'Locale settings can sometimes screw with the DateDiff function in unpredictable ways.  (For example, if a
    ' user is using PD on a pen drive, and they move between PCs with wildly different date representations)
    ' If something goes wrong at any point in this function, we'll simply disable update checks until next run.
    On Error GoTo noUpdates

    Dim allowedToUpdate As Boolean
    allowedToUpdate = False
    
    'Previous to v6.6, PD's update preference was a binary yes/no thing.  To make sure users who previously disabled
    ' updates are respected, if the new preference doesn't exist yet, we'll use the old preference value instead.
    Dim updateFrequency As PD_UPDATE_FREQUENCY
    updateFrequency = PDUF_EACH_SESSION
    If g_UserPreferences.doesValueExist("Updates", "Check For Updates") Then
        
        If Not g_UserPreferences.GetPref_Boolean("Updates", "Check For Updates", True) Then
            
            'Write a matching preference in the new format.
            g_UserPreferences.SetPref_Long "Updates", "Update Frequency", PDUF_NEVER
            
            'Overwrite the old preference, so it doesn't trigger this check again
            g_UserPreferences.SetPref_Boolean "Updates", "Check For Updates", True
            
        End If
        
    End If
    
    'In v6.6, PD's update strategy was modified to allow the user to specify an update frequency (rather than
    ' a binary yes/no preference).  Retrieve the allowed frequency now.
    If updateFrequency <> PDUF_NEVER Then
        updateFrequency = g_UserPreferences.GetPref_Long("Updates", "Update Frequency", PDUF_EACH_SESSION)
    End If
        
    'If updates ARE allowed, see when we last checked for an update.  If enough time has elapsed, check again.
    If updateFrequency <> PDUF_NEVER Then
    
        Dim lastCheckDate As String
        lastCheckDate = g_UserPreferences.GetPref_String("Updates", "Last Update Check", "")
        
        'If a "last update check date" was not found, request an immediate update check.
        If Len(lastCheckDate) = 0 Then
        
            allowedToUpdate = True
        
        'If a last update check date was found, check to see how much time has elapsed since that check.
        Else
        
            'Start by figuring out how many days need to have passed before we're allowed to check for updates
            ' again.  (This varies according to user preference.)
            Dim numAllowableDays As Long
            
            Select Case updateFrequency
            
                Case PDUF_EACH_SESSION
                    numAllowableDays = 0
                
                Case PDUF_WEEKLY
                    numAllowableDays = 7
                    
                Case PDUF_MONTHLY
                    numAllowableDays = 30
            
            End Select
            
            Dim currentDate As Date
            currentDate = Format$(Now, "Medium Date")
            
            'If the allowable date threshold has passed, allow the updater to perform a new check
            If CLng(DateDiff("d", CDate(lastCheckDate), currentDate)) >= numAllowableDays Then
                allowedToUpdate = True
            
            'If 10 days haven't passed, prevent an update
            Else
                Message "Update check postponed (a check has been performed recently)"
                allowedToUpdate = False
            End If
                    
        End If
    
    'If the user has specified "never" as their desired update frequency, we'll always return FALSE.
    Else
        allowedToUpdate = False
    End If
    
    isItTimeForAnUpdate = allowedToUpdate
    
    Exit Function

'In the event of an error, simply disallow updates for this session.
noUpdates:

    isItTimeForAnUpdate = False
    
End Function

'Given the XML string of a download language version XML report from photodemon.org, initiate downloads of any languages that need to be updated.
Public Sub processLanguageUpdateFile(ByRef srcXML As String)
    
    'This boolean array will be track which (if any) language files are in need of an update.  The matching "numOfUpdates"
    ' value will be > 0 if any files need updating.
    Dim numOfUpdates As Long
    numOfUpdates = 0
    
    Dim updateFlags() As Boolean
    
    'We will be testing a number of different languages to see if they qualify for an update.  This temporary object will
    ' be passed to the public translation class as necessary, to retrieve a copy of a given language file's data.
    Dim tmpLanguage As pdLanguageFile
        
    'A pdXML object handles XML parsing for us.
    Dim xmlEngine As pdXML
    Set xmlEngine = New pdXML
    
    Dim langVersion As String, langID As String, langRevision As Long
    
    'Validate the XML
    If xmlEngine.loadXMLFromString(srcXML) Then
    
        'Check for a few necessary tags, just to make sure this is actually a PhotoDemon language file
        If xmlEngine.isPDDataType("Language versions") Then
        
            'We're now going to enumerate all language tags in the file.  If one needs to be updated, a couple extra
            ' steps need to be taken.
            Dim langList() As String
            If xmlEngine.findAllAttributeValues(langList, "language", "updateID") Then
                
                'langList() now contains a list of all the unique language listings in the update file.
                ' We want to search this list for entries with an identical major/minor version to this PD build.
                
                'Start by retrieving the current PD executable version.
                Dim currentPDVersion As String
                currentPDVersion = getPhotoDemonVersionMajorMinorOnly
                
                'This step is simply going to flag language files in need of an update.  This array will be used to track
                ' such language; a separate step will initiate the actual update downloads.
                ReDim updateFlags(0 To UBound(langList)) As Boolean
                
                'Iterate the language update list, looking for version matches
                Dim i As Long
                For i = 0 To UBound(langList)
                
                    'Retrieve the major/minor version of this language file.  (String format is fine, as we're just
                    ' checking equality.)
                    langVersion = xmlEngine.getUniqueTag_String("version", , , "language", "updateID", langList(i))
                    langVersion = retrieveVersionMajorMinorAsString(langVersion)
                    
                    'Retrieve the language's revision as well.  This is explicitly retrieved as a LONG, because we need to perform
                    ' a >= check between it and the current language file revision.
                    langRevision = retrieveVersionRevisionAsLong(xmlEngine.getUniqueTag_String("revision", , , "language", "updateID", langList(i)))
                    
                    'If the version matches this .exe version, this language file is a candidate for updating.
                    If StrComp(currentPDVersion, langVersion, vbBinaryCompare) = 0 Then
                    
                        'Next, we need to compare the versions of the update language file and the installed language file.
                        
                        'Retrieve the language ID, which is a unique identifier.
                        langID = xmlEngine.getUniqueTag_String("id", , , "language", "updateID", langList(i))
                        
                        'Use a helper function to retrieve the language header for the currently installed copy of this language.
                        If g_Language.getPDLanguageFileObject(tmpLanguage, langID) Then
                        
                            'A matching language file was found.  Compare version numbers.
                            If StrComp(langVersion, retrieveVersionMajorMinorAsString(tmpLanguage.langVersion), vbBinaryCompare) = 0 Then
                            
                                'The major/minor version of this language file matches the current language.  Compare revisions.
                                If langRevision > retrieveVersionRevisionAsLong(tmpLanguage.langVersion) Then
                                
                                    'Holy shit, this language actually needs to be updated!  :P  Mark the corresponding location
                                    ' in the update array, and increment the update counter.
                                    updateFlags(i) = True
                                    numOfUpdates = numOfUpdates + 1
                                    
                                    Debug.Print "Language ID " & langID & " will be updated to revision " & langRevision & "."
                                
                                'The current file is up-to-date.
                                Else
                                
                                    Debug.Print "Language ID " & langID & " is already up-to-date."
                                
                                End If
                            
                            End If
                            
                        'This language ID was not found.  This could mean one of two things:
                        ' 1) This language file is a new one (e.g. it was not included in the original release of this PD version)
                        ' 2) The user forcibly deleted this language file at some point in the past.
                        '
                        'To avoid undoing (2), we must also ignore (1).  Do nothing if this language file doesn't exist in the
                        ' current languages folder.
                        Else
                            Debug.Print "Language ID " & langID & " does not exist on this PC.  No update will be performed."
                        End If
                    
                    End If
                    
                Next i
            
            End If
        
            'At this point, updateFlags() will contain TRUE for any language files that need to be updated, and numOfUpdates should
            ' be greater than 0 if updates are required.
            If numOfUpdates > 0 Then
                
                'TODO!
                Debug.Print numOfUpdates & " updated language files will now be downloaded."
                
            Else
                Debug.Print "All language files are up-to-date.  No new files will be downloaded."
            End If
        
        Else
            Debug.Print "WARNING! Language update XML did not pass basic validation.  Abandoning update process."
        End If
    
    Else
        Debug.Print "WARNING! Language update XML did not load successfully - check for an encoding error, maybe...?"
    End If
    
End Sub
