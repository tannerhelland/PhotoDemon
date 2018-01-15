Attribute VB_Name = "Updates"
'***************************************************************************
'Automatic Software Updater
'Copyright 2012-2018 by Tanner Helland
'Created: 19/August/12
'Last updated: 13/December/17
'Last update: clean up code, improve debug reporting, switch to https for patch downloads
'
'This module includes various support functions for determining if a new version of PhotoDemon is available for download.
'
'IMPORTANT NOTE: at present this doesn't technically DO the updating (e.g. overwriting program files), it just CHECKS for updates.
'
'As of March 2015, this module has been completely overhauled to support live-patching of PhotoDemon and its various support files
' (plugins, languages, etc).  Various bits of update code have been moved into the new update support app in the /Support folder.
' The use of a separate patching app greatly simplified things like updating in-use binary files.
'
'Note that this code interfaces with the user preferences file so the user can opt to not check for updates and never
' be notified again. (FYI - this option can be enabled/disabled from the 'Tools' -> 'Options' menu.)
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

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

'When patching PD itself, we make a backup copy of the update XML contents.  This file provides a second failsafe checksum reference, which is
' important when patching binary EXE and DLL files.
Private m_PDPatchXML As String

'When initially parsing the update XML file (above), if an update is found, the parse routine will note which track was used for the update,
' and where that track's data starts and ends inside the XML file.
Private m_SelectedTrack As Long, m_TrackStartPosition As Long, m_TrackEndPosition As Long

'If an update package is downloaded successfully, it will be forwarded to this module.  At program shutdown time, the package will be applied.
Private m_UpdateFilePath As String

'If an update is available, that update's release announcement will be stored in this persistent string.  UI elements can retrieve it as necessary.
Private m_UpdateReleaseAnnouncementURL As String

'If an update is available, that update's track will be stored here.  UI elements can retrieve it as necessary.
Private m_UpdateTrack As PD_UPDATE_TRACK

'Outside functions can also request the update version
Private m_UpdateVersion As String

'Beta releases use custom labeling, independent of the actual version (e.g. "PD 6.6 beta 3"), so we also retrieve and store this value as necessary.
Private m_BetaNumber As String


'Determine if the program should check online for update information.  This will return true IFF the following
' criteria are met:
' 1) User preferences allow us to check for updates (e.g. the user has not forcibly disabled update checks)
' 2) At least (n) days have passed since the last update check, where (n) matches the current user preference...
' 3) ...or (n) days haven't passed, but we have never checked for updates before, and this is NOT the first time the user
'    is running the program.  (This case covers 3rd-party download sites that don't maintain up-to-date download links)
Public Function IsItTimeForAnUpdate() As Boolean

    'Locale settings can sometimes screw with the DateDiff function in unpredictable ways.  (For example, if a
    ' user is using PD on a pen drive, and they move between PCs with wildly different date representations)
    ' If something goes wrong at any point in this function, we'll simply disable update checks until next run.
    On Error GoTo DontDoUpdates

    Dim allowedToUpdate As Boolean
    allowedToUpdate = False
    
    'Previous to v6.6, PD's update preference was a binary yes/no thing.  To make sure users who previously disabled
    ' updates are respected, if the new preference doesn't exist yet, we'll use the old preference value instead.
    Dim updateFrequency As PD_UPDATE_FREQUENCY
    updateFrequency = PDUF_EACH_SESSION
    If g_UserPreferences.DoesValueExist("Updates", "CheckForUpdates") Then
        
        'Write a matching preference in the new format, and overwrite the old preference (so it doesn't trigger this
        ' check again)
        If Not g_UserPreferences.GetPref_Boolean("Updates", "CheckForUpdates", True) Then
            g_UserPreferences.SetPref_Long "Updates", "Update Frequency", PDUF_NEVER
            g_UserPreferences.SetPref_Boolean "Updates", "CheckForUpdates", True
        End If
        
    End If
    
    'In v6.6, PD's update strategy was modified to allow the user to specify an update frequency (rather than
    ' a binary yes/no preference).  Retrieve the allowed frequency now.
    If (updateFrequency <> PDUF_NEVER) Then updateFrequency = g_UserPreferences.GetPref_Long("Updates", "Update Frequency", PDUF_EACH_SESSION)
    
    'If updates ARE allowed, see when we last checked for an update.  If enough time has elapsed, check again.
    If (updateFrequency <> PDUF_NEVER) Then
    
        Dim lastCheckDate As String
        lastCheckDate = g_UserPreferences.GetPref_String("Updates", "Last Update Check")
        
        'If a "last update check date" was not found, request an immediate update check.
        If (Len(lastCheckDate) = 0) Then
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
            allowedToUpdate = (CLng(DateDiff("d", CDate(lastCheckDate), currentDate)) >= numAllowableDays)
            If (Not allowedToUpdate) Then Message "Update check postponed (a check has been performed recently)"
            
        End If
    
    'If the user has specified "never" as their desired update frequency, we'll always return FALSE.
    Else
        allowedToUpdate = False
    End If
    
    IsItTimeForAnUpdate = allowedToUpdate
    
    Exit Function

'In the event of an error, simply disallow updates for this session.
DontDoUpdates:
    IsItTimeForAnUpdate = False
    
End Function

'tl;dr; given an XML report from photodemon.org, initiate a program update package download, as necessary.
' Long version: this function checks to see if PhotoDemon.exe is out of date against the current update track
' (stable, beta, or nightly, per the current user preference).  If the .exe *is* out of date, a full update
' package will be downloaded.
'
'Returns: TRUE if an update is available *and* its download was initiated successfully; FALSE otherwise
Public Function ProcessProgramUpdateFile(ByRef srcXML As String) As Boolean
    
    'In most cases, an update will *not* be available.
    ProcessProgramUpdateFile = False
    
    'A pdXML object handles XML parsing for us.
    Dim xmlEngine As pdXML
    Set xmlEngine = New pdXML
    
    'Validate the XML we were passed
    If xmlEngine.LoadXMLFromString(srcXML) Then
    
        'As an additional precaution, ensure that some PD-specific update tags exist in the XML
        If xmlEngine.IsPDDataType("Program version") Then
            
            'Next, figure out which update track we need to check.  The user can change this at any time, so it may not
            ' necessarily correlate to this .exe's build type.  (For example, maybe this is a stable PD build, but the user
            ' has decided to switch update checks to include beta and nightly builds - that's okay!)
            Dim curUpdateTrack As PD_UPDATE_TRACK
            curUpdateTrack = g_UserPreferences.GetPref_Long("Updates", "Update Track", PDUT_BETA)
            
            'From the update track, we need to generate a string that identifies the correct chunk of the XML file.
            ' Some update tracks can update to more than one type of build (for example, nightly builds can update
            ' to stable builds, if the stable version is newer, but stable builds can only update to nightly builds
            ' if the user's preferences explicit allow), so we may need to search multiple XML regions to find the
            ' most relevant update target for this .exe version.
            Dim updateTagIDs() As String, numUpdateTagIDs As Long
            ReDim updateTagIDs(0 To 2) As String
            updateTagIDs(0) = "stable"
            updateTagIDs(1) = "beta"
            updateTagIDs(2) = "nightly"
            
            Select Case curUpdateTrack
            
                Case PDUT_STABLE
                    numUpdateTagIDs = 1
                    
                Case PDUT_BETA
                    numUpdateTagIDs = 2
                
                Case PDUT_NIGHTLY
                    numUpdateTagIDs = 3
                
            End Select
            
            ReDim Preserve updateTagIDs(0 To numUpdateTagIDs - 1) As String
            
            'If we find an update track that provides a valid update target, this value will point at that track
            ' (0, 1, or 2, for stable, beta, or nightly, respectively).  If no update is found, it will remain at -1.
            Dim trackWithValidUpdate As Long
            trackWithValidUpdate = -1
            
            'We start with the current PD version as a baseline.  If newer update targets are found, this string will
            ' be updated with newer versions instead.
            Dim curVersionMatch As String
            curVersionMatch = GetPhotoDemonVersionCanonical()
            
            'If you want to perform testing against random version numbers, feel free to plug-in your own test version
            ' number here, e.g...
            'curVersionMatch = "6.4.0"
            
            'Next, search the update file for PhotoDemon.exe versions.  Each valid updateTagID region (as calculated above)
            ' will be searched.  If any return a hit, we will initiate the download of that update package.
            Dim i As Long
            For i = 0 To numUpdateTagIDs - 1
            
                'Find the bounding character markers for the relevant XML region (e.g. the one that corresponds to this update track)
                Dim tagAreaStart As Long, tagAreaEnd As Long
                If xmlEngine.GetTagCharacterRange(tagAreaStart, tagAreaEnd, "update", "track", updateTagIDs(i)) Then
                    
                    'Find the position of the reported PhotoDemon.exe version in this track
                    Dim pdTagPosition As Long
                    pdTagPosition = xmlEngine.GetLocationOfTagPlusAttribute("version", "component", "PhotoDemon.exe", tagAreaStart)
                    
                    'Make sure the tag position is within the valid range.  (This should always be TRUE, but it doesn't hurt to check.)
                    If (pdTagPosition >= tagAreaStart) And (pdTagPosition <= tagAreaEnd) Then
                    
                        'This is the version tag we want!  Retrieve its value.
                        Dim newPDVersionString As String
                        newPDVersionString = xmlEngine.GetTagValueAtPreciseLocation(pdTagPosition)
                        InternalDebugMsg "Update track " & i & " reports version " & newPDVersionString & " (our version: " & GetPhotoDemonVersionCanonical() & ")", "ProcessProgramUpdateFile"
                        
                        'If this value is newer than our current update target, mark it and proceed.  Note that this approach gives
                        ' us the highest possible update target from all available/enabled update tracks.
                        If IsNewVersionHigher(curVersionMatch, newPDVersionString) Then
                            
                            trackWithValidUpdate = i
                            
                            'Set some matching module-level values, which we'll need when it's time to actually patch the files
                            ' in question.
                            m_SelectedTrack = trackWithValidUpdate
                            m_TrackStartPosition = tagAreaStart
                            m_TrackEndPosition = tagAreaEnd
                            m_UpdateVersion = newPDVersionString
                            m_UpdateReleaseAnnouncementURL = xmlEngine.GetUniqueTag_String("raurl-" & updateTagIDs(i))
                            
                        End If
                        
                    End If
                
                'This Else branch should never trigger, as it means the update file doesn't contain the listed update track.
                Else
                    InternalDebugMsg "WARNING!  Update XML file is possibly corrupt, as the requested update track could not be located within the file.", "ProcessProgramUpdateFile"
                End If
                
            Next i
            
            'If we found a track with a valid update target, initiate its download
            If (trackWithValidUpdate >= 0) Then
            
                'Make a backup copy of the update XML string.  We'll need to refer to it later, after the patch files have downloaded,
                ' as it contains failsafe checksum values.
                m_PDPatchXML = xmlEngine.ReturnCurrentXMLString(True)
                
                'We also want to cache the current update track at module-level, so we can display customized update notifications to
                ' the user.
                m_UpdateTrack = trackWithValidUpdate
                
                'Retrieve the manually listed beta number, just in case we need it later.  (For example, the current .exe may be
                ' Beta 1, and we're gonna update to Beta 2.)
                m_BetaNumber = xmlEngine.GetUniqueTag_String("releasenumber-beta", "1")
                
                'Construct a URL that matches the selected update track.  GitHub currently hosts PD's update downloads.
                Dim updateURL As String
                updateURL = "https://raw.githubusercontent.com/tannerhelland/PhotoDemon-Updates/master/auto/"
                
                Select Case trackWithValidUpdate
                
                    Case PDUT_STABLE
                        updateURL = updateURL & "stable"
                    
                    Case PDUT_BETA
                        updateURL = updateURL & "beta"
                
                    Case PDUT_NIGHTLY
                        updateURL = updateURL & "nightly"
                
                End Select
                
                'Download files ship using a custom archive format
                updateURL = updateURL & ".pdz"

                'Request a download from the main form.  Note that we also use the reported checksum as the file's
                ' unique ID value. (Post-download and extraction, this value will be used to ensure that the extracted
                ' patch data matches what we originally uploaded.)
                If FormMain.RequestAsynchronousDownload("PD_UPDATE_PATCH", updateURL, PD_PATCH_IDENTIFIER, vbAsyncReadForceUpdate, g_UserPreferences.GetUpdatePath & "PDPatch.tmp") Then
                    InternalDebugMsg "Now downloading update summary from " & updateURL, "ProcessProgramUpdateFile"
                    ProcessProgramUpdateFile = True
                Else
                    InternalDebugMsg "WARNING! FormMain.RequestAsynchronousDownload refused to download update summary (" & updateURL & ")", "ProcessProgramUpdateFile"
                End If
                
            'No newer version was found.  Exit now.
            Else
                InternalDebugMsg "Update check performed successfully.  (No update available right now.)", "ProcessProgramUpdateFile"
            End If
        
        Else
            InternalDebugMsg "WARNING! Program update XML did not pass basic validation.  Abandoning update process.", "ProcessProgramUpdateFile"
        End If
    
    Else
        InternalDebugMsg "WARNING! Program update XML did not load successfully - check for an encoding error, maybe...?", "ProcessProgramUpdateFile"
    End If
    
End Function

'When live-patching program files, we double-check checksums of both the temp files and the final binary copies.  This prevents
' hijackers from intercepting the files mid-transit, and replacing them with their own.
Private Function GetFailsafeChecksum(ByRef xmlEngine As pdXML, ByVal relativePath As String) As Long

    'Find the position of this file's checksum
    Dim pdTagPosition As Long
    pdTagPosition = xmlEngine.GetLocationOfTagPlusAttribute("checksum", "component", relativePath, m_TrackStartPosition)
    
    'Make sure the tag position is within the valid range.  (This should always be TRUE, but it doesn't hurt to check.)
    If (pdTagPosition >= m_TrackStartPosition) And (pdTagPosition <= m_TrackEndPosition) Then
    
        'This is the checksum tag we want!  Retrieve its value.
        Dim thisChecksum As String
        thisChecksum = xmlEngine.GetTagValueAtPreciseLocation(pdTagPosition)
        
        'Convert the checksum to a long and return it
        GetFailsafeChecksum = thisChecksum
        
    'If the checksum doesn't exist in the file, return 0
    Else
        GetFailsafeChecksum = 0
    End If
    
End Function

'If a program update file has successfully downloaded during this session, FormMain calls this function at program termination.
' This lovely function actually patches any/all relevant files.
Public Function PatchProgramFiles() As Boolean
    
    On Error GoTo ProgramPatchingFailure
    
    'If no update file is available, exit without doing anything
    If (Len(m_UpdateFilePath) = 0) Then
        PatchProgramFiles = True
        Exit Function
    End If
    
    'Write the update XML file out to file, so the separate patching app can access it
    Dim tmpXML As pdXML
    Set tmpXML = New pdXML
    tmpXML.LoadXMLFromString m_PDPatchXML
    tmpXML.WriteXMLToFile g_UserPreferences.GetUpdatePath & "patch.xml", True
    
    'The patching .exe is embedded inside the update package.  Extract it now.
    Dim cPackage As pdPackagerLegacy
    Set cPackage = New pdPackagerLegacy
    cPackage.Init_ZLib vbNullString, True, PluginManager.IsPluginCurrentlyEnabled(CCP_zLib)
    
    Dim patchFileName As String
    patchFileName = "PD_Update_Patcher.exe"
    
    If cPackage.ReadPackageFromFile(m_UpdateFilePath, PD_PATCH_IDENTIFIER) Then
        cPackage.AutoExtractSingleFile g_UserPreferences.GetProgramPath, patchFileName, , 99
    Else
        InternalDebugMsg "WARNING!  Patch program wasn't found inside the update package.  Patching will not proceed.", "PatchProgramFiles"
    End If
    
    'All that's left to do is shell the patch .exe.  It will wait for PD to close, then initiate the patching process.
    Dim patchParams As String
    If g_UserWantsRestart Then patchParams = "/restart"
    
    'We must tell the patcher where to find the update information
    patchParams = patchParams & " /start " & m_TrackStartPosition & " /end " & m_TrackEndPosition
    
    Dim targetPath As String
    targetPath = g_UserPreferences.GetProgramPath & patchFileName
    
    Dim shellReturn As Long
    shellReturn = ShellExecute(0, 0, StrPtr(targetPath), StrPtr(patchParams), 0, 0)
    
    'ShellExecute returns a value > 32 if successful (https://msdn.microsoft.com/en-us/library/windows/desktop/bb762153%28v=vs.85%29.aspx)
    If (shellReturn < 32) Then
        InternalDebugMsg "ShellExecute could not launch the updater!  It returned code #" & shellReturn & ".", "PatchProgramFiles"
    Else
        InternalDebugMsg "ShellExecute returned success.  Updater should launch momentarily...", "PatchProgramFiles"
    End If
    
    'Exit now
    PatchProgramFiles = True
    Exit Function
    
ProgramPatchingFailure:

    PatchProgramFiles = False
    
End Function

'Rather than apply updates mid-session, any patches are applied by a separate application, at shutdown time
Public Sub NotifyUpdatePackageAvailable(ByVal tmpUpdateFile As String)
    m_UpdateFilePath = tmpUpdateFile
End Sub

Public Function IsUpdatePackageAvailable() As Boolean
    IsUpdatePackageAvailable = (Len(m_UpdateFilePath) <> 0)
End Function

'Replacing files at run-time is unpredictable; sometimes we can delete the files, sometimes we can't.
'
'As such, this function is called when PD starts. It scans the update folder for old temp files and deletes them as encountered.
Public Sub CleanPreviousUpdateFiles()
    
    InternalDebugMsg "Looking for previous update files and cleaning as necessary...", "CleanPreviousUpdateFiles"
    
    'Use pdFSO to generate a list of .tmp files in the Update folder
    Dim tmpFileList As pdStringStack
    Set tmpFileList = New pdStringStack
    
    Dim tmpFile As String
        
    'First, we hard-code a few XML files that may exist due to old PD update methods
    tmpFileList.AddString g_UserPreferences.GetUpdatePath & "patch.xml"
    tmpFileList.AddString g_UserPreferences.GetUpdatePath & "pdupdate.xml"
    tmpFileList.AddString g_UserPreferences.GetUpdatePath & "updates.xml"
    
    'Next, we auto-add any .tmp files in the update folder, which should cover all other potential use-cases
    Files.RetrieveAllFiles g_UserPreferences.GetUpdatePath, tmpFileList, False, False, "TMP|tmp"
    
    'If temp files exist, remove them now.
    Do While tmpFileList.PopString(tmpFile)
        If Files.FileExists(tmpFile) Then
            Files.FileDeleteIfExists tmpFile
            InternalDebugMsg "deleting update file: " & tmpFile, "CleanPreviousUpdateFiles"
        End If
    Loop
        
    'Do the same thing for temp files in the base PD folder
    Set tmpFileList = Nothing
    If Files.RetrieveAllFiles(g_UserPreferences.GetProgramPath, tmpFileList, False, False, "TMP|tmp") Then
        
        Do While tmpFileList.PopString(tmpFile)
            If Files.FileExists(tmpFile) Then
                Files.FileDeleteIfExists tmpFile
                InternalDebugMsg "deleting update file: " & tmpFile, "CleanPreviousUpdateFiles"
            End If
        Loop
        
    End If
        
    '...And just to be safe, do the same thing for temp files in the plugin folder
    Set tmpFileList = Nothing
    If Files.RetrieveAllFiles(PluginManager.GetPluginPath, tmpFileList, False, False, "TMP|tmp") Then
        
        Do While tmpFileList.PopString(tmpFile)
            If Files.FileExists(tmpFile) Then
                Files.FileDelete tmpFile
                InternalDebugMsg "deleting update file: " & tmpFile, "CleanPreviousUpdateFiles"
            End If
        Loop
        
    End If
    
    'Finally, delete the patch exe itself, which will have closed by now
    Files.FileDeleteIfExists g_UserPreferences.GetProgramPath & "PD_Update_Patcher.exe"
    
End Sub

'At start-up, PD calls this function to find out if the program started via a PD-generated restart event (e.g. the presence of restart.bat).
' Returns TRUE if restart.bat is found; FALSE otherwise.
' (Also, this function deletes restart.bat if present)
Public Function WasProgramStartedViaRestart() As Boolean
    
    Dim restartFile As String
    restartFile = g_UserPreferences.GetProgramPath & "PD_Update_Patcher.exe"
    
    If Files.FileExists(restartFile) Then
        Files.FileDelete restartFile
        InternalDebugMsg "this session was started by an update process (PD_Update_Patcher is present)", "WasProgramStartedViaRestart"
        WasProgramStartedViaRestart = True
    Else
        InternalDebugMsg "FYI: this session was started by the user (PD_Update_Patcher is not present)", "WasProgramStartedViaRestart"
        WasProgramStartedViaRestart = False
    End If
    
End Function

'Every time PD is run, we have to do things like "see if it's time to check for an update".  This meta-function wraps all those
' behaviors into a single, caller-friendly function (currently called by FormMain_Load()).
Public Sub StandardUpdateChecks()
    
    'If PD is running in non-portable mode, we don't have write access to our own folder - which makes updates impossible.
    If g_UserPreferences.IsNonPortableModeActive() Then Exit Sub
    
    'See if this PD session was initiated by a PD-generated restart.  This happens after an update patch is successfully applied, for example.
    g_ProgramStartedViaRestart = Updates.WasProgramStartedViaRestart
        
    'Before updating, clear out any temp files leftover from previous updates.  (Replacing files at run-time is messy business, and Windows
    ' is unpredictable about allowing replaced files to be deleted.)
    Updates.CleanPreviousUpdateFiles
        
    'Start by seeing if we're allowed to check for software updates (the user can disable this check, and we want to honor their selection)
    Dim allowedToUpdate As Boolean
    allowedToUpdate = Updates.IsItTimeForAnUpdate()
    
    'If PD was restarted by an internal restart, disallow an update check now, as we would have just applied one (which caused the restart)
    If g_ProgramStartedViaRestart Then allowedToUpdate = False
    
    'If this is the user's first time using the program, don't pester them with update notifications
    If g_IsFirstRun Then allowedToUpdate = False
    
    'If we're STILL allowed to update, do so now (unless this is the first time the user has run the program; in that case, suspend updates,
    ' as it is assumed the user already has an updated copy of the software - and we don't want to bother them already!)
    If allowedToUpdate Then
    
        Message "Initializing software updater (this feature can be disabled from the Tools -> Options menu)..."
        
        'Initiate an asynchronous download of the standard PD update file (currently hosted @ GitHub).
        ' When the asynchronous download completes, the downloader will place the completed update file in the /Data/Updates subfolder.
        ' On exit (or subsequent program runs), PD will check for the presence of that file, then proceed accordingly.
        FormMain.RequestAsynchronousDownload "PROGRAM_UPDATE_CHECK", "https://raw.githubusercontent.com/tannerhelland/PhotoDemon-Updates/master/summary/pdupdate.xml", , vbAsyncReadForceUpdate, g_UserPreferences.GetUpdatePath & "updates.xml"
        
    End If
    
    'With all potentially required downloads added to the queue, we can now begin downloading everything
    FormMain.asyncDownloader.SetAutoDownloadMode True
    
End Sub

'If an update is ready, you may call this function to display an update notification to the user
Public Sub DisplayUpdateNotification()
    
    'If a modal dialog is active, raising a new window will cause a crash; we must deal with this accordingly
    On Error GoTo CouldNotDisplayUpdateNotification
    
    'Suspend any previous update notification flags
    g_ShowUpdateNotification = False
    
    'Check user preferences; they can choose to ignore update notifications
    If g_UserPreferences.GetPref_Boolean("Updates", "Update Notifications", True) Then
        
        'Display the dialog, while yielding for the rare case that a modal dialog is already active
        If Interface.IsModalDialogActive() Then
            g_ShowUpdateNotification = True
        Else
            FormUpdateNotify.Show vbModeless, FormMain
        End If
        
    End If
    
    Exit Sub
    
CouldNotDisplayUpdateNotification:

    'Set a global flag; PD's central processor will use this to display the notification as soon as it reasonably can
    g_ShowUpdateNotification = True

End Sub

'PD should always be able to provide a release announcement URL, but I still recommend testing this string for emptiness prior to displaying
' it to the user.
Public Function GetReleaseAnnouncementURL() As String
    GetReleaseAnnouncementURL = m_UpdateReleaseAnnouncementURL
End Function

'Outside functions can also the track of the currently active update.  Note that this doesn't predictably correspond to the user's current
' update preference, as most users will allow updates from multiple potential tracks (e.g. both stable and beta).
Public Function GetUpdateTrack() As PD_UPDATE_TRACK
    GetUpdateTrack = m_UpdateTrack
End Function

'Outside functions can also request a human-readable string of the literal update number (e.g. Major.Minor.Build, untouched).
Public Function GetUpdateVersion_Literal(Optional ByVal forceRevisionDisplay As Boolean = False) As String
    
    'Parse the version string, which is currently on the form Major.Minor.Build.Revision
    Dim verStrings() As String
    verStrings = Split(m_UpdateVersion, ".")
    If (UBound(verStrings) < 2) Then verStrings = Split(m_UpdateVersion, ",")
    
    'We always want major and minor version numbers
    If (UBound(verStrings) >= 1) Then
        
        GetUpdateVersion_Literal = verStrings(0) & "." & verStrings(1)
        
        'If the revision value is non-zero, or the user demands a revision number, include it
        If (UBound(verStrings) >= 3) Or forceRevisionDisplay Then
            
            'If the revision number exists, use it
            If (UBound(verStrings) >= 3) Then
                If Strings.StringsNotEqual(verStrings(3), "0", False) Then GetUpdateVersion_Literal = GetUpdateVersion_Literal & "." & verStrings(3)
            
            'If the revision number does not exist, append 0 in its place
            Else
                GetUpdateVersion_Literal = GetUpdateVersion_Literal & ".0"
            End If
            
        End If
        
    Else
        GetUpdateVersion_Literal = m_UpdateVersion
    End If
    
End Function

'Outside functions can use this to request a human-readable string of the "friendly" update number (e.g. beta releases are
' properly identified and bumped up to the next stable release).
Public Function GetUpdateVersion_Friendly() As String
    
    'Start by retrieving the literal version number
    Dim litVersion As String
    litVersion = GetUpdateVersion_Literal(True)
    
    'If the current update track is *NOT* a beta, the friendly string matches the literal string.  Return it now.
    If (m_UpdateTrack <> PDUT_BETA) Then
        GetUpdateVersion_Friendly = litVersion
    
    'If the current update track *IS* a beta, we need to manually update the number prior to returning it
    Else
        
        On Error GoTo VersionFormatError
        
        'Start by extracting all version numbers
        Dim vSplit() As String
        vSplit = Split(litVersion, ".")
        
        Dim vMajor As Long, vMinor As Long, vRevision As Long
        
        vMajor = vSplit(0)
        vMinor = vSplit(1)
        vRevision = vSplit(2)
        
        'Bump minor by 1
        vMinor = vMinor + 1
        
        'Account for .10, which means a release to the next major version (e.g. 6.9 leads to 7.0, not 6.10)
        If (vMinor = 10) Then
            vMinor = 0
            vMajor = vMajor + 1
        End If
        
        'Construct a new version string
        GetUpdateVersion_Friendly = g_Language.TranslateMessage("%1.%2 Beta %3", vMajor, vMinor, m_BetaNumber)
        
    End If
    
    Exit Function
    
VersionFormatError:

    GetUpdateVersion_Friendly = litVersion

End Function

'Retrieve PD's current name and version, modified against "beta" labels, etc
Public Function GetPhotoDemonNameAndVersion() As String
    GetPhotoDemonNameAndVersion = App.Title & " " & Updates.GetPhotoDemonVersion()
End Function

'Retrieve PD's current version, modified against "beta" labels, etc
Public Function GetPhotoDemonVersion() As String
    
    'Even-numbered releases are "official" releases, so simply return the full version string
    If (CLng(App.Minor) Mod 2 = 0) Then
        GetPhotoDemonVersion = App.Major & "." & App.Minor
        If (App.Revision <> 0) Then GetPhotoDemonVersion = GetPhotoDemonVersion & "." & App.Revision
        
    Else
    
        'Odd-numbered development releases of the pattern X.9 are production builds for the next major version, e.g. (X+1).0
        
        'Build state can be retrieved from the public const PD_BUILD_QUALITY
        Dim buildStateString As String
        
        Select Case PD_BUILD_QUALITY
        
            Case PD_PRE_ALPHA
                If (g_Language Is Nothing) Then
                    buildStateString = "pre-alpha"
                Else
                    buildStateString = g_Language.TranslateMessage("pre-alpha")
                End If
            
            Case PD_ALPHA
                If (g_Language Is Nothing) Then
                    buildStateString = "alpha"
                Else
                    buildStateString = g_Language.TranslateMessage("alpha")
                End If
            
            Case PD_BETA
                If (g_Language Is Nothing) Then
                    buildStateString = "beta"
                Else
                    buildStateString = g_Language.TranslateMessage("beta")
                End If
        
        End Select
        
        'Assemble a full title string, while handling the special case of .9 version numbers, which serve as production
        ' builds for the next .0 release.
        If (App.Minor = 9) Then
            GetPhotoDemonVersion = CStr(App.Major + 1) & ".0 " & buildStateString & " (build " & CStr(App.Revision) & ")"
        Else
            GetPhotoDemonVersion = CStr(App.Major) & "." & CStr(App.Minor + 1) & " " & buildStateString & " (build " & CStr(App.Revision) & ")"
        End If
        
    End If
    
End Function

'Retrieve PD's current version witout any appended tags (e.g. "beta"), and with a "0" automatically plugged in for build.
Public Function GetPhotoDemonVersionCanonical() As String
    GetPhotoDemonVersionCanonical = Trim$(Str(App.Major)) & "." & Trim$(Str(App.Minor)) & ".0." & Trim$(Str(App.Revision))
End Function

'Retrieve PD's current version (not revision!) as a pure major/minor string.  This is not generally recommended for displaying
' to the user, but it's helpful for things like update checks.
Public Function GetPhotoDemonVersionMajorMinorOnly() As String
    GetPhotoDemonVersionMajorMinorOnly = Trim$(Str(App.Major)) & "." & Trim$(Str(App.Minor))
End Function

Public Function GetPhotoDemonVersionRevisionOnly() As String
    GetPhotoDemonVersionRevisionOnly = Trim$(Str(App.Revision))
End Function

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

'Given two version numbers, return TRUE if the second version is larger than the first.
' If the second version equals the first, FALSE is returned.
Public Function IsNewVersionHigher(ByVal oldVersion As String, ByVal newVersion As String) As Boolean
    
    'Normalize version separators
    If InStr(1, oldVersion, ",", vbBinaryCompare) Then oldVersion = Replace$(oldVersion, ",", ".")
    If InStr(1, newVersion, ",", vbBinaryCompare) Then oldVersion = Replace$(newVersion, ",", ".")
    
    'If the string representations are identical, we can exit now
    If Strings.StringsEqual(oldVersion, newVersion, False) Then
        IsNewVersionHigher = False
        
    'If the strings are not equal, a more detailed comparison is required.
    Else
    
        'Parse the versions by "."
        Dim oldV() As String, newV() As String
        oldV = Split(oldVersion, ".")
        newV = Split(newVersion, ".")
        
        'Fill in any missing version entries
        Dim i As Long, oldUBound As Long
        
        If (UBound(oldV) < 3) Then
            
            oldUBound = UBound(oldV)
            ReDim Preserve oldV(0 To 3) As String
            
            For i = oldUBound + 1 To 3
                oldV(i) = "0"
            Next i
            
        End If
        
        If (UBound(newV) < 3) Then
            
            oldUBound = UBound(newV)
            ReDim Preserve newV(0 To 3) As String
            
            For i = oldUBound + 1 To 3
                newV(i) = "0"
            Next i
            
        End If
        
        'To simplify comparisons, convert the string arrays to numeric ones
        Dim oldVersionNums(0 To 3) As Long
        Dim newVersionNums(0 To 3) As Long
        For i = 0 To 3
            oldVersionNums(i) = CLng(oldV(i))
            newVersionNums(i) = CLng(newV(i))
        Next i
        
        'With both version numbers normalized, compare each entry in turn.
        Dim newIsNewer As Boolean: newIsNewer = False
        
        'For each version, we will compare numbers in turn, starting with the major version and working
        ' our way down.  We only check subsequent values if all preceding ones are equal.  (This ensures
        ' that e.g. 6.6.0 does not update to 6.5.1.)
        Dim majorIsEqual As Boolean, minorIsEqual As Boolean, revIsEqual As Boolean
                
        For i = 0 To 3
            
            Select Case i
            
                'Major version updates always trigger an update
                Case 0
                
                    If (newVersionNums(i) > oldVersionNums(i)) Then
                        newIsNewer = True
                        Exit For
                    Else
                        majorIsEqual = (newVersionNums(i) = oldVersionNums(i))
                    End If
                
                'Minor version updates trigger an update only if the major version matches (e.g. 1.0 will update to 1.2,
                ' but 2.0 will not update to 1.2)
                Case 1
                
                    If majorIsEqual Then
                        If (newVersionNums(i) > oldVersionNums(i)) Then
                            newIsNewer = True
                            Exit For
                        Else
                            minorIsEqual = (newVersionNums(i) = oldVersionNums(i))
                        End If
                    End If
                
                'Build and revision updates follow the pattern above
                Case 2
                
                    If minorIsEqual Then
                        If (newVersionNums(i) > oldVersionNums(i)) Then
                            newIsNewer = True
                            Exit For
                        Else
                            revIsEqual = (newVersionNums(i) = oldVersionNums(i))
                        End If
                    End If
                
                Case Else
                
                    If revIsEqual Then
                        newIsNewer = (newVersionNums(i) > oldVersionNums(i))
                        Exit For
                    End If
                
            End Select
            
        Next i
        
        IsNewVersionHigher = newIsNewer
        
    End If
    
End Function

'Updates involve the Internet, so any number of things can go wrong.  PD versions (post-7.0) post a lot of internal
' debug data throughout this module, to help us identify problems where we can.  Instead of inserting your own
' "#IF DEBUGMODE = 1" lines, please route debugging data through here.
Private Sub InternalDebugMsg(ByRef srcMsg As String, ByRef srcFunctionName As String, Optional ByVal errNumber As Long = 0, Optional ByRef errDescription As String = vbNullString)

    #If DEBUGMODE = 1 Then
        If (errNumber <> 0) Then
            pdDebug.LogAction "WARNING!  Updates." & srcFunctionName & " reported an error (#" & CStr(errNumber) & "): " & errDescription & ".  Further details: " & srcMsg
        Else
            pdDebug.LogAction "Updates." & srcFunctionName & " reports: " & srcMsg, , True
        End If
    #End If

End Sub
