Attribute VB_Name = "Updates"
'***************************************************************************
'Automatic App Updater
'Copyright 2012-2026 by Tanner Helland
'Created: 19/August/12
'Last updated: 02/April/24
'Last update: rework to prep for a more rapid release schedule (using year.month version numbers)
'
'This module includes support functions for determining if a new version of PhotoDemon is available
' for automatic patching.
'
'IMPORTANT NOTE: this module doesn't do the actual updating (e.g. overwriting program files); it just
' CHECKS for updates.  Patching is handled by a separate exe.
'
'As of March 2015, this module has been completely overhauled to support live-patching of PhotoDemon
' and its various support files (plugins, languages, etc).  Various bits of update code have been moved
' into the new update support app in the /Support folder.  The use of a separate patching app greatly
' simplifies things like updating in-use binary files.
'
'Note that this code interfaces with the user preferences file so the user can opt to not check for
' updates and never be notified again. (FYI - this option can always be toggled from the 'Tools' ->
' 'Options' menu.)
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

Private Declare Function ShellExecuteW Lib "shell32" (ByVal hWnd As Long, ByVal ptrToOperationString As Long, ByVal ptrToFileString As Long, ByVal ptrToParameters As Long, ByVal ptrToDirectory As Long, ByVal nShowCmd As Long) As Long

'When initially parsing the update XML file (above), if an update is found, the parse routine will note which
' track was used for the update, and where that track's data starts and ends inside the XML file.
Private m_SelectedTrack As Long

'If an update package is downloaded successfully, it will be forwarded to this module.  At program shutdown time,
' the package will be applied.
Private m_UpdateFilePath As String

'If an update is available, that update's release announcement will be stored in this persistent string.
' UI elements can retrieve it as necessary.
Private m_UpdateReleaseAnnouncementURL As String

'If an update is available, that update's track will be stored here.  UI elements can retrieve it as necessary.
Private m_UpdateTrack As PD_UpdateTrack

'Outside functions can also request the update version
Private m_UpdateVersion As String

'Beta releases use custom labeling, independent of the actual version (e.g. "PD 6.6 beta 3"), so we also retrieve
' and store this value as necessary.
Private m_BetaNumber As String

'If the user wants to restart the program after applying an update (vs just applying it at shutdown)
Private m_UserRequestedRestart As Boolean

'If an update is available, this will be set to TRUE.  (Other parts of the app can query this value
' and raise a notification accordingly.)
Private m_UpdateReady As Boolean

'Canonical app version; calculated once, on first access, then cached for subsequent calls
Private m_CanonicalVersion As String

Public Function GetRestartAfterUpdate() As Boolean
    GetRestartAfterUpdate = m_UserRequestedRestart
End Function

Public Sub SetRestartAfterUpdate(ByVal newState As Boolean)
    m_UserRequestedRestart = newState
End Sub

Public Function IsUpdateReadyToInstall() As Boolean
    IsUpdateReadyToInstall = m_UpdateReady
End Function

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
    If UserPrefs.DoesValueExist("Updates", "CheckForUpdates") Then
        
        'Write a matching preference in the new format, and overwrite the old preference (so it doesn't trigger this
        ' check again)
        If (Not UserPrefs.GetPref_Boolean("Updates", "CheckForUpdates", True)) Then
            UserPrefs.SetPref_Long "Updates", "Update Frequency", PDUF_NEVER
            UserPrefs.SetPref_Boolean "Updates", "CheckForUpdates", True
        End If
        
    End If
    
    'In v6.6, PD's update strategy was modified to allow the user to specify an update frequency (rather than
    ' a binary yes/no preference).  Retrieve the allowed frequency now.
    If (updateFrequency <> PDUF_NEVER) Then updateFrequency = UserPrefs.GetPref_Long("Updates", "Update Frequency", PDUF_EACH_SESSION)
    
    'If updates ARE allowed, see when we last checked for an update.  If enough time has elapsed, check again.
    If (updateFrequency <> PDUF_NEVER) Then
    
        Dim lastCheckDate As String
        lastCheckDate = UserPrefs.GetPref_String("Updates", "Last Update Check")
        
        'If a "last update check date" was not found, request an immediate update check.
        If (LenB(lastCheckDate) = 0) Then
            allowedToUpdate = True
        
        'If a last update check date was found, check to see how much time has elapsed since that check.
        Else
        
            'Start by figuring out how many days need to have passed before we're allowed to check for updates
            ' again.  (This varies according to user preference.)
            Dim numAllowableDays As Long
            If (updateFrequency = PDUF_EACH_SESSION) Then
                numAllowableDays = 0
            ElseIf (updateFrequency = PDUF_WEEKLY) Then
                numAllowableDays = 7
            ElseIf (updateFrequency = PDUF_MONTHLY) Then
                numAllowableDays = 30
            
            'This else should never trigger.
            Else
                numAllowableDays = 180
            End If
            
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

'tl;dr: given an XML report from photodemon.org, initiate a program update package download, as necessary.
'
'Long explanation: this function checks to see if PhotoDemon.exe is out of date against the current
' update track (stable, beta, or nightly, per the current user preference).  If the .exe *is* out of date,
' the latest update package will be downloaded.
'
'Returns: TRUE if an update is available *and* its download was initiated successfully; FALSE otherwise
Public Function ProcessProgramUpdateFile(ByRef srcXML As String) As Boolean
    
    'Start by figuring out which update track we need to check.  The user can change this at any time,
    ' so it may not correlate to this .exe's build type.  (For example, maybe this is a stable PD build,
    ' but the user switched update checks to include beta and nightly builds - that's okay!)
    Dim curUpdateTrack As PD_UpdateTrack
    curUpdateTrack = UserPrefs.GetPref_Long("Updates", "Update Track", ut_Beta)
    
    'From the update track, we need to generate a string that identifies the correct tag to check for
    ' version numbers.  This is a little tricky because some update tracks can update to more than one type
    ' of build - for example, nightly builds can update to stable builds, if the stable version is newer,
    ' but stable builds can't update to nightly builds unless the user's preferences explicitly allow it.
    '
    'As such, we may need to search multiple HTML tags to find a relevant update target.
    Dim updateTagIDs() As String, numUpdateTagIDs As Long
    ReDim updateTagIDs(0 To 2) As String
    updateTagIDs(0) = "stable"
    updateTagIDs(1) = "beta"
    updateTagIDs(2) = "developer"
    
    Select Case curUpdateTrack
    
        Case ut_Stable
            numUpdateTagIDs = 1
            
        Case ut_Beta
            numUpdateTagIDs = 2
        
        Case ut_Developer
            numUpdateTagIDs = 3
        
    End Select
    
    ReDim Preserve updateTagIDs(0 To numUpdateTagIDs - 1) As String
    
    'If we find an update track that provides a valid update target, this value will point at that
    ' track's index (0, 1, or 2, for stable, beta, or nightly, respectively).
    '
    'If no update is found, it will remain at -1.
    Dim trackWithValidUpdate As PD_UpdateTrack
    trackWithValidUpdate = ut_None
    
    'We start with the current PD version as a baseline.  If newer update targets are found,
    ' this string will be updated with a newer version number, instead.
    Dim curVersionMatch As String
    curVersionMatch = GetPhotoDemonVersionCanonical()
    
    'If you want to test against random version numbers, feel free to plug in a custom test version number...
    'curVersionMatch = "6.4.0"
    
    Dim i As Long
    Dim newPDVersionString As String
        
    'The new update file is (literally) just the index.html page of PD's GitHub Pages update server
    ' (https://tannerhelland.github.io/PhotoDemon-Updates-v2/).
    '
    'We want to compare against the specific release numbers listed on that page.
    
    'To do that, we're gonna search for each updateTagID region (as calculated above).
    ' If any return a hit, we'll take the newest one and start downloading its update package.
    For i = 0 To numUpdateTagIDs - 1
    
        'Find the tag in question by looking for specifically formatted html bounding regions.
        Dim startTagText As String, endTagText As String
        startTagText = "<a id=""pdv_start_" & updateTagIDs(i) & """></a>"
        endTagText = "<a id=""pdv_end_" & updateTagIDs(i) & """></a>"
        
        Dim tagStartPos As Long, tagEndPos As Long
        tagStartPos = Strings.StrStrBM(srcXML, startTagText)
        If (tagStartPos > 0) Then
            tagStartPos = tagStartPos + Len(startTagText)
            tagEndPos = InStr(tagStartPos, srcXML, endTagText)
        End If
        
        'If valid positions were found, retrieve the text between them
        newPDVersionString = vbNullString
        If (tagStartPos <> 0) And (tagEndPos <> 0) And (tagEndPos > tagStartPos) Then
            
            newPDVersionString = Mid$(srcXML, tagStartPos, tagEndPos - tagStartPos)
            InternalDebugMsg "Update track " & i & " reports version " & newPDVersionString & " (our version: " & GetPhotoDemonVersionCanonical() & ")", "ProcessProgramUpdateFile"
            
            'If this value is newer than our current update target, mark it and proceed.  Note that this approach gives
            ' us the highest possible update target from all available/enabled update tracks.
            If IsNewVersionHigher(curVersionMatch, newPDVersionString) Then
                
                trackWithValidUpdate = i
                
                'Set some matching module-level values, which we'll need when it's time to actually patch the files
                ' in question.
                m_SelectedTrack = trackWithValidUpdate
                m_UpdateVersion = newPDVersionString
                
                'Retrieving the announcement URL is a little strange, since it's embedded in the page as a clickable link,
                ' but we've marked it specially to make it easy to find.
                m_UpdateReleaseAnnouncementURL = vbNullString
                
                Dim hRefStart As Long, hRefEnd As Long
                hRefStart = InStr(1, srcXML, "<a id=""pdra_" & updateTagIDs(i) & """ href=""", vbBinaryCompare)
                If (hRefStart > 0) Then hRefEnd = InStr(hRefStart, srcXML, "</a>", vbBinaryCompare)
                If (hRefEnd > hRefStart) Then
                
                    hRefStart = InStr(hRefStart, srcXML, "href=""", vbBinaryCompare)
                    If (hRefStart <> 0) Then
                        hRefStart = hRefStart + Len("href=""")
                        hRefEnd = InStr(hRefStart + 1, srcXML, """", vbBinaryCompare)
                        If (hRefEnd > hRefStart) Then m_UpdateReleaseAnnouncementURL = Mid$(srcXML, hRefStart, hRefEnd - hRefStart)
                    End If
                
                End If
                
            End If
        
        Else
            InternalDebugMsg "invalid tag positions found: " & tagStartPos & ", " & tagEndPos, "ProcessProgramUpdateFile"
        End If
        
    Next i
    
    'If we found a track with a valid update target, initiate its download
    If (trackWithValidUpdate >= 0) Then
    
        'Cache the current update track at module-level, so we can display customized update notifications
        ' to the user.
        m_UpdateTrack = trackWithValidUpdate
        
        'Retrieve the manually listed beta number, just in case we need it later.
        ' (For example, the current .exe may be Beta 1, and we're gonna update to Beta 2.)
        startTagText = "<a id=""pdv_beta_num_start""></a>"
        endTagText = "<a id=""pdv_beta_num_end""></a>"
        
        tagStartPos = Strings.StrStrBM(srcXML, startTagText)
        If (tagStartPos > 0) Then
            tagStartPos = tagStartPos + Len(startTagText)
            tagEndPos = InStr(tagStartPos, srcXML, endTagText)
        End If
        
        If (tagStartPos <> 0) And (tagEndPos <> 0) And (tagEndPos > tagStartPos) Then
            m_BetaNumber = Mid$(srcXML, tagStartPos, tagEndPos - tagStartPos)
        Else
            m_BetaNumber = g_Language.TranslateMessage("unknown")
        End If
        
        'Construct a URL that matches the selected update track.  GitHub currently hosts PD's update downloads.
        Dim updateURL As String
        updateURL = "https://tannerhelland.github.io/PhotoDemon-Updates-v2/auto/"
        
        Select Case trackWithValidUpdate
        
            Case ut_Stable
                updateURL = updateURL & "stable"
            
            Case ut_Beta
                updateURL = updateURL & "beta"
        
            Case ut_Developer
                updateURL = updateURL & "nightly"
        
        End Select
        
        'Download files ship using a custom archive format
        updateURL = updateURL & ".pdz2"
        
        'Request a download from the main form.
        If FormMain.RequestAsynchronousDownload("PD_UPDATE_PATCH", updateURL, PD_PATCH_IDENTIFIER, vbAsyncReadForceUpdate, UserPrefs.GetUpdatePath & "PDPatch.tmp") Then
            InternalDebugMsg "Now downloading update summary from " & updateURL, "ProcessProgramUpdateFile"
            ProcessProgramUpdateFile = True
        Else
            InternalDebugMsg "WARNING! FormMain.RequestAsynchronousDownload refused to download update patch (" & updateURL & ")", "ProcessProgramUpdateFile"
        End If
        
    'No newer version was found.  Exit now.
    Else
        InternalDebugMsg "Update check performed successfully.  (No update available right now.)", "ProcessProgramUpdateFile"
    End If
    
End Function

'If a program update file has successfully downloaded during this session, FormMain calls this function at program termination.
' This lovely function actually patches any/all relevant files.
Public Function PatchProgramFiles() As Boolean
    
    On Error GoTo ProgramPatchingFailure
    
    'If no update file is available, exit without doing anything
    If (LenB(m_UpdateFilePath) = 0) Then
        PatchProgramFiles = True
        Exit Function
    End If
    
    'The patching .exe is embedded inside the update package.  Extract it now; it will handle the rest
    ' of the patching process after we exit.
    Dim cPackage As pdPackageLegacyV2
    Set cPackage = New pdPackageLegacyV2
    
    Dim patchFileName As String
    patchFileName = "\PD_Update_Patcher.exe"
    
    If cPackage.ReadPackageFromFile(m_UpdateFilePath, PD_PATCH_IDENTIFIER) Then
        cPackage.AutoExtractSingleFile UserPrefs.GetProgramPath, patchFileName, False, 99
    Else
        InternalDebugMsg "WARNING!  Patch program wasn't found inside the update package.  Patching will not proceed.", "PatchProgramFiles"
    End If
    
    'All that's left to do is shell the patch .exe.  It will wait for PD to close, then initiate the patching process.
    Dim patchParams As String
    If m_UserRequestedRestart Then patchParams = "/restart"
    patchParams = patchParams & " /sourceIsPD"
    
    Dim targetPath As String
    targetPath = UserPrefs.GetProgramPath & patchFileName
    
    Dim shellReturn As Long
    shellReturn = ShellExecuteW(0, 0, StrPtr(targetPath), StrPtr(patchParams), 0, 0)
    
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
Public Sub NotifyUpdatePackageAvailable(ByRef tmpUpdateFile As String)
    m_UpdateFilePath = tmpUpdateFile
End Sub

Public Function IsUpdatePackageAvailable() As Boolean
    IsUpdatePackageAvailable = (LenB(m_UpdateFilePath) <> 0)
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
    tmpFileList.AddString UserPrefs.GetUpdatePath & "patch.xml"
    tmpFileList.AddString UserPrefs.GetUpdatePath & "pdupdate.xml"
    tmpFileList.AddString UserPrefs.GetUpdatePath & "updates.xml"
    
    'Next, we auto-add any .tmp files in the update folder, which should cover all other potential use-cases
    Files.RetrieveAllFiles UserPrefs.GetUpdatePath, tmpFileList, False, False, "TMP|tmp"
    
    'If temp files exist, remove them now.
    Do While tmpFileList.PopString(tmpFile)
        If Files.FileExists(tmpFile) Then
            Files.FileDeleteIfExists tmpFile
            InternalDebugMsg "deleting update file: " & tmpFile, "CleanPreviousUpdateFiles"
        End If
    Loop
        
    'Do the same thing for temp files in the base PD folder
    Set tmpFileList = Nothing
    If Files.RetrieveAllFiles(UserPrefs.GetProgramPath, tmpFileList, False, False, "TMP|tmp") Then
        
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
    Files.FileDeleteIfExists UserPrefs.GetProgramPath & "PD_Update_Patcher.exe"
    
End Sub

'At start-up, PD calls this function to find out if the program started via a PD-generated restart event (e.g. the presence of restart.bat).
' Returns TRUE if restart.bat is found; FALSE otherwise.
' (Also, this function deletes restart.bat if present)
Public Function WasProgramStartedViaRestart() As Boolean
    
    Dim restartFile As String
    restartFile = UserPrefs.GetProgramPath & "PD_Update_Patcher.exe"
    
    If Files.FileExists(restartFile) Then
        Files.FileDelete restartFile
        InternalDebugMsg "this session was started by an update process (PD_Update_Patcher is present)", "WasProgramStartedViaRestart"
        WasProgramStartedViaRestart = True
    Else
        InternalDebugMsg "FYI: this session was started by the user (PD_Update_Patcher is not present)", "WasProgramStartedViaRestart"
        WasProgramStartedViaRestart = False
    End If
    
End Function

'Every time PD is run, we have to do things like "see if it's time to check for an update".  This meta-function
' wraps all those behaviors into a single, caller-friendly function (currently called by FormMain_Load()).
Public Sub StandardUpdateChecks()
    
    'If PD is running in non-portable mode, we don't have write access to our own folder;
    ' this makes updates impossible, so we skip the entire process.
    If UserPrefs.IsNonPortableModeActive() Then Exit Sub
    
    'See if this PD session was initiated by a PD-generated restart.  This happens after an update patch is
    ' successfully applied, for example - and if it happens, we want to know later in the function, so we
    ' can skip an update check this session.
    Dim appWasJustRestarted As Boolean
    appWasJustRestarted = Updates.WasProgramStartedViaRestart
        
    'Before updating, clear out any temp files leftover from previous updates.  (Replacing files at run-time
    ' is messy business, and Windows is sometimes unpredictable about allowing replaced files to be deleted.)
    Updates.CleanPreviousUpdateFiles
        
    'Start by seeing if we're even allowed to check for software updates.  (Note that this step is multifaceted;
    ' the user can disable update checks entirely, or they can enable them at a specific interval.  If either test
    ' fails, we skip further checks, without caring about the reason "why".)
    Dim allowedToUpdate As Boolean
    allowedToUpdate = Updates.IsItTimeForAnUpdate()
    
    'If this PD session was the result of an internal restart (e.g. an automatic update *just* finished),
    ' disallow this session's update check.
    If appWasJustRestarted Then allowedToUpdate = False
    
    'If this is the user's first time using the program, don't pester them with update notifications
    If g_IsFirstRun Then allowedToUpdate = False
    
    'If we're STILL allowed to update, do so now (unless this is the first time the user has run the program; in that case, suspend updates,
    ' as it is assumed the user already has an updated copy of the software - and we don't want to bother them already!)
    If allowedToUpdate Then
    
        Message "Initializing software updater (this feature can be disabled from the Tools -> Options menu)..."
        
        'Initiate an asynchronous download of the standard PD update file (currently hosted @ GitHub).
        ' When the asynchronous download completes, the downloader will place the completed update file in the /Data/Updates subfolder.
        ' On exit (or subsequent program runs), PD will check for the presence of that file, then proceed accordingly.
        Dim srcPath As String
        srcPath = "https://tannerhelland.github.io/PhotoDemon-Updates-v2/"
        FormMain.RequestAsynchronousDownload "PROGRAM_UPDATE_CHECK", srcPath, , vbAsyncReadForceUpdate, UserPrefs.GetUpdatePath & "updates.xml"
        
    End If
    
    'With all potentially required downloads added to the queue, we can now begin downloading everything
    FormMain.AsyncDownloader.SetAutoDownloadMode True
    
End Sub

'If an update is ready, you may call this function to display an update notification to the user
Public Sub DisplayUpdateNotification()
    
    'If a modal dialog is active, raising a new window will cause a crash; we must deal with this accordingly
    On Error GoTo CouldNotDisplayUpdateNotification
    
    'Suspend any previous update notification flags
    m_UpdateReady = False
    
    'Check user preferences; they can choose to ignore update notifications
    If UserPrefs.GetPref_Boolean("Updates", "Update Notifications", True) Then
        
        'Display the dialog, while yielding for the rare case that a modal dialog is already active
        If Interface.IsModalDialogActive() Then
            m_UpdateReady = True
        Else
            FormUpdateNotify.Show vbModeless, FormMain
        End If
        
    End If
    
    Exit Sub
    
CouldNotDisplayUpdateNotification:

    'Set a global flag; PD's central processor will use this to display the notification as soon as it reasonably can
    m_UpdateReady = True

End Sub

'PD should always be able to provide a release announcement URL, but I still recommend testing this string for emptiness prior to displaying
' it to the user.
Public Function GetReleaseAnnouncementURL() As String
    GetReleaseAnnouncementURL = m_UpdateReleaseAnnouncementURL
End Function

'Outside functions can also the track of the currently active update.  Note that this doesn't predictably correspond to the user's current
' update preference, as most users will allow updates from multiple potential tracks (e.g. both stable and beta).
Public Function GetUpdateTrack() As PD_UpdateTrack
    GetUpdateTrack = m_UpdateTrack
End Function

'Produce a human-readable string of the literal update number (e.g. Major.Minor.Build).
Private Function GetUpdateVersion_Literal(Optional ByVal forceRevisionDisplay As Boolean = False) As String
    
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
    If (m_UpdateTrack <> ut_Beta) Then
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
    
    'Build state can be retrieved from the compile-time const PD_BUILD_QUALITY
    Dim buildStateString As String
    
    Select Case PD_BUILD_QUALITY
    
        Case PD_ALPHA
            If (g_Language Is Nothing) Then
                buildStateString = "alpha"
            Else
                buildStateString = g_Language.TranslateMessage("alpha", SPECIAL_TRANSLATION_OBJECT_PREFIX & "version-alpha")
            End If
        
        Case PD_BETA
            If (g_Language Is Nothing) Then
                buildStateString = "beta"
            Else
                buildStateString = g_Language.TranslateMessage("beta", SPECIAL_TRANSLATION_OBJECT_PREFIX & "version-beta")
            End If
        
        'No special text is required for production builds
        Case Else
            buildStateString = vbNullString
    
    End Select
    
    'Strip exact build numbers from production builds; on all other builds, include a text description
    ' (e.g. "alpha" or "beta") and full build number so I can more easily track bug reports.
    GetPhotoDemonVersion = CStr(VBHacks.AppMajor_Safe()) & "." & CStr(VBHacks.AppMinor_Safe())
    If (PD_BUILD_QUALITY <> PD_PRODUCTION) Then GetPhotoDemonVersion = GetPhotoDemonVersion & " " & buildStateString & " (build " & CStr(VBHacks.AppRevision_Safe()) & ")"
    
End Function

'Retrieve PD's current version witout any appended tags (e.g. "beta"), and with a "0" automatically plugged in for build.
Public Function GetPhotoDemonVersionCanonical() As String
    If (LenB(m_CanonicalVersion) = 0) Then m_CanonicalVersion = Trim$(Str$(VBHacks.AppMajor_Safe())) & "." & Trim$(Str$(VBHacks.AppMinor_Safe())) & ".0." & Trim$(Str$(VBHacks.AppRevision_Safe()))
    GetPhotoDemonVersionCanonical = m_CanonicalVersion
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

'Raise a modal message box that offers to download a new copy of some third-party plugin.
' This function uses general language appropriate for any plugin; it's up to the caller to do something
' interesting based on the return of this function (TRUE if the user consents to the update).
Public Function OfferPluginUpdate(Optional ByRef pluginName As String = vbNullString, Optional ByRef curVersion As String = vbNullString, Optional ByRef newVersion As String = vbNullString) As VbMsgBoxResult
    
    Dim uiMsg As pdString
    Set uiMsg = New pdString
    
    uiMsg.AppendLine g_Language.TranslateMessage("This action relies on a third-party plugin (%1).  An updated version of this plugin is available.", pluginName)
    uiMsg.AppendLineBreak
    
    If (LenB(curVersion) >= 0) And (LenB(newVersion) >= 0) Then
        uiMsg.AppendLine g_Language.TranslateMessage("Your version: %1", curVersion)
        uiMsg.AppendLine g_Language.TranslateMessage("New version: %1", newVersion)
        uiMsg.AppendLineBreak
    End If
    
    uiMsg.AppendLine g_Language.TranslateMessage("Would you like PhotoDemon to update this plugin for you?")
    
    OfferPluginUpdate = PDMsgBox(uiMsg.ToString, vbYesNoCancel Or vbApplicationModal Or vbInformation, "Update available")
    
End Function

'Attempt to update a plugin by downloading a .pdz file (from GitHub) and auto-extracting its contents
' to the local PD plugin folder.
Public Function DownloadPluginUpdate(ByVal pluginID As PD_PluginCore, ByRef srcURL As String, Optional ByVal numFilesExpected As Long = 0, Optional ByVal numBytesExpected As Long = 0) As Boolean
    
    Const FUNC_NAME As String = "DownloadPluginUpdate"
    
    Dim dstFileTemp As String
    
    'Before downloading anything, ensure we have write access on the plugin folder.
    dstFileTemp = PluginManager.GetPluginPath()
    If Not Files.PathExists(dstFileTemp, True) Then
        PDMsgBox g_Language.TranslateMessage("You have placed PhotoDemon in a restricted system folder.  Because PhotoDemon does not have administrator access, it cannot download files for you.  Please move PhotoDemon to an unrestricted folder and try again."), vbOKOnly Or vbApplicationModal Or vbCritical, g_Language.TranslateMessage("Error")
        DownloadPluginUpdate = False
        Exit Function
    End If
    
    PDDebug.LogAction "Attempting to update " & PluginManager.GetPluginName(pluginID) & "..."
    
    'Previously, PhotoDemon downloaded each plugin file as-is.  Now we package them into a single pdPackage file
    ' and extract them post-download.  (This cuts download size significantly.)
    
    'Generate a temporary filename based on the plugin being downloaded
    dstFileTemp = PluginManager.GetPluginPath() & PluginManager.GetPluginName(pluginID) & ".tmp"
    
    'If the destination temp file exists, kill it
    ' (This condition is unexpected, but maybe a previous update attempt failed?)
    Files.FileDeleteIfExists dstFileTemp
    
    'Download the user's source URL to a system-generated temp file, then copy it into our local PD folder
    Dim tmpFile As String
    tmpFile = Web.DownloadURLToTempFile(srcURL, False)
    
    If Files.FileExists(tmpFile) Then Files.FileCopyW tmpFile, dstFileTemp
    Files.FileDeleteIfExists tmpFile
    
    'With the pdPackage file successfully downloaded, extract all files to the plugins folder.
    PDDebug.LogAction "Extracting update for " & PluginManager.GetPluginName(pluginID) & "..."
    Dim cPackage As pdPackageChunky
    Set cPackage = New pdPackageChunky
    
    Dim dstFilename As String
    Dim tmpStream As pdStream, tmpChunkName As String, tmpChunkSize As Long
    
    Dim numSuccessfulFiles As Long, numBytesExtracted As Long
    numSuccessfulFiles = 0
    numBytesExtracted = 0
    
    'Load the file into a temporary package manager
    If cPackage.OpenPackage_File(dstFileTemp) Then
        
        'I use a custom-built tool to assemble pdPackage files; individual files are stored as simple name-value pairs
        Do While cPackage.GetNextChunk(tmpChunkName, tmpChunkSize, tmpStream)
            
            'Ensure the chunk name is actually a "NAME" chunk
            If (tmpChunkName = "NAME") Then
                
                'Convert the filename to a full path into the user's plugin folder
                dstFilename = PluginManager.GetPluginPath() & tmpStream.ReadString_UTF8(tmpChunkSize)
                
                'Next, extract the chunk's data
                If cPackage.GetNextChunk(tmpChunkName, tmpChunkSize, tmpStream) Then
                    
                    'Ensure the chunk data is a "DATA" chunk
                    If (tmpChunkName = "DATA") Then
                        
                        'Write the chunk's contents to file
                        If Files.FileCreateFromPtr(tmpStream.Peek_PointerOnly(0, tmpChunkSize), tmpChunkSize, dstFilename, True) Then
                            numSuccessfulFiles = numSuccessfulFiles + 1
                            numBytesExtracted = numBytesExtracted + tmpChunkSize
                        Else
                            InternalDebugMsg "failed to create target file " & dstFilename, FUNC_NAME
                        End If
                    
                    '/Validate DATA chunk
                    End If
                        
                '/Unexpected chunk
                Else
                    InternalDebugMsg "bad data chunk: " & tmpChunkName, FUNC_NAME
                End If
            
            '/Unexpected chunk
            Else
                InternalDebugMsg "bad name chunk: " & tmpChunkName, FUNC_NAME
            End If
        
        'Iterate all remaining package items
        Loop
        
    Else
        InternalDebugMsg "download failed!  " & PluginManager.GetPluginName(pluginID) & " will *not* be available to this PhotoDemon instance.", FUNC_NAME
    End If
    
    'Free the underlying package object
    Set cPackage = Nothing
    
    'Double-check expected number of files and total size of extracted bytes.
    If (numSuccessfulFiles <> numFilesExpected) Then InternalDebugMsg "unexpected extraction file count: " & numSuccessfulFiles, FUNC_NAME
    If (numBytesExtracted = numBytesExpected) Then
        PDDebug.LogAction "Successfully extracted " & numSuccessfulFiles & " files totaling " & numBytesExtracted & " bytes."
    Else
        InternalDebugMsg "unexpected extraction size: " & numBytesExtracted & " vs " & numBytesExpected, FUNC_NAME
    End If
    
    'Delete the temporary package file
    Files.FileDeleteIfExists dstFileTemp
    
    'Attempt to initialize both the import and export plugins, and return whatever PD's central plugin manager
    ' says is the state of these libraries (it may perform multiple initialization steps, including testing OS compatibility)
    PluginManager.LoadPluginGroup False
    DownloadPluginUpdate = PluginManager.IsPluginCurrentlyEnabled(pluginID)
    
End Function

'Updates involve the Internet, so any number of things can go wrong.  PD versions (post-7.0) post a lot
' of internal debug data throughout this module, to help us identify problems where we can.
Private Sub InternalDebugMsg(ByRef srcMsg As String, ByRef srcFunctionName As String, Optional ByVal errNumber As Long = 0, Optional ByRef errDescription As String = vbNullString)
    If (errNumber <> 0) Then
        PDDebug.LogAction "WARNING!  Updates." & srcFunctionName & " reported an error (#" & CStr(errNumber) & "): " & errDescription & ".  Further details: " & srcMsg
    Else
        PDDebug.LogAction "Updates." & srcFunctionName & " reports: " & srcMsg, , True
    End If
End Sub
