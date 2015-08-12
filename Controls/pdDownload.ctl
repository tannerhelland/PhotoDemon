VERSION 5.00
Begin VB.UserControl pdDownload 
   BackColor       =   &H8000000D&
   ClientHeight    =   480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   480
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000D&
   HasDC           =   0   'False
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   32
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   32
   ToolboxBitmap   =   "pdDownload.ctx":0000
   Windowless      =   -1  'True
   Begin VB.Timer tmrReset 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "pdDownload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Asynchronous Download control
'Copyright 2014-2015 by Tanner Helland
'Created: 24/January/15
'Last updated: 26/January/15
'Last update: wrapped up initial build
'
'In keeping with PD's grand tradition of NIH syndrome, this custom user control was built to facilitate asynchronous
' downloads of various PD elements.  It is rather tightly integrated with PD itself, relying on things like publicly
' available functions, so as usual, you'll need to do some editing before sticking it in another project.
'
'For a better standalone project, here are some excellent choices:
'
'http://www.vbforums.com/showthread.php?733409-VB6-Simple-Async-Download-Ctl-for-multiple-Files
'http://visualstudiomagazine.com/articles/2008/03/27/simple-asynchronous-downloads.aspx
'https://github.com/Kroc/blu/blob/master/bluDownload.ctl
'
'Many thanks in particular to that last link; Kroc Camen's bluDownload control served PD very well prior to pdDownload,
' and it would be a great choice if you need a simple, standalone single-file download UC.
'
'The goal with this control is simple: to provide silent background downloads of relevant PD files.  At present, it is
' primarily focused on update files (including language updates), but in the future it will likely be expanded to cover
' other items, including patches for PD itself.
'
'Still TODO:
' - Implement a "try again later" flag.  This would retry any failed downloads within [x] seconds, using a timer to track [x].
'    This could be helpful for weird intermittent Internet issues, when downloading non-mission-critical items.
' - Raise progress events, as necessary.  I've deliberately avoided these at present, because they clutter up the code but
'    aren't particularly useful in PD, as the asynchronicity means we aren't bothering the user with progress updates.
'    If these ever prove helpful, I'll drop 'em in.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Two "download successful" events are raised: one for each individual file, as they complete, and a separate one for when every
' download in the queue completes.  Note that the "every download complete" event may not be accurate if you are starting
' downloads willy-nilly; instead, if you plan on using that event, I suggest not starting the download process until ALL RELEVANT
' URLS have been added to the queue.  (If you don't do this, items may download before subsequent ones are added to the queue,
' so the FinishedAllItems event may raise multiple times.)
'
'The FinishedOneItem() event provides a pointer to the downloaded data.  You can retrieve it immediately, or wait until later
' as this class will automatically make an internal copy of the data.
'
'The FinishedAllItems() event cannot easily return a pointer for all data for all items, so instead, the caller must retrieve
' relevant keys manually.
'
'TODO: pass a ByRef "try again" flag, so the recipient can easily re-initate downloads that weren't successful
Event FinishedOneItem(ByVal downloadSuccessful As Boolean, ByVal entryKey As String, ByVal OptionalType As Long, ByRef downloadedData() As Byte, ByVal savedToThisFile As String)
Event FinishedAllItems(ByVal allDownloadsSuccessful As Boolean)

'Because we often have to download multiple files at once, a custom type is used.  This type includes enums
' relevant to each currently downloading file.
Public Enum pdDownloadStatus
    PDS_NOT_YET_STARTED = 0
    PDS_DOWNLOADING = 1
    PDS_DOWNLOAD_COMPLETE = 2
    PDS_FAILURE_CALLER_CANCELED_DOWNLOAD = 3
    PDS_FAILURE_BUT_WILL_TRY_AGAIN_SOON = 4
    PDS_FAILURE_NOT_TRYING_AGAIN = 5
    PDS_FAILURE_CHECKSUM_MISMATCH = 6
End Enum

#If False Then
    Private Const PDS_NOT_YET_STARTED = 0, PDS_DOWNLOADING = 1, PDS_DOWNLOAD_COMPLETE = 2
    Private Const PDS_FAILURE_CALLER_CANCELED_DOWNLOAD = 3, PDS_FAILURE_BUT_WILL_TRY_AGAIN_SOON = 4, PDS_FAILURE_NOT_TRYING_AGAIN = 5
    Private Const PDS_FAILURE_CHECKSUM_MISMATCH = 6
#End If

Private Type pdDownloadEntry
    Key As String
    DownloadTypeOptional As Long
    CurrentStatus As pdDownloadStatus
    DownloadURL As String
    DownloadFlags As AsyncReadConstants
    TargetFileWhenComplete As String
    BytesDownloaded As Long
    BytesTotal As Long
    LastAsyncStatusCode As Long
    LastAsyncStatus As Variant      'This may be a string, but VB doesn't guarantee the type.  The return varies according
                                    ' to the last status code.
    DataBytes() As Byte
    DataBytesMarkedForRelease As Boolean
    ExpectedChecksum As Long        'Callers can supply the downloader with an expected checksum.  If present, pdDownload will automatically verify
                                    ' the download's integrity, and raise an error state if the download completes but the checksum doesn't mathc.
End Type


'The number of files currently being downloaded is important, so we can raise a special "all files done" notification.
' The user may choose to respond to this instead of a single file download.
Private m_NumOfFiles As Long

'This value tracks the number of files whose downloading has ceased, whether by error or successful completion.  It is used
' to raise the FinishedAllItems() event.
Private m_NumOfFilesFinishedDownloading As Long

'This array actually stores the list of currently downloading files.
Private m_DownloadList() As pdDownloadEntry

'To keep things simple, pdDownload uses a single "download/don't download right now" flag.  If TRUE, downloads for any
' files in the queue can proceed.  If FALSE, no downloading will occur.  PD uses this flag to postpone downloading until
' all desired downloads have been added to the queue.  (So we can deal with them in a single batch, by tracking the
' "all downloads complete" event.)
Private m_DownloadsAllowed As Boolean

'If an error occurs, it will be tracked here.  In PD, asynchronous downloads aren't generally used for mission-critical items,
' so errors just mean "try again later".  However, any returned error codes are stored here if you want 'em.
Private m_LastErrorNumber As Long, m_LastErrorDescription As String

'VB's asynchronous model is pretty simple for arrays; the cLocks value in the SAFEARRAY header is incremented and decremented
' for each lock/unlock action.  Access errors are a strong potential for a class like this, where accesses and resets may happen
' inside raised events, so as a precaution, I track "Reset" instructions via this module-level boolean.  Once a reset has been
' requested, this will be set to TRUE alongside the tmrReset control.  Every second, that control will check to see if all locks
' have been released for a given array.  If they have, it knows it can safely erase the master array.
Private m_ResetActive As Boolean

Private Sub tmrReset_Timer()

    On Error GoTo arrayNotReadyForRelease

    If m_ResetActive Then
    
        'Reset the master tracking array.  Note that this may fail when in an error state, as the array will be locked
        ' when we try to ReDim it.  In that case, simply carry on.
        ReDim m_DownloadList(0 To 3) As pdDownloadEntry
        
        'If we didn't error out, the ReDim was successful.  Reset the timer and tracking variable.
        m_ResetActive = False
        tmrReset.Enabled = False
    
    End If
    
arrayNotReadyForRelease:

End Sub

'Something has finished downloading.  Update our internal tracking directory, and raise events as necessary.
' (FYI: this function is also raised when download stops due to an error state; PD will try to auto-detect this,
'       and update its raised event accordingly.)
Private Sub UserControl_AsyncReadComplete(AsyncProp As AsyncProperty)
    
    'Per MSDN, any async errors will trigger here, when accessing AsyncProp.Value; this is the only mechanism we have of reliably
    ' checking download failures.
    On Error GoTo DownloadError
    
    'Start by finding the matching internal struct
    Dim itemIndex As Long
    itemIndex = doesKeyExist(AsyncProp.PropertyName)
    
    If itemIndex >= 0 Then
    
        'Check the download state.  If the download was incomplete, we might be able to detect it here, rather than
        ' having to rely on the error handler.
        If AsyncProp.StatusCode <> vbAsyncStatusCodeEndDownloadData Then
            
            'Download failed.  Populate struct elements anyway, then raise a completion event with the FAIL flag set.
            With m_DownloadList(itemIndex)
                .CurrentStatus = PDS_FAILURE_NOT_TRYING_AGAIN
                .BytesDownloaded = AsyncProp.BytesRead
                .BytesTotal = AsyncProp.BytesMax
                .LastAsyncStatusCode = AsyncProp.StatusCode
                .LastAsyncStatus = AsyncProp.Status
                
                'This line is likely to trigger an error if the download failed, so we hit it last.
                .DataBytes = AsyncProp.Value
                
                'Raise a failure event
                RaiseEvent FinishedOneItem(False, .Key, .DownloadTypeOptional, .DataBytes, "")
                
            End With
            
            'Increment the "download finished" counter by 1.
            m_NumOfFilesFinishedDownloading = m_NumOfFilesFinishedDownloading + 1
            
            'Check to see if all downloads have completed
            checkAllDownloadsComplete
                        
        'Download appears to be successful, but we won't know for sure until we try to access the .Value property...
        Else
        
            'Start populate internal struct elements
            With m_DownloadList(itemIndex)
                .CurrentStatus = PDS_DOWNLOAD_COMPLETE
                .BytesDownloaded = AsyncProp.BytesRead
                .LastAsyncStatusCode = AsyncProp.StatusCode
                .LastAsyncStatus = AsyncProp.Status
                
                'This line is likely to trigger an error if the download failed, so we hit it last.
                .DataBytes = AsyncProp.Value
                
                'BytesMax returns 0 when downloading to a byte array - I'm not sure why.  As such, we'll calculate
                ' BytesTotal separately.
                .BytesTotal = (UBound(.DataBytes) - LBound(.DataBytes)) + 1
                
            End With
            
            'If we made it all the way here, the download appears to be successful.
            
            'If the user supplied us with a checksum, verify it now.  We will report an error state if the checksums don't match.
            Dim checkSumPassed As Boolean
            checkSumPassed = True
            
            If m_DownloadList(itemIndex).ExpectedChecksum <> 0 Then
                
                'Use pdPackager to checksum the retrieved data
                Dim cPackage As pdPackager
                Set cPackage = New pdPackager
                
                Dim chksumVerify As Long
                chksumVerify = cPackage.checkSumArbitraryArray(m_DownloadList(itemIndex).DataBytes)
                
                'Check for equality; the downloader will report failure if the checksums do not match
                checkSumPassed = (chksumVerify = m_DownloadList(itemIndex).ExpectedChecksum)
                
            End If
            
            'If requested, copy the contents out to file.
            With m_DownloadList(itemIndex)
            
                If (Len(.TargetFileWhenComplete) > 0) Then
                    
                    'Make sure the checksum passed (if one was specified).
                    If checkSumPassed Then
                        
                        'All file interactions are handled through pdFSO
                        Dim cFile As pdFSO
                        Set cFile = New pdFSO
                        
                        'Kill the destination file if it already exists
                        If cFile.FileExist(.TargetFileWhenComplete) Then cFile.KillFile .TargetFileWhenComplete
                        
                        'Dump the downloaded data to file
                        Dim hFile As Long
                        If cFile.CreateFileHandle(.TargetFileWhenComplete, hFile, True, True, OptimizeSequentialAccess) Then
                            
                            cFile.WriteDataToFile hFile, VarPtr(.DataBytes(0)), UBound(.DataBytes) + 1
                            cFile.CloseFileHandle hFile
                            
                        Else
                            Debug.Print "WARNING! File was downloaded successfully, but we couldn't write it to the hard drive.  Check the path: " & .TargetFileWhenComplete
                        End If
                                            
                    Else
                        Debug.Print "WARNING! File was downloaded successfully, but checksum failed.  Please investigate: " & m_DownloadList(itemIndex).DownloadURL
                    End If
                    
                End If
            
            End With
            
            'Raise a success/failure event, based on the checksum result (if any; note that checkSumPassed defaults to TRUE)
            With m_DownloadList(itemIndex)
                If checkSumPassed Then
                    RaiseEvent FinishedOneItem(True, .Key, .DownloadTypeOptional, .DataBytes, .TargetFileWhenComplete)
                Else
                    RaiseEvent FinishedOneItem(False, .Key, .DownloadTypeOptional, .DataBytes, .TargetFileWhenComplete)
                End If
            End With
            
            'Increment the "download finished" counter by 1.
            m_NumOfFilesFinishedDownloading = m_NumOfFilesFinishedDownloading + 1
            
            'Check to see if all downloads have completed
            checkAllDownloadsComplete
        
        End If
    
    'It shouldn't technically be possible for this function to return a key that doesn't exist, but better safe than sorry...
    Else
        Debug.Print "WARNING! AsyncReadComplete in pdDownload returned an invalid key.  No event will be raised."
        Exit Sub
    End If
    
    Exit Sub
    
'If something went horribly wrong during the download process, this chunk of code will be triggered
DownloadError:
    
    Debug.Print "WARNING!  An error occurred in pdDownload's AsyncReadComplete event.  Download abandoned."
    
    'Download failed.  Populate struct elements anyway, then raise a completion event with the FAIL flag set.
    With m_DownloadList(itemIndex)
        .CurrentStatus = PDS_FAILURE_NOT_TRYING_AGAIN
        .LastAsyncStatusCode = AsyncProp.StatusCode
        .LastAsyncStatus = AsyncProp.Status
        Erase .DataBytes
    End With
    
    'Raise a failure event
    RaiseEvent FinishedOneItem(False, m_DownloadList(itemIndex).Key, m_DownloadList(itemIndex).DownloadTypeOptional, m_DownloadList(itemIndex).DataBytes, "")
    
    'Increment the "download finished" counter by 1.
    m_NumOfFilesFinishedDownloading = m_NumOfFilesFinishedDownloading + 1
    
    'Check to see if all downloads have completed
    checkAllDownloadsComplete
    
End Sub

'Call this to see if all downloads are complete.  If they are, a "FinishedAllItems" event will be raised.
Public Sub checkAllDownloadsComplete()

    If m_NumOfFilesFinishedDownloading = m_NumOfFiles Then
    
        'All files have finished downloading.  Check to see if any failed.
        Dim allFilesSuccessful As Boolean
        allFilesSuccessful = True
        
        Dim i As Long
        For i = 0 To m_NumOfFiles - 1
            If m_DownloadList(i).CurrentStatus <> PDS_DOWNLOAD_COMPLETE Then
                allFilesSuccessful = False
                Exit For
            End If
        Next i
        
        'Raise the event; it is up to the caller to retrieve any desired data after this point.
        RaiseEvent FinishedAllItems(allFilesSuccessful)
    
    End If

End Sub

'Callers can query individual items for success/failure
Public Function wasDownloadSuccessful(ByVal itemKey As String) As Boolean
    
    Dim itemIndex As Long
    itemIndex = doesKeyExist(itemKey)
    
    If itemKey >= 0 Then
    
        'Check the download status, and make sure at least one byte was retrieved.
        With m_DownloadList(itemIndex)
        
            If (.CurrentStatus = PDS_DOWNLOAD_COMPLETE) And (UBound(.DataBytes) >= LBound(.DataBytes)) Then
                wasDownloadSuccessful = True
            Else
                wasDownloadSuccessful = False
            End If
        
        End With
    
    Else
        wasDownloadSuccessful = False
    End If

End Function

'Callers can use this to retrieve the downloaded contents of a given key.  (Note that at present, the downloaded data remains
' in memory, even if the caller requested it written out to file.)
Public Function copyDownloadArray(ByVal itemKey As String, ByRef targetBytes() As Byte) As Boolean

    Dim itemIndex As Long
    itemIndex = doesKeyExist(itemKey)
    
    If itemKey >= 0 Then
    
        'Check the download status, and make sure at least one byte was retrieved.
        With m_DownloadList(itemIndex)
        
            If (.CurrentStatus = PDS_DOWNLOAD_COMPLETE) And (UBound(.DataBytes) >= LBound(.DataBytes)) Then
                
                ReDim targetBytes(LBound(.DataBytes) To UBound(.DataBytes)) As Byte
                CopyMemory ByVal VarPtr(targetBytes(LBound(.DataBytes))), ByVal VarPtr(.DataBytes(LBound(.DataBytes))), (UBound(.DataBytes) - LBound(.DataBytes)) + 1
                copyDownloadArray = True
                
            Else
                copyDownloadArray = False
            End If
        
        End With
    
    Else
        copyDownloadArray = False
    End If

End Function

'When a caller is done with a given download item, they can call this function to release all resources associated with that item.
Public Sub freeResourcesForItem(ByVal itemKey As String)
    
    Dim itemIndex As Long
    itemIndex = doesKeyExist(itemKey)
    
    If itemIndex >= 0 Then
    
        'Check the download status, and make sure at least one byte was retrieved.
        With m_DownloadList(itemIndex)
            
            'Erase the data chunk, but leave any other indicators as things like "all downloads complete" may rely
            ' on that data.
            Erase .DataBytes
            
            'Also erase the name, so new downloads with that name can be initiated
            .Key = ""
            
        End With
        
    End If
    
End Sub

'Use this sub to reset everything to its virgin state.
Public Sub Reset(Optional ByVal setFailsafeTimer As Boolean = True)
    
    'Cancel any downloads currently in progress
    Dim i As Long
    If m_NumOfFiles > 0 Then
    
        For i = 0 To m_NumOfFiles - 1
        
            With m_DownloadList(i)
            
                If .CurrentStatus = PDS_DOWNLOADING Then
                    UserControl.CancelAsyncRead .Key
                    .CurrentStatus = PDS_FAILURE_CALLER_CANCELED_DOWNLOAD
                End If
            
            End With
        
        Next i
    
    End If
    
    'Reset all tracking variables
    m_NumOfFiles = 0
    m_NumOfFilesFinishedDownloading = 0
    m_DownloadsAllowed = False
    ReDim m_DownloadList(0 To 31) As pdDownloadEntry
    
    'The master tracking array is likely locked, as this function will likely be accessed from inside a raised event.
    ' To prevent asynchronicity issues, launch a separate timer.  It will handle the actual erasing of the array.
    If setFailsafeTimer Then
        m_ResetActive = True
        tmrReset.Enabled = True
    End If

End Sub

Private Sub UserControl_Initialize()
    
    'Reset everything to its default state
    Reset False
    
End Sub

'At termination, all downloads are forcibly stopped and any existing data is deleted.
' This all happens automatically, so we don't need to do our own clean-up.
Private Sub UserControl_Terminate()
    Reset False
End Sub

'Add a file to the queue.  Note that this DOES NOT start the download, unless setDownloadState has been passed TRUE at some
' prior point, or the startDownloadingNow flag is passed.
'
'FYI: for privacy and portability reasons, pdDownload does not download to a temp file.  It downloads directly to a byte array.
' (This is a good solution for the small files PD deals with, as well as portable users who don't want programs auto-downloading
'  stuff to a foregin "Temporary Internet Files" folder.  That said, this is obviously impractical for those who want to download
' very large files - consider yourself warned.)
'
'Inputs:
' 1) Download Item Key.  This key MUST BE UNIQUE for each entry in the download queue.
' 2) URL to download.  Standard VB download rules apply (e.g. https:// will cause issues)
' 3) Optional download flags.  By default, PD is set to download a new copy of the file only if the server version is newer
'    (vbAsyncReadResynchronize), but downloads can also be forcibly requested regardless of version (vbAsyncReadForceUpdate)
' 4) Optional startDownloadImmediately parameter, which does exactly what you think it does.  NOTE: if set, this flag operates
'    independent of the class-wide m_DownloadsAllowed flag.  This can make it a little unwieldy if you are mixing and matching
'    immediate download state for different files.  Plan accordingly.
' 5) Optional saveToThisFileWhenComplete parameter, which instructs pdDownload to immediately save the file contents to this file
'    when download is complete.  Note that it will do this *before* raising the download complete event, so you can make use
'    of the file immediately.
'
'Returns: success/fail.  Fail is unlikely, unless the caller does something stupid like specifying a duplicate key.
Public Function addToQueue(ByVal downloadKey As String, ByVal urlString As String, Optional ByVal OptionalDownloadType As Long = 0, Optional ByVal asyncFlags As AsyncReadConstants = vbAsyncReadResynchronize, Optional ByVal startDownloadImmediately As Boolean = False, Optional ByVal saveToThisFileWhenComplete As String = "", Optional ByVal checksumToVerify As Long = 0) As Boolean

    'Make sure this key is unique in the collection
    If doesKeyExist(downloadKey) >= 0 Then
    
        'Duplicate keys are not allowed.
        Debug.Print "WARNING: duplicate download key requested in pdDownload addToQueue.  Invalid usage; download abandoned."
        addToQueue = False
        Exit Function
    
    End If
    
    'On Error GoTo addToQueueFailure
    
    'This key is unique; add it now
    Dim itemIndex As Long
    itemIndex = m_NumOfFiles
    
    With m_DownloadList(itemIndex)
        .Key = downloadKey
        .DownloadTypeOptional = OptionalDownloadType
        .DownloadURL = urlString
        .DownloadFlags = asyncFlags
        .CurrentStatus = PDS_NOT_YET_STARTED
        .BytesDownloaded = 0
        .BytesTotal = 0
        .TargetFileWhenComplete = saveToThisFileWhenComplete
        .ExpectedChecksum = checksumToVerify
    End With
        
    'Update the size of the directory, as necessary
    m_NumOfFiles = m_NumOfFiles + 1
    If m_NumOfFiles > UBound(m_DownloadList) Then ReDim Preserve m_DownloadList(0 To m_NumOfFiles * 2 - 1) As pdDownloadEntry
    
    'If the user requested an immediate download, initiate it now and mirror that return value to addToQueue
    If m_DownloadsAllowed Or startDownloadImmediately Then
        addToQueue = startDownloadingByIndex(itemIndex)
    Else
        addToQueue = True
    End If
    
    'Success!
    addToQueue = True
    Exit Function
    
addToQueueFailure:

    addToQueue = False
    m_LastErrorNumber = Err.Number
    m_LastErrorDescription = Err.Description
    
    'TODO: implement "try again later" status
    m_DownloadList(itemIndex).CurrentStatus = PDS_FAILURE_NOT_TRYING_AGAIN
    
End Function

'doesKeyExist looks for a given downloadKey (a KEY, not a URL, although they may be the same thing depending on your usage)
' in the current collection.  Like other places in PD, binary compare mode is enforced.  Plan accordingly.
'
'Returns: index of found key (>= 0) if key exists.  -1 if it does not exist.
Public Function doesKeyExist(ByVal downloadKey As String) As Long

    If m_NumOfFiles > 0 Then
        
        Dim i As Long
        For i = 0 To m_NumOfFiles - 1
            If StrComp(downloadKey, m_DownloadList(i).Key, vbBinaryCompare) = 0 Then
                doesKeyExist = i
                Exit Function
            End If
        Next i
        
        'If we made it here, the key does not exist
        
    End If
    
    doesKeyExist = -1

End Function

'pdDownload will automatically resize its download directory as files are added to it, and it starts with a default directory
' size of FOUR.  If you know that you will be downloading many files, you can set the directory size in advance; this spares
' expensive ReDim Preserve operations down the road.
Public Function forceDownloadQueueSize(ByVal newSize As Long) As Boolean

    'Perform a failsafe check against current queue contents
    If (newSize < m_NumOfFiles) Or (newSize < 0) Then
        Debug.Print "WARNING! forceDownloadQueueSize requested an invalid size; queue cannot be reduced while downloads are in progress."
        forceDownloadQueueSize = False
        Exit Function
    End If
    
    'ReDim or ReDim Preserve as necessary.  (Generally speaking, I wouldn't advise against using this function *after* files
    ' have been added to the queue, even though this function allows it.)
    If m_NumOfFiles > 0 Then
        ReDim Preserve m_DownloadList(0 To newSize - 1) As pdDownloadEntry
    Else
        ReDim m_DownloadList(0 To newSize - 1) As pdDownloadEntry
    End If

End Function

'Downloads can be started at any time.  My personal preferences is to add all required downloads to the engine, then download
' them all together (which makes it more meaningful to raise an "all files done downloading" event), but there's nothing that
' prevents individual files from starting their downloads immediately.  Call this function to set the global "download
' immediately" flag.
'
'Note that by design, this flag does not auto-unset itself when files finish downloading.  This is intentional, to cover the
' case where files download faster than PD can add subsequent ones to the queue.  To stop auto-download mode, you must
' manually pass FALSE to this function.  (Note that there is no penalty to leaving the object in auto-download mode, as it
' won't do anything if there aren't files to download.)
Public Sub setAutoDownloadMode(ByVal newMode As Boolean)
    
    m_DownloadsAllowed = newMode
    
    'If files are in the queue, start download them immediately.
    If m_DownloadsAllowed Then
    
        Dim i As Long
        For i = 0 To m_NumOfFiles - 1
            If m_DownloadList(i).CurrentStatus = PDS_NOT_YET_STARTED Then
                startDownloadingByIndex i
            End If
        Next i
    
    End If
    
End Sub

'Start downloading an individual item.  Generally speaking, this should be used internally, rather than called randomly from
' outside sources.  Returned value is TRUE if the download appears to be initiated successfully; FALSE if we couldn't start
' the download.  For detailed failure information, use the separate error retrieval functions.
Public Function startDownloadingByIndex(ByVal keyIndex As Long) As Boolean

    'This function actually makes use of error tracking, as .AsyncRead may throw errors for various Internet issues.
    On Error GoTo startDownloadingFailure
    
    If (keyIndex >= 0) And (keyIndex < m_NumOfFiles) Then
        
        'Make sure the file isn't already downloading
        If m_DownloadList(keyIndex).CurrentStatus <> PDS_DOWNLOADING Then
            
            'Everything is good to go!  Start downloading the file in question, and return SUCCESS
            With m_DownloadList(keyIndex)
                .CurrentStatus = PDS_DOWNLOADING
                UserControl.AsyncRead .DownloadURL, vbAsyncTypeByteArray, .Key, .DownloadFlags
            End With
            
            startDownloadingByIndex = True
            
        Else
            Debug.Print "WARNING! This item is already downloading!"
            startDownloadingByIndex = False
        End If
        
    Else
        Debug.Print "WARNING! Could not start download, because the specific item index is invalid."
        startDownloadingByIndex = False
    End If
    
'UserControl.AsyncRead may throw errors for various Internet issues.  Rather than raise errors, we simply return FALSE,
' and the caller can choose to retrieve more specific error information as necessary.
startDownloadingFailure:

    startDownloadingByIndex = False
    m_LastErrorNumber = Err.Number
    m_LastErrorDescription = Err.Description
    
    'TODO: implement "try again later" status
    If keyIndex >= 0 Then m_DownloadList(keyIndex).CurrentStatus = PDS_FAILURE_NOT_TRYING_AGAIN

End Function

'Thin wrapper to startDownloadingByIndex, above
Public Function startDownloadingByKey(ByVal itemKey As String) As Boolean
    
    'Retrieve an index for the specified key
    Dim keyIndex As Long
    keyIndex = doesKeyExist(itemKey)
    
    If keyIndex >= 0 Then
        startDownloadingByKey = startDownloadingByIndex(keyIndex)
    Else
        Debug.Print "WARNING! Could not start download, because itemKey does not exist in collection."
        startDownloadingByKey = False
    End If
    
End Function

'If a function tied to downloading returns a FALSE state, last error data can be retrieved here.
Public Function getLastErrorNumber() As Long
    getLastErrorNumber = m_LastErrorNumber
End Function

Public Function getLastErrorDescription() As String
    getLastErrorDescription = m_LastErrorDescription
End Function
