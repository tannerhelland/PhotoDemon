Attribute VB_Name = "Autosaves"
'***************************************************************************
'Image Autosave Handler
'Copyright 2014-2026 by Tanner Helland
'Created: 18/January/14
'Last updated: 18/October/21
'Last update: use autosave engine to (silently) restore previous sessions after forced system reboot
'
'PhotoDemon's Autosave engine is closely tied to the pdUndo class, so some understanding of that class is necessary
' to appreciate how this module operates.
'
'All Undo/Redo data is saved to the hard drive, in a temp folder of the user's choosing (the Windows temp folder
' by default).  The data is cleared whenever an image is unloaded, and an extra pass is made at program shutdown
' "just to be safe".
'
'In the event of an unclean shutdown, this module searches the temp folder for any PhotoDemon-specific data.  If
' some is found, the user is given a choice to restore those files.  If the user declines, that data is wiped
' (to prevent future unclean shutdown checks from re-detecting it).
'
'As part of its Autosave functionality, this module also handles the creation and subsequent destruction of a
' "clean shutdown" file.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'A collection of valid Autosave XML entries in the user's Data\Autosave folder.  In all but the worst-case
' scenarios (e.g. program failure during generating Undo/Redo data), these *should* correspond to raw image
' data in the Undo/Redo list.
Public Type AutosaveXML
    xmlPath As String
    parentImageID As String
    friendlyName As String
    currentFormat As PD_IMAGE_FORMAT
    originalFormat As PD_IMAGE_FORMAT
    originalPath As String
    originalSessionID As String
    undoStackHeight As Long
    undoStackAbsoluteMaximum As Long
    undoStackPointer As Long
    undoNumAtLastSave As Long
End Type

'Collection of Autosave XML entries found
Private m_numOfXMLFound As Long
Private m_XmlEntries() As AutosaveXML

'If a function wants to quickly check for previous unclean shutdowns, but *not* generate a new safe shutdown file, use this
' function instead of WasLastShutdownClean(), below.  Note that this function should only be used during Loading stages,
' because once PD has been loaded, the function will no longer be accurate.
Public Function PeekLastShutdownClean() As Boolean

    Dim safeShutdownPath As String
    safeShutdownPath = UserPrefs.GetPresetPath & "SafeShutdown.xml"
    
    'If a previous program session terminated unexpectedly, its safe shutdown file will still be present
    PeekLastShutdownClean = (Not Files.FileExists(safeShutdownPath))

End Function

'Check to make sure the last program shutdown was clean.  If it was, return TRUE (and write out a new safe shutdown file).
' If it was not, return FALSE.
Public Function WasLastShutdownClean() As Boolean

    Dim safeShutdownPath As String
    safeShutdownPath = UserPrefs.GetPresetPath & "SafeShutdown.xml"
    
    'If a previous program session terminated unexpectedly, its safe shutdown file will still be present
    If Files.FileExists(safeShutdownPath) Then
        WasLastShutdownClean = False

    'The previous shutdown was clean.  Write a new safe shutdown file.
    Else
    
        Dim xmlEngine As pdXML
        Set xmlEngine = New pdXML
        
        xmlEngine.PrepareNewXML "Safe shutdown"
        
        xmlEngine.WriteBlankLine
        xmlEngine.WriteComment "This file is used to detect unsafe shutdowns from previous PhotoDemon sessions."
        xmlEngine.WriteBlankLine
        xmlEngine.WriteTag "SessionDate", Format$(Now, "Long Date")
        xmlEngine.WriteTag "SessionTime", Format$(Now, "h:mm AMPM")
        xmlEngine.WriteTag "SessionID", OS.UniqueSessionID()
        xmlEngine.WriteBlankLine
        
        xmlEngine.WriteXMLToFile safeShutdownPath
        
        WasLastShutdownClean = True
    
    End If
    
    
End Function

'If the program has shut itself down without incident, the last thing it does will be notifying this sub.
' (This sub clears the safe shutdown file.)
Public Sub NotifyCleanShutdown()
    
    Dim safeShutdownPath As String
    safeShutdownPath = UserPrefs.GetPresetPath & "SafeShutdown.xml"
    
    Files.FileDeleteIfExists safeShutdownPath

End Sub

'During program initialization, FormMain will call this sub to handle any startup behavior related to old AutoSave data.
Public Sub InitializeAutosave()

    'DO NOT CHECK FOR AUTOSAVE DATA if another PhotoDemon session is active.
    If Mutex.IsThisOnlyInstance() Then
        
        'See if the previous PD session crashed.
        Dim shutDownClean As Boolean
        shutDownClean = Autosaves.WasLastShutdownClean
        
        'Notify the debugger; it may use this information to generate additional debug data
        PDDebug.NotifyLastSessionState shutDownClean
        
        'If our last shutdown was clean, skip further processing
        If (Not shutDownClean) Then
            
            'Oh no!  Something went horribly wrong with the last PD session.
            PDDebug.LogAction "WARNING!  Previous shutdown was *not* clean (autosave data found)."
            
            'See if there's any image autosave data worth recovering.
            If (Autosaves.SaveableImagesPresent > 0) Then
                
                PDDebug.LogAction "Saveable images found.  Preparing to restore..."
                
                'Autosave data was found!  Present it to the user.
                Dim userWantsAutosaves As VbMsgBoxResult
                Dim listOfFilesToSave() As AutosaveXML
                
                'If this is an auto-recovery session after a forced system reboot, and the user allows
                ' us to auto-recover their previous session, skip displaying any UI and just jump
                ' right to restoration.
                Dim cmdParams As pdStringStack, showUI As Boolean
                showUI = True
                
                If OS.CommandW(cmdParams, True) Then
                    
                    Dim i As Long
                    For i = 0 To cmdParams.GetNumOfStrings - 1
                        If Strings.StringsEqual(cmdParams.GetString(i), "/system-reboot", True) Then
                            showUI = False
                            Exit For
                        End If
                    Next i
                    
                    'If the autosave data comes from a perfectly good session interrupted by a forced system reboot,
                    ' restore all work silently.
                    If (Not showUI) Then
                        PDDebug.LogAction "System reboot terminated the previous session.  Restoring it now..."
                        Dim tmpXMLCount As Long
                        Autosaves.GetXMLAutosaveEntries listOfFilesToSave, tmpXMLCount
                        userWantsAutosaves = vbYes
                    End If
                    
                End If
                
                If showUI Then userWantsAutosaves = DisplayAutosaveWarning(listOfFilesToSave)
                
                'If the user wants to restore old Autosave data, do so now.
                If (userWantsAutosaves = vbYes) Then
                
                    'listOfFilesToSave contains the list of Autosave files the user wants restored.
                    ' Hand them off to the autosave handler, which will load and restore each file in turn.
                    Autosaves.LoadTheseAutosaveFiles listOfFilesToSave
                    Interface.SyncInterfaceToCurrentImage
                                
                Else
                    
                    'The user has no interest in recovering AutoSave data.  Purge all the entries we found,
                    ' so they don't show up in future AutoSave checks.
                    Autosaves.PurgeOldAutosaveData
                
                End If
                
            
            'There's not any AutoSave data worth recovering.  This is okay, as it means the unsafe shutdown
            ' occurred without any images being loaded.  Do nothing.
            Else
                PDDebug.LogAction "No saveable images found.  Skipping session restoration."
            End If
        
        Else
            PDDebug.LogAction "Previous shutdown was clean (no autosave data found)."
        End If
        
    Else
        PDDebug.LogAction "Multiple PhotoDemon sessions active; autosave check abandoned."
    End If
    
End Sub

'After an unclean shutdown is detected, this function can be called to search the temp directory for saveable Undo/Redo data.
' It will return a value larger than 0 if Undo/Redo data was found.
Public Function SaveableImagesPresent() As Long

    'Search the temporary folder for any files matching PhotoDemon's Undo/Redo file pattern.  Because PD's Undo/Redo engine
    ' is awesome, it automatically saves very nice Undo XML files that contain key data for each pdImage opened by the program.
    ' In the event of an unsafe shutdown, these XML files help us easily reconstruct any "lost" images.
    m_numOfXMLFound = 0
    ReDim m_XmlEntries(0 To 3) As AutosaveXML
    
    'We'll use PD's standard XML engine to validate any discovered autosave entries
    Dim xmlEngine As pdXML
    Set xmlEngine = New pdXML
    
    'Retrieve the first matching file from the folder (if any)
    Dim chkFile As String
    Dim listOfFiles As pdStringStack
    
    If Files.RetrieveAllFiles(UserPrefs.GetTempPath & "~PDU_StackSummary_*.pdtmp", listOfFiles, False, False) Then
    
        'Continue checking potential autosave XML entries until all have been analyzed
        Do While listOfFiles.PopString(chkFile)
            
            'First, make sure the file actually contains XML data
            If xmlEngine.LoadXMLFile(chkFile) Then
            
                'If it does, make sure the XML data is valid, and that at least one Undo entry is listed in the file
                If xmlEngine.IsPDDataType("Undo stack") And xmlEngine.ValidateLoadedXMLData("pdUndoVersion") Then
                
                    'The file checks out!  Add it to our XML entries array
                    With m_XmlEntries(m_numOfXMLFound)
                        .xmlPath = chkFile
                        .friendlyName = xmlEngine.GetUniqueTag_String("friendlyName")
                        .originalPath = xmlEngine.GetUniqueTag_String("originalPath")
                        .originalFormat = xmlEngine.GetUniqueTag_Long("originalFormat", PDIF_PDI)
                        .currentFormat = xmlEngine.GetUniqueTag_Long("currentFormat", PDIF_PDI)
                        .parentImageID = xmlEngine.GetUniqueTag_String("imageID")
                        .undoNumAtLastSave = xmlEngine.GetUniqueTag_Long("UndoNumAtLastSave", 0)
                        .undoStackAbsoluteMaximum = xmlEngine.GetUniqueTag_Long("StackAbsoluteMaximum", 0)
                        .undoStackHeight = xmlEngine.GetUniqueTag_Long("StackHeight", 1)
                        .undoStackPointer = xmlEngine.GetUniqueTag_Long("CurrentStackPointer", 0)
                    End With
                    
                    'Increment the "number found" counter and resize the array as necessary
                    m_numOfXMLFound = m_numOfXMLFound + 1
                    If (m_numOfXMLFound > UBound(m_XmlEntries)) Then ReDim Preserve m_XmlEntries(0 To (UBound(m_XmlEntries) + 1) * 2) As AutosaveXML
                    
                End If
                
            End If
            
        Loop
        
    End If
    
    'Trim the XML array to its smallest relevant size
    If (m_numOfXMLFound > 0) Then
        ReDim Preserve m_XmlEntries(0 To m_numOfXMLFound - 1) As AutosaveXML
    Else
        ReDim m_XmlEntries(0) As AutosaveXML
    End If
    
    'Return the number of images found
    SaveableImagesPresent = m_numOfXMLFound

End Function

'If the user declines to restore old AutoSave data, purge it from the system (to prevent it from showing up in future searches).
Public Sub PurgeOldAutosaveData()
    
    If (m_numOfXMLFound > 0) Then
    
        'Create a dummy pdUndo object.  This object will help us generate relevant filenames using PD's standard Undo filename formula.
        Dim tmpUndoEngine As pdUndo
        Set tmpUndoEngine = New pdUndo
        
        Dim tmpFilename As String
        Dim i As Long, j As Long
        
        'Loop through all XML files found.  We will not only be deleting the XML files themselves, but also any child
        ' files they may reference
        For i = 0 To m_numOfXMLFound - 1
        
            'Delete all possible child references for this image.
            For j = 0 To m_XmlEntries(i).undoStackAbsoluteMaximum
                
                tmpFilename = tmpUndoEngine.GenerateUndoFilenameExternal(m_XmlEntries(i).parentImageID, j)
            
                'Check image data first...
                Files.FileDeleteIfExists tmpFilename
            
                '...followed by layer data
                Files.FileDeleteIfExists tmpFilename & ".layer"
            
                '...followed by selection data
                Files.FileDeleteIfExists tmpFilename & ".selection"
            
            Next j
            
            'Finally, kill the Autosave XML file and preview image associated with this entry
            Files.FileDeleteIfExists m_XmlEntries(i).xmlPath
            Files.FileDeleteIfExists m_XmlEntries(i).xmlPath & ".pdasi"
        
        Next i
        
    End If
    
    'As a nice gesture, release any module-level data associated with the Autosave engine
    m_numOfXMLFound = 0
    ReDim m_XmlEntries(0) As AutosaveXML
    
End Sub

'External functions can retrieve a copy of the XML autosave entries we've found by using this function.
Public Function GetXMLAutosaveEntries(ByRef autosaveArray() As AutosaveXML, ByRef autosaveCount As Long) As Boolean

    ReDim autosaveArray(0 To m_numOfXMLFound - 1) As AutosaveXML
    autosaveCount = m_numOfXMLFound
    
    Dim i As Long
    For i = 0 To autosaveCount - 1
        autosaveArray(i) = m_XmlEntries(i)
    Next i
    
    GetXMLAutosaveEntries = True
    
End Function

'After any autosave images have been loaded into PD, call this function to replace those images' data (such as "location on disk")
' with information from the Autosave XML files.
Public Sub AlignLoadedImageWithAutosave(ByRef srcPDImage As pdImage)

    Dim i As Long
    
    If (Not srcPDImage Is Nothing) Then
        If srcPDImage.IsActive Then
        
            'Find a corresponding Autosave XML file for this image (if one exists)
            For i = 0 To m_numOfXMLFound - 1
            
                'If this file's location on disk matches the binary buffer associated with a given XML entry,
                ' ask the pdImage object to rewrite its internal data to match the XML file.
                If Strings.StringsEqual(srcPDImage.ImgStorage.GetEntry_String("CurrentLocationOnDisk"), m_XmlEntries(i).xmlPath, True) Then
                    
                    'Load the XML data, then attempt to initialize a pdImage object from it
                    Dim srcString As String
                    If Files.FileLoadAsString(m_XmlEntries(i).xmlPath, srcString) Then
                        If (Not srcPDImage.SetHeaderFromXML(srcString)) Then PDDebug.LogAction "WARNING!  Autosaves.AlignLoadedImageWithAutoSave failed to created a valid pdImage header."
                    Else
                        PDDebug.LogAction "WARNING!  Couldn't load autosave data: " & m_XmlEntries(i).xmlPath
                    End If
                    
                    Exit For
                    
                End If
            
            Next i
        
        End If
    End If
    
End Sub

'If the user opts to restore one (or more) autosave entries, PD's main form will pass the list of XML files
' to this function.  It is our job to then load those files.
Public Sub LoadTheseAutosaveFiles(ByRef fullXMLList() As AutosaveXML)

    Dim i As Long, autosaveFile As String
    
    'Before starting our processing loop, create a dummy pdUndo object.  This object will help us generate
    ' relevant filenames using PD's standard Undo filename formula.
    Dim tmpUndoEngine As pdUndo
    Set tmpUndoEngine = New pdUndo
    
    'An XML engine will be used to update each image's new Undo/Redo engine so that it exactly matches the
    ' state of its original Undo/Redo engine.
    Dim xmlEngine As pdXML
    Set xmlEngine = New pdXML
    
    'Process each XML entry in turn.  Because of the way we are reconstructing the Undo entries, we can't load
    ' all the files in a single request (despite PD's load function supporting a stack of filenames).
    
    'Instead, we must load each image individually, do a bunch of post-processing to the image (and its Undo files)
    ' to restore it to its proper state, *then* move on to the next image.
    For i = 0 To UBound(fullXMLList)
        
        'Make a copy of the current Undo XML file for this image, as it will be overwritten as soon as we load the first
        ' Undo entry as a new image.
        xmlEngine.LoadXMLFile fullXMLList(i).xmlPath
        
        'We now have everything we need.  Load the base Undo entry as a new image.
        autosaveFile = tmpUndoEngine.GenerateUndoFilenameExternal(fullXMLList(i).parentImageID, 0)
        Loading.LoadFileAsNewImage autosaveFile, fullXMLList(i).friendlyName, False
        
        'It is possible, but extraordinarily rare, for the LoadFileAsNewImage function to fail (for example, if the user removed
        ' a portable drive containing Autosave data in the midst of the load).  We can identify a fail state by the expected pdImage
        ' object being freed prematurely.
        If PDImages.IsImageNonNull() Then
            
            'Restore the file's original Unique ID.  (Note that this will erase the Undo/Redo data
            ' automatically created by the load function, above - that's by design!)
            PDImages.GetActiveImage.SetUniqueID fullXMLList(i).parentImageID
            
            'The new image has been successfully noted, but we must now overwrite some of the data PD has assigned it with
            ' its original data (such as its "location on disk", which should reflect its original location - not its
            ' temporary file location!)
            PDImages.GetActiveImage.ImgStorage.AddEntry "CurrentLocationOnDisk", vbNullString
            PDImages.GetActiveImage.ImgStorage.AddEntry "OriginalFileName", Files.FileGetName(fullXMLList(i).friendlyName, True)
            PDImages.GetActiveImage.ImgStorage.AddEntry "OriginalFileExtension", Files.FileGetExtension(fullXMLList(i).friendlyName)
            
            'Attempt to set its filetype.  (We rely on the file extension for this.)
            PDImages.GetActiveImage.SetOriginalFileFormat fullXMLList(i).originalFormat
            PDImages.GetActiveImage.SetCurrentFileFormat fullXMLList(i).currentFormat
            
            'Mark the image as unsaved
            PDImages.GetActiveImage.SetSaveState False, pdSE_AnySave
            
            'Reset all save dialog flags (as they should be re-displayed after autosave recovery)
            PDImages.GetActiveImage.ImgStorage.AddEntry "hasSeenJPEGPrompt", False
            PDImages.GetActiveImage.ImgStorage.AddEntry "hasSeenJP2Prompt", False
            PDImages.GetActiveImage.ImgStorage.AddEntry "hasSeenWebPPrompt", False
            PDImages.GetActiveImage.ImgStorage.AddEntry "hasSeenJXRPrompt", False
            
            'It is now time to artificially reconstruct the image's Undo/Redo stack, using the data from the autosave file.
            ' The Undo engine itself handles this step.
            If PDImages.GetActiveImage.UndoManager.ReconstructStackFromExternalSource(xmlEngine.ReturnCurrentXMLString) Then
            
                'The Undo stack was reconstructed successfully.  Ask it to advance the stack pointer to its location from
                ' the last session.
                PDImages.GetActiveImage.UndoManager.MoveToSpecificUndoPoint fullXMLList(i).undoStackPointer
                Message "Autosave reconstruction complete for %1", fullXMLList(i).friendlyName
            
            Else
                Message "Autosave could not be fully reconstructed.  Partial reconstruction attempted instead."
            End If
            
        End If
    
    Next i
    
End Sub
