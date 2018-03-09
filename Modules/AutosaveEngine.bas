Attribute VB_Name = "Autosaves"
'***************************************************************************
'Image Autosave Handler
'Copyright 2014-2018 by Tanner Helland
'Created: 18/January/14
'Last updated: 09/March/18
'Last update: previous session crashes now notify pdDebug; it may choose to activate the debugger, even in stable builds,
'             as a failsafe.
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
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'A collection of valid Autosave XML entries in the user's Data\Autosave folder.  In all but the worst-case
' scenarios (e.g. program failure during generating Undo/Redo data), these *should* correspond to raw image
' data in the Undo/Redo list.
Public Type AutosaveXML
    xmlPath As String
    parentImageID As Long
    friendlyName As String
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
    If (Not App.PrevInstance) Then
        
        'See if the previous PD session crashed.
        Dim shutDownClean As Boolean
        shutDownClean = Autosaves.WasLastShutdownClean
        
        'Notify the debugger; it may use this information to generate additional debug data
        pdDebug.NotifyLastSessionState shutDownClean
        
        'If our last shutdown was clean, skip further processing
        If (Not shutDownClean) Then
            
            'Oh no!  Something went horribly wrong with the last PD session.
            pdDebug.LogAction "WARNING!  Previous shutdown was *not* clean (autosave data found)."
            
            'See if there's any image autosave data worth recovering.
            If (Autosaves.SaveableImagesPresent > 0) Then
            
                'Autosave data was found!  Present it to the user.
                Dim userWantsAutosaves As VbMsgBoxResult
                Dim listOfFilesToSave() As AutosaveXML
                
                userWantsAutosaves = DisplayAutosaveWarning(listOfFilesToSave)
                
                'If the user wants to restore old Autosave data, do so now.
                If (userWantsAutosaves = vbYes) Then
                
                    'listOfFilesToSave contains the list of Autosave files the user wants restored.
                    ' Hand them off to the autosave handler, which will load and restore each file in turn.
                    Autosaves.LoadTheseAutosaveFiles listOfFilesToSave
                    SyncInterfaceToCurrentImage
                                
                Else
                    
                    'The user has no interest in recovering AutoSave data.  Purge all the entries we found, so they don't show
                    ' up in future AutoSave searches.
                    Autosaves.PurgeOldAutosaveData
                
                End If
                
            
            'There's not any AutoSave data worth recovering.  This is okay, as it means the unsafe shutdown
            ' occurred without any images being loaded.  Do nothing.
            Else
            
            End If
        
        Else
            pdDebug.LogAction "Previous shutdown was clean (no autosave data found)."
        End If
        
    Else
        pdDebug.LogAction "Multiple PhotoDemon sessions active; autosave check abandoned."
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
    
    If Files.RetrieveAllFiles(UserPrefs.GetTempPath & "~PDU_StackSummary_*_.pdtmp", listOfFiles, False, False) Then
    
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
                    .originalSessionID = xmlEngine.GetUniqueTag_String("OriginalSessionID")
                    .parentImageID = xmlEngine.GetUniqueTag_Long("imageID", -1)
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
    
    'Sort the file array in ascending order, according to the image's original image ID values.  If the user chooses to load these
    ' autosave files, the generated pdImage objects will likely get assigned a different ID value than what they had in the previous
    ' session.  To make the existing Undo files align with the newly assigned ID value, the Undo files must be renamed (because
    ' imageID is part of each Undo file's name - that's how we track separate Undo chains for each loaded image).  The trickiness
    ' starts when a loaded autosave image is assigned a new ID value, and that ID happens to correspond to one of the ID values from
    ' the *previous* session.  When it comes time to rename the Undo files, we may inadvertently overwrite another autosave image's
    ' Undo data with the new image's data, if the new image ID matches the other image's old ID!  Needless to say, this causes all
    ' sorts of havoc.  To prevent this from ever occurring, we manually sort images by ID order, to ensure that when new ID values
    ' are assigned out, they never inadvertently overwrite another autosave image's original ID value.  (This works because ID values
    ' are assigned in ascending order, so as long as the Autosave files are also loaded in ascending order, no new image ID will
    ' ever overwrite an old image's ID.)
    If (m_numOfXMLFound > 0) Then SortAutosaveEntries
    
    'Return the number of images found
    SaveableImagesPresent = m_numOfXMLFound

End Function

'Sort the m_XmlEntries() array in ascending order, using original image ID as the sort parameter
Private Sub SortAutosaveEntries()

    Dim i As Long, j As Long
    
    'Loop through all entries in the autosave array, sorting them as we go
    For i = 0 To m_numOfXMLFound - 1
        For j = 0 To m_numOfXMLFound - 1
            If (m_XmlEntries(i).parentImageID < m_XmlEntries(j).parentImageID) Then SwapAutosaveData m_XmlEntries(i), m_XmlEntries(j)
        Next j
    Next i

End Sub

'Swap the values of two Autosave entries
Private Sub SwapAutosaveData(ByRef asOne As AutosaveXML, ByRef asTwo As AutosaveXML)
    Dim asTmp As AutosaveXML
    asTmp = asOne
    asOne = asTwo
    asTwo = asTmp
End Sub

'If the user declines to restore old AutoSave data, purge it from the system (to prevent it from showing up in future searches).
Public Sub PurgeOldAutosaveData()
    
    If (m_numOfXMLFound > 0) Then
    
        Message "Purging old autosave data..."
        
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
                
                tmpFilename = tmpUndoEngine.GenerateUndoFilenameExternal(m_XmlEntries(i).parentImageID, j, m_XmlEntries(i).originalSessionID)
            
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
                    srcPDImage.ReadExternalData m_XmlEntries(i).xmlPath
                    Exit For
                End If
            
            Next i
        
        End If
    End If
    
End Sub

'If the user opts to restore one (or more) autosave entries, PD's main form will pass the list of XML files
' to this function.  It is our job to then load those files.
Public Sub LoadTheseAutosaveFiles(ByRef fullXMLList() As AutosaveXML)

    Dim i As Long, newImageID As Long, oldImageID As Long
    Dim autosaveFile As String
    
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
    
        'Before doing anything else, we are going to rename the Undo files associated with this Autosave entry.
        ' PD assigns image IDs sequentially in each session, starting with image ID #1.  Because the image ID is immutable
        ' (it corresponds to the image's location in the master pdImages() array), we cannot simply change it to match
        ' the ID of the Undo files - instead, we must rename the Undo files to match the new image ID.
        newImageID = CanvasManager.GetProvisionalImageID()
        oldImageID = fullXMLList(i).parentImageID
        
        RenameAllUndoFiles fullXMLList(i), newImageID, oldImageID
        
        'Make a copy of the current Undo XML file for this image, as it will be overwritten as soon as we load the first
        ' Undo entry as a new image.
        xmlEngine.LoadXMLFile fullXMLList(i).xmlPath
        
        'We now have everything we need.  Load the base Undo entry as a new image.
        autosaveFile = tmpUndoEngine.GenerateUndoFilenameExternal(newImageID, 0, OS.UniqueSessionID())
        Loading.LoadFileAsNewImage autosaveFile, fullXMLList(i).friendlyName, False
        
        'It is possible, but extraordinarily rare, for the LoadFileAsNewImage function to fail (for example, if the user removed
        ' a portable drive containing Autosave data in the midst of the load).  We can identify a fail state by the expected pdImage
        ' object being freed prematurely.
        If (Not pdImages(g_CurrentImage) Is Nothing) Then
        
            'The new image has been successfully noted, but we must now overwrite some of the data PD has assigned it with
            ' its original data (such as its "location on disk", which should reflect its original location - not its
            ' temporary file location!)
            pdImages(g_CurrentImage).ImgStorage.AddEntry "CurrentLocationOnDisk", vbNullString
            pdImages(g_CurrentImage).ImgStorage.AddEntry "OriginalFileName", Files.FileGetName(fullXMLList(i).friendlyName, True)
            pdImages(g_CurrentImage).ImgStorage.AddEntry "OriginalFileExtension", Files.FileGetExtension(fullXMLList(i).friendlyName)
            
            'Mark the image as unsaved
            pdImages(g_CurrentImage).SetSaveState False, pdSE_AnySave
            
            'Reset all save dialog flags (as they should be re-displayed after autosave recovery)
            pdImages(g_CurrentImage).ImgStorage.AddEntry "hasSeenJPEGPrompt", False
            pdImages(g_CurrentImage).ImgStorage.AddEntry "hasSeenJP2Prompt", False
            pdImages(g_CurrentImage).ImgStorage.AddEntry "hasSeenWebPPrompt", False
            pdImages(g_CurrentImage).ImgStorage.AddEntry "hasSeenJXRPrompt", False
            
            'It is now time to artificially reconstruct the image's Undo/Redo stack, using the data from the autosave file.
            ' The Undo engine itself handles this step.
            If pdImages(g_CurrentImage).UndoManager.ReconstructStackFromExternalSource(xmlEngine.ReturnCurrentXMLString) Then
            
                'The Undo stack was reconstructed successfully.  Ask it to advance the stack pointer to its location from
                ' the last session.
                pdImages(g_CurrentImage).UndoManager.MoveToSpecificUndoPoint fullXMLList(i).undoStackPointer
                Message "Autosave reconstruction complete for %1", fullXMLList(i).friendlyName
            
            Else
                Message "Autosave could not be fully reconstructed.  Partial reconstruction attempted instead."
            End If
            
        End If
    
    Next i
    
End Sub

'loadTheseAutosaveFiles(), above, uses this function to rename Undo files so that they match a new image ID.
Private Sub RenameAllUndoFiles(ByRef autosaveData As AutosaveXML, ByVal newImageID As Long, ByVal oldImageID As Long)

    Dim oldFilename As String, newFilename As String
    
    'Before starting our processing loop, create a dummy pdUndo object.  This object will help us generate
    ' relevant filenames using PD's standard Undo filename formula.
    Dim tmpUndoEngine As pdUndo
    Set tmpUndoEngine = New pdUndo
    
    Dim cFile As pdFSO
    Set cFile = New pdFSO
    
    'The autosaveData object knows how many autosave files are available
    Dim i As Long
    For i = 0 To autosaveData.undoStackAbsoluteMaximum
    
        oldFilename = tmpUndoEngine.GenerateUndoFilenameExternal(oldImageID, i, autosaveData.originalSessionID)
        newFilename = tmpUndoEngine.GenerateUndoFilenameExternal(newImageID, i, OS.UniqueSessionID())
        
        'Check image data first...
        If Files.FileExists(oldFilename) Then
            Files.FileDeleteIfExists newFilename
            cFile.FileCopyW oldFilename, newFilename
        End If
        
        '...followed by layer data
        If Files.FileExists(oldFilename & ".layer") Then
            Files.FileDeleteIfExists newFilename & ".layer"
            cFile.FileCopyW oldFilename & ".layer", newFilename & ".layer"
        End If
        
        '...followed by selection data
        If Files.FileExists(oldFilename & ".selection") Then
            Files.FileDeleteIfExists newFilename & ".selection"
            cFile.FileCopyW oldFilename & ".selection", newFilename & ".selection"
        End If
        
    Next i

End Sub
