VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdUndo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'Undo/Redo Handler
'Copyright 2001-2026 by Tanner Helland
'Created: 2/April/01
'Last updated: 26/April/19
'Last update: modify File > Revert behavior; it now creates a *new* entry on the Undo stack (similar to Photoshop),
'             which allows the user to return to previous states even after reverting.
'
'PhotoDemon has had Undo/Redo capabilities for a long time.  The first incarnation (way back in 2001) stored a single
' byte-array copy of the image in RAM, allowing for a lone Undo step if necessary.  Eventually this was converted
' to a more robust HDD-based solution, and further updates have tacked on features to (the very poor) original
' Undo/Redo framework.
'
'In 2013, I finally had enough of the program's shoddy Undo/Redo code, so I merged all the various Undo/Redo subs
' into a single class, and gave each pdImage a copy of said class.  As part of the rewrite, the Undo/Redo engine
' itself was rewritten in a much cleaner fashion, including moving all interface-related code outside the class.
' (Specifically, UI considerations were handed off to PD's central SyncInterfaceToCurrentImage function.)
'
'In 2014, as part of adding Layers support, the Undo/Redo engine was further modified from a lazy implementation to
' an active one.  What does this mean?  Basically, the previous "lazy" implementation did not create Undo/Redo data
' until absolutely necessary (e.g. after at least one image-changing action had been initiated, and right before a
' new action was about to occur).  Such an ultra-conservative approach had benefits; for example, Undo data was
' never created until an image was actually edited, improving image load time, and Redo data was never created
' unless Undo was actually pressed.  While good for conserving memory, this system ultimately proved untenable, as it
' required a lot of special handling re: "what happened previously", which was necessary for handling Undo creation
' correctly for odd edge cases (like when an image edit is applied after a selection modification). To simplify the
' engine and make it capable of storing Undo/Redo data other than just the full pdImage stack, the engine was
' converted to an Active implementation.  Now, any time the image is modified, Undo data is immediately written
' to file following the completion of the modification action.
'
'Similarly, when an image is first loaded, a full backup of the image is immediately written to a temp file.  This
' incurs a slight penalty in load-time performance, but I believe the trade-off is worth it because this backup file
' is also useful for the Autosave engine, as it gives us something to work from in the event of an unsafe shutdown.
'
'Later in 2014, writing a full pdImage stack for each program action proved a terrible idea, due to the size,
' slowness, and complexity of each Undo/Redo file.  To combat this, the Undo/Redo engine was rewritten yet again -
' but this time, as a full-featured "diff engine", not unlike what powers Git.  Specifically, PD's Undo/Redo engine
' now saves only the minimum amount of data necessary to reconstruct an image from its immediately previous Undo/Redo
' entry.  This allows for very small Undo/Redo files, at a trade-off with a large amount of new complexity in the
' Undo/Redo engine itself.  I believe this approach is the best one possible, so I'm hoping 2014 marks the end of
' massive Undo/Redo engine rewrites.  (Fingers crossed.)
'
'Undo/Redo data is still saved to the HDD, and will remain that way for the forseeable future.  Unlimited operations
' are allowed.  This class does not check HDD space availability, so errors may be thrown if the user runs out of free
' space.  Similarly, clearing the specified TEMP folder in the middle of program operation is a bad idea.  PD will not
' crash (probably), but as expected, Undo/Redo functionality will stop working.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'If you need to debug Undo/Redo data, set this to TRUE; this will forcibly print the current Undo stack contents
' to the debug window whenever the stack changes.
Private Const UNDO_DEBUG_MODE As Boolean = False

'Similarly, detailed timing reports can also be provided
Private Const UNDO_DETAILED_REPORTS As Boolean = False

'In the Undo/Redo stack, which entry is this image currently on?
Private m_UndoNum As Long

'How many Undo/Redo entries exist in our stack?
Private m_UndoMax As Long
Private m_UndoTotal As Long

'Are Undo/Redo operations allowed for this image?  (This is used to determine the state of undo/redo buttons and
' menu items.)
Private m_UndoState As Boolean, m_RedoState As Boolean

'Information about each Undo entry is stored in an array; the array is dynamically resized as necessary when new
' Undos are created.  We track the ID of each action in preparation for a future History browser that allows the
' user to jump to any arbitrary Undo/Redo state.  (Also, to properly update the text of the Undo/Redo menu and
' buttons so the user knows which action they are undo/redoing.)
Private m_UndoStack() As PD_UndoEntry

'If the user performs 10 actions, then presses Undo 5 times and saves the image, we need to disable the Save button
' at that Undo point, because it was the last point at which the image was saved.  Thus we must track the last point
' in the Undo/Redo stack where the image was saved.  Note that we separately track flat image saves (JPEG, PNG, etc),
' and layered image saves (PDI).  External functions can request either or both of these values as they need 'em.
Private m_UndoNumAtLastSave_Flat As Long, m_UndoNumAtLastSave_PDI As Long

'We need access to the parent class so we can request things like "write your current data out to file".
Public parentPDImage As pdImage

'As part of PD's Undo/Redo chain, we can write the Undo/Redo stack to an XML file.  These constants are used to help
' us ID and parse such data correctly.
Private Const PD_UNDO_VERSION_ID As Long = 64

'The size of the Undo/Redo Thumbnail
Private Const UNDO_THUMB_LARGE As Long = 128

'Certain parts of the Undo/Redo header are written out to persistent files.  For example, a full copy of the current
' Undo/Redo stack is written to an XML file, and in the case of a crash, we use that file to reconstruct the active image.
' Rather than regenerate these filenames on every call, we generate them once and cache them.  (Note that the filename,
' by design, does not include an extension - the idea is that you append the extension you want, depending on what kind
' of data you're writing.)
Private m_FSO As pdFSO
Private m_HeaderFilename As String, m_HeaderXMLHandle As Long, m_UTF8() As Byte, m_ThumbnailFileHandle As Long

'In 7.0, we started tracking Undo generation times on a per-type basis.  If operations get too slow, we crank up lz4
' acceleration levels to compensate.  (NOTE: this only occurs under default PD settings, e.g. compression level 1.
' Custom Undo/Redo compression uses other, hard-coded settings.)
Private m_UndoAcceleration() As Long
Private Const UNDO_TIME_LIMIT_MS As Long = 100

'Some external functions may need to modify their behavior when called during Undo/Redo operations.  As such, this class
' sets internal flags while Undo/Redo operations are live.
Private m_UndoRedoActive As Boolean

'Returns TRUE if this class instance is in the middle of an operation
Friend Function IsUndoEngineActive() As Boolean
    IsUndoEngineActive = True
End Function

'Returns TRUE if this class has successfully started tracking Undo data
Friend Function HasFirstUndoWriteOccurred() As Boolean
    HasFirstUndoWriteOccurred = (m_UndoTotal >= 1)
End Function

'When an image is saved, this class needs to be notified so that it can make a note of the current undo index
Friend Sub NotifyImageSaved(ByVal typeOfSave As PD_SAVE_EVENT)
    
    Dim saveIndex As Long
    If (m_UndoNum > 0) Then saveIndex = m_UndoNum Else saveIndex = 1
    
    Select Case typeOfSave
        
        'This value should never be passed in!
        Case pdSE_AnySave
            Debug.Print "The pdSE_AnySave type is only for GETTING image save state, not SETTING it!  Fix this!"
            m_UndoNumAtLastSave_Flat = saveIndex
            m_UndoNumAtLastSave_PDI = saveIndex
                
        'Image has been saved to PDI format, meaning layers are intact
        Case pdSE_SavePDI
            m_UndoNumAtLastSave_PDI = saveIndex
                
        'Image has been saved to some flat format (JPEG, PNG, etc).  Layers are not intact.
        Case pdSE_SaveFlat
            m_UndoNumAtLastSave_Flat = saveIndex
        
    End Select
    
End Sub

Friend Function GetUndoState() As Boolean
    GetUndoState = m_UndoState
End Function

Friend Function GetRedoState() As Boolean
    GetRedoState = m_RedoState
End Function

'Create an Undo entry (e.g. save a copy of the present image or selection to the temp directory)
' Inputs:
'  1) the ID string of the process that called this action (e.g. "Gaussian blur")
'  2) optionally, the type of Undo that needs to be created.  Make absolutely certain that this value is correct, because
'     it determines what type of Undo data is saved (e.g. full layers stack, single layer, selection, etc).
'  3) optionally, the tool associated with this action.  This value is not currently used, but in the future it could
'     be helpful for automatically selecting the tool associated with a given Undo/Redo point.
'  4) optionally, setting coalescenceAllowed to TRUE lets the Undo engine know that it is acceptable to overwrite the
'     previous Undo entry with this new one, if the layer ID and processID match.  The caller must have *explicit* knowledge
'     that this behavior is appropriate, as there is no way to retrieve the overwritten Undo entry.  (PD uses this for
'     non-destructive changes to a layer, if the non-destructive type doesn't change between Undo requests.)
Friend Sub CreateUndoData(ByRef srcProcData As PD_ProcessCall, Optional ByVal targetLayerID As Long = -1, Optional ByVal coalescenceAllowed As Boolean = False)
    
    m_UndoRedoActive = True
    
    Dim timeAtUndoStart As Currency
    VBHacks.GetHighResTime timeAtUndoStart
    
    'If the caller is okay with us coalescing this into the previous entry, test for coalescence acceptability
    If (coalescenceAllowed And (m_UndoNum > 0)) Then
        
        'If this processID and layerID match the previous ones, we are okay to just overwrite that entry with this one.
        If Strings.StringsEqual(srcProcData.pcID, m_UndoStack(m_UndoNum - 1).srcProcCall.pcID) And (targetLayerID = m_UndoStack(m_UndoNum - 1).undoLayerID) Then
            
            'As an additional failsafe, make sure the Undo type is layer-safe, and vector-safe.  This guarantees that only
            ' non-destructive changes are allowed to coalesce.
            If (srcProcData.pcUndoType = UNDO_Layer_VectorSafe) And (m_UndoStack(m_UndoNum - 1).srcProcCall.pcUndoType = UNDO_Layer_VectorSafe) Then
                m_UndoNum = m_UndoNum - 1
            End If
            
        End If
        
    End If
    
    'Forward a save request to the actual SaveUndoData function.  It handles the messy business of actually writing the correct
    ' filetypes (which it determines based on the current processType and targetLayer)
    Dim finalUndoFileSize As Long
    Saving.SaveUndoData parentPDImage, GetCurrentUndoFilename(), srcProcData.pcUndoType, targetLayerID, m_UndoAcceleration(srcProcData.pcUndoType), finalUndoFileSize
    
    'Store all relevant information about this Undo action, so we can access it in the future if necessary
    If (m_UndoNum <> UBound(m_UndoStack)) Then ReDim Preserve m_UndoStack(0 To m_UndoNum) As PD_UndoEntry
    With m_UndoStack(m_UndoNum)
        .srcProcCall = srcProcData
        .undoLayerID = targetLayerID
        .undoFileSize = finalUndoFileSize
        If UNDO_DETAILED_REPORTS Then PDDebug.LogAction "Undo file size: " & Format$(CDbl(.undoFileSize) / 1000#, "#,#0") & " kb"
    End With
    
    'If the present image is NOT the last image in the undo train (i.e., the user pressed Undo before running
    ' the current operation), treat this action as the new "last" Undo entry in the stack.  However, we still need
    ' to remember that Undo files exist past this point, so we can clean them up when we exit - hence the need for
    ' a separate m_UndoTotal variable.
    m_UndoMax = m_UndoNum
    
    'Next, we will store a copy of the current image thumbnail in two sizes: 256x256, and 48x48
    Dim startTime As Currency
    VBHacks.GetHighResTime startTime
    
    'First, obtain a large copy of the thumbnail.  The source of the thumbnail changes depending on the undo type.
    Dim largeThumbSize As Long: largeThumbSize = Interface.FixDPI(UNDO_THUMB_LARGE)
    If (m_UndoStack(m_UndoNum).thumbnailLarge Is Nothing) Then Set m_UndoStack(m_UndoNum).thumbnailLarge = New pdDIB
    
    Select Case srcProcData.pcUndoType
    
        Case UNDO_Nothing
            parentPDImage.RequestThumbnail m_UndoStack(m_UndoNum).thumbnailLarge, largeThumbSize, False
        
        Case UNDO_Everything
            parentPDImage.RequestThumbnail m_UndoStack(m_UndoNum).thumbnailLarge, largeThumbSize, False
            
        Case UNDO_Image, UNDO_Image_VectorSafe
            parentPDImage.RequestThumbnail m_UndoStack(m_UndoNum).thumbnailLarge, largeThumbSize, False
        
        Case UNDO_ImageHeader
            parentPDImage.RequestThumbnail m_UndoStack(m_UndoNum).thumbnailLarge, largeThumbSize, False
        
        Case UNDO_Layer, UNDO_Layer_VectorSafe
            parentPDImage.GetLayerByID(targetLayerID).RequestThumbnail m_UndoStack(m_UndoNum).thumbnailLarge, largeThumbSize, False
        
        Case UNDO_LayerHeader
            parentPDImage.RequestThumbnail m_UndoStack(m_UndoNum).thumbnailLarge, largeThumbSize, False
        
        Case UNDO_Selection
            parentPDImage.MainSelection.RequestThumbnail m_UndoStack(m_UndoNum).thumbnailLarge, largeThumbSize
            
    End Select
    
    If UNDO_DETAILED_REPORTS Then PDDebug.LogAction "Undo thumbnail request took: " & Format$(VBHacks.GetTimerDifferenceNow(startTime) * 1000, "#0.00") & " ms"
    
    'Increase our Undo count
    m_UndoNum = m_UndoNum + 1
    m_UndoMax = m_UndoMax + 1
    
    'If the present image is *truly* the last one in the Undo train, note that we have more Undo files to delete
    ' when our parent is unloaded.
    If (m_UndoMax > m_UndoTotal) Then m_UndoTotal = m_UndoMax
    
    'If this is not the first Undo entry, make a note that the user is now allowed to initiate Undo actions.
    m_UndoState = (m_UndoNum > 1)
    
    'Because we have just created Undo data, the user is not allowed to initiate Redo actions.
    m_RedoState = False
    
    'If this action changed pixel data, mark the image as NOT being saved in its current state.
    If (srcProcData.pcUndoType <> UNDO_Nothing) And (srcProcData.pcUndoType <> UNDO_Selection) And (m_UndoNum > 1) Then
        parentPDImage.SetSaveState False, pdSE_AnySave
    End If
    
    'New to version 6.4 is our persistent creation of Undo XML header files.  These are extraordinarily helpful
    ' for both regular Undo/Redo behavior, as well as AutoSave recreation.  We always create a copy of the Undo
    ' stack, in XML format, in the temporary folder.
    Me.WriteUndoDataToXML GenerateUndoHeaderFilename()
    
    'To inspect the current state of the Undo/Redo stack, uncomment this line.
    If UNDO_DEBUG_MODE Then DebugUndo "new Undo entry created"
    If UNDO_DETAILED_REPORTS Then PDDebug.LogAction "Time taken for total Undo process: " & Format$(VBHacks.GetTimerDifferenceNow(timeAtUndoStart) * 1000, "#0.00") & " ms"
    
    m_UndoRedoActive = False
    
    'If the default compression setting is active, calculate an acceleration ratio to keep Undo generation as snappy as possible.
    If (g_UndoCompressionLevel = 1) Then CalculateUndoAcceleration srcProcData.pcUndoType, VBHacks.GetTimerDifferenceNow(timeAtUndoStart) * 1000#
    
End Sub

'After creating Undo data, call this function calculate new Undo acceleration parameters.  (We dynamically decrease compression
' ratios in order to improve export speed.)
Private Sub CalculateUndoAcceleration(ByVal processType As PD_UndoType, ByVal timeTaken As Double)

    'When the user uses default Undo/Redo compression settings, we use lz4 to apply a high-speed compression to our outgoing
    ' save data.  Our goal is to keep Undo/Redo generation under a hard-coded limit (UNDO_TIME_LIMIT_MS).  If Undo/Redo writes
    ' start taking more time than this, we'll crank up lz4 acceleration settings to compensate.  This all happens silently,
    ' and it's calculated by Undo type (because some types, e.g. header-only writes, are much faster than other types, so we
    ' can happily maximize compression without worrying about performance).
    Dim curLz4Level As Long
    curLz4Level = m_UndoAcceleration(processType)
    
    Dim lz4Levels As Long, timeRatio As Double
    
    'If the time taken was unacceptably long, take proactive action
    If (timeTaken > UNDO_TIME_LIMIT_MS) Then
    
        'How badly are we over our target?
        timeRatio = (timeTaken / CDbl(UNDO_TIME_LIMIT_MS))
        
        'lz4's acceleration levels give a 3-4% performance boost per level.  Translate our "over budget" ratio
        ' into discrete lz4 acceleration levels.
        lz4Levels = Int((timeRatio - 1#) * 300# + 0.5)
        If (lz4Levels < 1) Then lz4Levels = 1
        
        'Add that to our existing compression level, then we'll try again on our next request
        m_UndoAcceleration(processType) = m_UndoAcceleration(processType) + lz4Levels
        
        'Make sure we're within acceptable lz4 compression level boundaries
        If (m_UndoAcceleration(processType) > Plugin_lz4.Lz4_GetMaxAccelerationLevel()) Then m_UndoAcceleration(processType) = Plugin_lz4.Lz4_GetMaxAccelerationLevel()
        If (m_UndoAcceleration(processType) < 2) Then m_UndoAcceleration(processType) = 2
        
        If UNDO_DETAILED_REPORTS Then
            If (m_UndoAcceleration(processType) <> curLz4Level) Then PDDebug.LogAction "Undo is too slow, accelerating lz4 acceleration to compensate. Level changed from " & curLz4Level & " to " & m_UndoAcceleration(processType) & " (" & Format$(timeTaken, "0") & " ms)"
        End If
        
    Else
    
        'We came in under-budget.  If we previously cranked up our lz4 acceleration level, drop it slightly to try and find
        ' the "sweet spot" for this image and undo type.
        If (m_UndoAcceleration(processType) <> -1) Then
            
            If (timeTaken > 0#) Then timeRatio = CDbl(UNDO_TIME_LIMIT_MS) / timeTaken Else timeRatio = 0.01
            lz4Levels = Int((timeRatio - 1#) * 100# + 0.5)
            If (lz4Levels < 1) Then lz4Levels = 1
        
            m_UndoAcceleration(processType) = m_UndoAcceleration(processType) - lz4Levels
            
            'If we get too low, restore default settings
            If (m_UndoAcceleration(processType) < 2) Then m_UndoAcceleration(processType) = -1
            
            If UNDO_DETAILED_REPORTS Then
                If (m_UndoAcceleration(processType) <> curLz4Level) Then PDDebug.LogAction "Undo is fast enough, decreasing lz4 acceleration to compensate. Level changed from " & curLz4Level & " to " & m_UndoAcceleration(processType)
            End If
        
        'If we're already running at maximum lz4 compression, we can't improve compression further.
        'Else
        
        End If
        
    End If
    
End Sub

'This (rather odd) function is required by the Crop tool.  Because the Crop tool forcibly removes the current selection -
' but only after making use of it - it causes the Undo/Redo engine's selection tracking to fall out of sync.
' To work around this without adding unnecessary dummy entries to the Undo chain, the Crop command calls this function
' prior to processing.
'
'This function simply updates the previous Undo entry to include a full record of everything: the whole image,
' and any active selections.  This ensures that if the user uses Undo to return to a pre-Crop state, everything is
' restored properly.  (Without this function, Undo-ing a crop action would cause the selection to be lost unless the
' user returned to the point when the selection was first created.)
Friend Sub ForceLastUndoDataToIncludeEverything()
    
    m_UndoRedoActive = True
    
    'Overwrite the previous action's Undo data with a full image backup.  Note that all the settings from the previous
    ' action are retained; the only thing we're doing differently is specifying the Undo type as UNDO_EVERYTHING.
    Saving.SaveUndoData parentPDImage, GenerateUndoFilename(m_UndoNum - 1), UNDO_Everything, -1
    
    'Update the stack entry at that point to reflect the UNDO_EVERYTHING usage
    With m_UndoStack(m_UndoNum - 1)
        .srcProcCall.pcUndoType = UNDO_Everything
        .undoLayerID = -1
    End With
    
    'Update the Undo XML header as well.
    WriteUndoDataToXML GenerateUndoHeaderFilename()
        
    m_UndoRedoActive = False
        
End Sub

'This special command can be used to revert the image to its last saved state.  For images that have never been
' saved (e.g. those pasted from the clipboard, etc), it will undo to the image's original state.
Friend Sub RevertToLastSavedState()
    
    m_UndoRedoActive = True
    
    Dim targetUndoPoint As Long
    
    If (m_UndoNumAtLastSave_Flat > m_UndoNumAtLastSave_PDI) Then
        targetUndoPoint = m_UndoNumAtLastSave_Flat
    Else
        targetUndoPoint = m_UndoNumAtLastSave_PDI
    End If
    
    'Failsafe check to make sure the specified index is valid
    If (targetUndoPoint < 1) Then targetUndoPoint = 1
    
    'In older versions of PD, we would now simply move the Undo pointer to this index in the stack.
    ' This is identical to using Edit > Undo History, but with us auto-calculating the index for you.
    'Me.MoveToSpecificUndoPoint targetUndoPoint
    
    'However, in studying Photoshop's implementation of Revert, they actually treat it as an *edit*
    ' and it gets added to the Undo stack as its own entry!  I actually like this approach better,
    ' as it protects against actual File > Revert clicks (probably useful, given its proximity to
    ' File > Save As).
    
    'To implement something similar, we need to use a custom hybrid of the MoveToSpecificUndoPoint
    ' (to restore the correct image data) and a regular Undo event (to add the restored data to
    ' the stack).
    
    'Here is the required code from MoveToSpecificUndoPoint, indented deliberately:
        Message "Reconstructing image..."
        
        'Start by resetting the image to its default state, e.g. Undo 0.
        ImageImporter.LoadUndo GenerateUndoFilename(0), UNDO_Everything, UNDO_Everything, , True
        
        'From this point, we now want to advance the Undo stack to the requested index by applying successive diff entries
        ' until the final index is reached.
        If (targetUndoPoint > 1) Then
        
            Dim i As Long
            For i = 1 To targetUndoPoint - 1
                LoadUndo GenerateUndoFilename(i), m_UndoStack(i).srcProcCall.pcUndoType, m_UndoStack(i).srcProcCall.pcUndoType, m_UndoStack(i).undoLayerID, True
            Next i
        
        End If
        
    'The image's pixel data now matches the last-saved copy of the image (or the image's original state,
    ' if the user hasn't saved it this session).  We don't need to do any further work; PD's standard
    ' Undo/Redo path will take care of creating a new stack entry at this point.
    m_UndoRedoActive = False
    
End Sub

'Move to a specific location in the Undo/Redo chain.  Note that this function may take some time if the requested
' point is quite far into the Undo/Redo chain.
Friend Sub MoveToSpecificUndoPoint(ByVal undoIndex As Long)
    
    m_UndoRedoActive = True
    
    Message "Reconstructing image..."
    
    'Start by resetting the image to its default state, e.g. Undo 0.
    ImageImporter.LoadUndo GenerateUndoFilename(0), UNDO_Everything, UNDO_Everything, , True
    
    'From this point, we now want to advance the Undo stack to the requested index by applying successive diff entries
    ' until the final index is reached.
    If (undoIndex > 1) Then
    
        Dim i As Long
        For i = 1 To undoIndex - 1
            LoadUndo GenerateUndoFilename(i), m_UndoStack(i).srcProcCall.pcUndoType, m_UndoStack(i).srcProcCall.pcUndoType, m_UndoStack(i).undoLayerID, True
        Next i
    
    End If
    
    'Our parent pdImage object has now been successively reconstructed.  Set all Undo variables to their proper values,
    ' then redraw the screen and exit.
    
    'm_UndoNum is always one higher than the state of the image.  (This is easily remembered because the base Undo copy
    ' is created at point 0 right when an image is loaded, and because the m_UndoNum is immediately advanced,
    ' it starts at value 1.)
    m_UndoNum = undoIndex
    
    'Allow saving contingent on the image's state
    parentPDImage.SetSaveState False, pdSE_AnySave
    If (m_UndoNum = m_UndoNumAtLastSave_PDI) Then parentPDImage.SetSaveState True, pdSE_SavePDI
    If (m_UndoNum = m_UndoNumAtLastSave_Flat) Then parentPDImage.SetSaveState True, pdSE_SaveFlat
    
    'See if Undoing is allowed
    m_UndoState = (m_UndoNum > 1)
    
    'See if redoing is allowed
    m_RedoState = (m_UndoNum <> m_UndoMax)
    
    'Select the relevant tool for this action, if one is provided
    'NOTE: as of 7.0, I have disabled this capability.  It's not as useful as I thought it would be, and it's very
    'confusing when the program swaps tools without your input.
    'If (GetUndoTool > -1) Then toolbar_Toolbox.SelectNewTool GetUndoTool
    
    'To inspect the current state of the Undo/Redo stack, uncomment this line.
    If UNDO_DEBUG_MODE Then DebugUndo "Undo pointer moved by request to a specific location"
    
    'New to version 6.4 is our persistent creation of Undo XML header files.  These are extraordinarily helpful
    ' for both regular Undo/Redo behavior, as well as AutoSave recreation.
    Me.WriteUndoDataToXML GenerateUndoHeaderFilename()
    
    'Now that all layers have been updated against their most recent diff, we can finally redraw the screen
    Viewport.Stage1_InitializeBuffer PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    
    Message "Image successfully restored."
    
    m_UndoRedoActive = False
    
End Sub

'Safe XML-based wrapper for MoveToSpecificUndoPoint(), above.
Friend Sub MoveToSpecificUndoPoint_XML(ByVal paramList As String)
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString paramList
    Me.MoveToSpecificUndoPoint cParams.GetLong("UndoHistoryPoint", 1)
End Sub

'Restore the last undo entry, e.g. "the user just clicked Undo".
Friend Sub RestoreUndoData()
    
    m_UndoRedoActive = True
    
    'Decrement the Undo stack pointer
    m_UndoNum = m_UndoNum - 1
    If (m_UndoNum < 0) Then m_UndoNum = 0
    
    'If this is not the last Undo state, continue to allow the user to Undo, but if we are at the bottom of
    ' the Undo stack (e.g. the image is in its original state), disable further Undo-ing.
    m_UndoState = (m_UndoNum > 1)
    
    'If we are not at the top of the Undo stack, also enable Redo-ing
    m_RedoState = (m_UndoNum < m_UndoMax)
        
    'If this image initially existed on the user's hard drive (e.g. it was not imported from the clipboard or
    ' some other source), note that the image has now BEEN SAVED (because it has not changed at all from its
    ' original appearance).
    If (m_UndoNum = m_UndoNumAtLastSave_Flat) Or (m_UndoNum = m_UndoNumAtLastSave_PDI) Then
    
        'In the special case of m_UndoNumAtLastSave = 0, check to see if this image exists on the hard drive
        If (LenB(parentPDImage.ImgStorage.GetEntry_String("CurrentLocationOnDisk", vbNullString)) <> 0) Then
            
            If (parentPDImage.GetCurrentFileFormat = PDIF_PDI) Then
                parentPDImage.SetSaveState True, pdSE_SavePDI
            Else
                parentPDImage.SetSaveState True, pdSE_SaveFlat
            End If
            
        Else
            parentPDImage.SetSaveState False, pdSE_AnySave
        End If
        
    'If the image was imported (e.g. from a scanner or the clipboard), note that it has NOT BEEN SAVED just because
    ' the user Undo'd all the way back to its original state.
    Else
        parentPDImage.SetSaveState False, pdSE_AnySave
    End If
    
    'Select the relevant tool for this action, if we were supplied one previously
    'NOTE: as of 7.0, I have disabled this capability.  It's not as useful as I thought it would be, and it's very
    'confusing when the program swaps tools without your input.
    'If (GetUndoTool > -1) Then toolbar_Toolbox.SelectNewTool GetUndoTool
    
    'With all prep work complete, we now need to figure out the most recent diff file that possesses data we need.
    ' (Remember that PhotoDemon doesn't save a full copy of the image on every Undo action.  Instead, it only saves
    ' the data that has changed.  This makes Undo-ing a bit more complicated, as we sometimes have to revert to a
    ' much earlier state, then "replay" all subsequent steps to arrive at the correct image state.)
    
    'First, find the nearest relevant diff for this type of Undo operation.  For a typical UNDO_LAYER operation,
    ' for example, this can be an UNDO_EVERYTHING, UNDO_IMAGE, or UNDO_LAYER operation - if the UNDO_LAYER operation
    ' matches this layer ID!  (Similarly, some parameters - like undoLayerID-  will be ignored if this undo entry
    ' isn't layer-specific.)
    Dim nearestDiffIndex As Long
    nearestDiffIndex = FindNearestDiffIndex(m_UndoNum, m_UndoStack(m_UndoNum).srcProcCall.pcUndoType, m_UndoStack(m_UndoNum).undoLayerID)
    
    'Want to know what decision the Undo engine made?  Here it is:
    'Debug.Print "Nearest diff index: " & nearestDiffIndex & "; undo type: " & m_UndoStack(nearestDiffIndex).srcProcCall.pcUndoType
    
    'We now have to handle some tricky diff-related business.  If the Undo type is a full pdImage reconstruction (meaning the entire
    ' layer stack is being recreated as part of the Undo process), we have to re-apply all diffs in two steps: first, we have to
    ' reconstruct the layer stack by finding a diff that contains full image data, e.g. a file of type UNDO_IMAGE (or in some special
    ' cases, UNDO_IMAGEHEADER).  Once the layer stack has been properly recreated, we have to update each layer and layer header
    ' against their most recent diff patch(es).  To prevent the screen from flickering like crazy while we do all this background
    ' processing, we forcibly suspend all screen redraws until our work is complete.
    '
    'Similarly, if we are undoing a single layer action (UNDO_LAYER, UNDO_LAYER_VECTORSAFE), we may need to apply optional
    ' UNDO_LAYERHEADER diffs after the fact, to ensure the layer's header is up-to-date against any header-only changes that
    ' occurred since the latest full layer DIB write-to-file.
    '
    'Selections follow the same rule, but for UNDO_SELECTION types only
    Dim suspendRedraws As Boolean
    If (m_UndoStack(m_UndoNum).srcProcCall.pcUndoType = UNDO_Everything) Or (m_UndoStack(m_UndoNum).srcProcCall.pcUndoType = UNDO_Image) Or (m_UndoStack(m_UndoNum).srcProcCall.pcUndoType = UNDO_Image_VectorSafe) Or (m_UndoStack(m_UndoNum).srcProcCall.pcUndoType = UNDO_ImageHeader) Or (m_UndoStack(m_UndoNum).srcProcCall.pcUndoType = UNDO_Layer) Or (m_UndoStack(m_UndoNum).srcProcCall.pcUndoType = UNDO_Layer_VectorSafe) Then
        
        'If we are reverting the image to its original state, we do not need to process any extra layer data, as we already know
        ' the the source file will be accurate (because a full copy of the image is always created at Undo point 0).
        suspendRedraws = (m_UndoNum > 0)
        
    Else
        suspendRedraws = False
    End If
    
    'With the most recent relevant diff file identified, we can now launch the undo loading routine for that diff.  It will
    ' automatically apply any changes against the relevant parts of the current pdImage object.
    ImageImporter.LoadUndo GenerateUndoFilename(nearestDiffIndex), m_UndoStack(nearestDiffIndex).srcProcCall.pcUndoType, m_UndoStack(m_UndoNum).srcProcCall.pcUndoType, m_UndoStack(m_UndoNum).undoLayerID, suspendRedraws
    
    'If redraws have been suspended, we know that we need to reconstruct one or more layers in the image against their
    ' most recent diff.
    If suspendRedraws Then
    
        Dim nearestLayerDiffDIB As Long, nearestLayerDiffHeader As Long, nearestImageDiffHeader As Long, curLayerID As Long
        Dim i As Long
        
        'The amount of diffs we need to reconstruct varies by Undo type.
        
        'If this is a full image Undo, all layers need to be checked against their most recent diffs.
        If (m_UndoStack(m_UndoNum).srcProcCall.pcUndoType = UNDO_Everything) Or (m_UndoStack(m_UndoNum).srcProcCall.pcUndoType = UNDO_Image) Or (m_UndoStack(m_UndoNum).srcProcCall.pcUndoType = UNDO_Image_VectorSafe) Or (m_UndoStack(m_UndoNum).srcProcCall.pcUndoType = UNDO_ImageHeader) Then
        
            'For UNDO_EVERYTHING, UNDO_IMAGE, and UNDO_IMAGE_VECTORSAFE entries, we may need to reconstruct the layer stack
            ' before any further processing occurs.
            If (m_UndoStack(m_UndoNum).srcProcCall.pcUndoType = UNDO_Everything) Or (m_UndoStack(m_UndoNum).srcProcCall.pcUndoType = UNDO_Image) Or (m_UndoStack(m_UndoNum).srcProcCall.pcUndoType = UNDO_Image_VectorSafe) Then
            
                'Search for the nearest relevant pdImage header diff, as it may have changed independent of the rest
                ' of the pdImage stack (for example, if the user has rearranged layer order)
                nearestImageDiffHeader = FindNearestDiffIndex(m_UndoNum, UNDO_ImageHeader)
                
                'If a more recent layer header was found, apply it now.
                If (nearestImageDiffHeader > nearestDiffIndex) Then
                    ImageImporter.LoadUndo GenerateUndoFilename(nearestImageDiffHeader), m_UndoStack(nearestImageDiffHeader).srcProcCall.pcUndoType, UNDO_ImageHeader, , suspendRedraws
                End If
                
                'Also reset the nearestDiffIndex value to match the header's location; subsequent layer checks
                ' will use this as their baseline for determining whether they need to load additional data for that layer.
                nearestDiffIndex = nearestImageDiffHeader
            
            End If
                        
            'Now we know that the layer stack has the correct number of layers, in the correct order.
            ' Loop through each layer in the image, searching for more recent diffs for either the layer itself or the
            ' layer header (relevant if the user has made on-canvas changes that don't necessarily affect the layer DIB).
            For i = 0 To PDImages.GetActiveImage.GetNumOfLayers - 1
            
                'Grab the current layer ID (all undo functions operate on layer ID, not index, as index may change unexpectedly)
                curLayerID = PDImages.GetActiveImage.GetLayerByIndex(i).GetLayerID
            
                'Start by reconstructing a layer DIB using the default UNDO_LAYER type.  (Note that this check will also
                ' return UNDO_LAYER_VECTORSAFE and UNDO_LAYER_HEADER results.)
                nearestLayerDiffDIB = FindNearestDiffIndex(m_UndoNum, UNDO_Layer, curLayerID)
                
                'If the nearest diff for this layer is *not* the same as the nearest diff for the full image, apply this diff
                ' to the layer in question.
                If (nearestLayerDiffDIB > nearestDiffIndex) Then
                    ImageImporter.LoadUndo GenerateUndoFilename(nearestLayerDiffDIB), m_UndoStack(nearestLayerDiffDIB).srcProcCall.pcUndoType, UNDO_Layer, curLayerID, suspendRedraws
                End If
                
                'Next, we must reconstruct the layer header, as it too may have changed (for example, if the user moved the layer
                ' since the last action)
                nearestLayerDiffHeader = FindNearestDiffIndex(m_UndoNum, UNDO_LayerHeader, curLayerID)
                
                'If a more recent layer header was found, apply it now.
                If (nearestLayerDiffHeader > nearestDiffIndex) And (nearestLayerDiffHeader > nearestLayerDiffDIB) Then
                    ImageImporter.LoadUndo GenerateUndoFilename(nearestLayerDiffHeader), m_UndoStack(nearestLayerDiffHeader).srcProcCall.pcUndoType, UNDO_LayerHeader, curLayerID, suspendRedraws
                End If
                
            Next i
            
            'Finally, repeat the above steps for any selection-specific changes that may have occurred.
            
            ' (NOTE: I added this fix on 23 October 2017, and to be entirely honest, I'm not sure I comprehend why it works,
            ' or what other ramifications it may have.  All I know is that there's an extremely obnoxious - and reproducible -
            ' bug that occurs when you apply the following actions:
            ' 1) Create selection, 2) Move selection, 3) Crop image, 4) Undo back to (1)
            ' At this point, your initial selection should be on-screen, but for some reason, it isn't!  The Undo engine has
            ' instead chosen to restore the original image state, without an active selection.  (Weirdly, if you UNDO one
            ' more time, back to the original image state, then REDO once, the selection appears, as it should.)  Adding this
            ' additional check solves the problem completely.
            Dim nearestSelectionDiff As Long
            nearestSelectionDiff = FindNearestDiffIndex(m_UndoNum, UNDO_Selection)
            If (nearestSelectionDiff > nearestDiffIndex) Then
                ImageImporter.LoadUndo GenerateUndoFilename(nearestSelectionDiff), m_UndoStack(nearestSelectionDiff).srcProcCall.pcUndoType, UNDO_Selection, , suspendRedraws
            End If
        
        'If this is a just a layer Undo, we simply need to verify the layer header against any diffs more recent than the last
        ' full DIB reconstruction for this layer.
        ElseIf (m_UndoStack(m_UndoNum).srcProcCall.pcUndoType = UNDO_Layer) Or (m_UndoStack(m_UndoNum).srcProcCall.pcUndoType = UNDO_Layer_VectorSafe) Then
        
            'Grab the current layer ID (all undo functions operate on layer ID, not index, as index may change unexpectedly)
            curLayerID = m_UndoStack(m_UndoNum).undoLayerID
            
            'Next, look for the most recent layer header diff, as it may have changed independent of the layer DIB
            ' (for example, if the user moves a layer after painting on it)
            nearestLayerDiffHeader = FindNearestDiffIndex(m_UndoNum, UNDO_LayerHeader, curLayerID)
                
            'If a more recent layer header was found, apply it now.
            If (nearestLayerDiffHeader > nearestDiffIndex) Then
                LoadUndo GenerateUndoFilename(nearestLayerDiffHeader), m_UndoStack(nearestLayerDiffHeader).srcProcCall.pcUndoType, UNDO_LayerHeader, curLayerID, suspendRedraws
            End If
        
        End If
        
        'Now that all layers have been updated against their most recent diff, we can finally redraw the screen
        Viewport.Stage1_InitializeBuffer PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    
    End If
    
    'To inspect the current state of the Undo/Redo stack, uncomment this line.
    If UNDO_DEBUG_MODE Then DebugUndo "user clicked Undo"
    
    'New to version 6.4 is our persistent creation of Undo XML header files.  These are extraordinarily helpful
    ' for both regular Undo/Redo behavior, as well as AutoSave recreation.
    Me.WriteUndoDataToXML GenerateUndoHeaderFilename()
    
    m_UndoRedoActive = False
    
End Sub

'Given an undo Index, find the nearest Diff file that contains the information we need to reconstruct the image's state
' at that point in time.
Private Function FindNearestDiffIndex(ByVal undoIndex As Long, ByVal undoTypeCheck As PD_UndoType, Optional ByVal relevantLayerID As Long = -1) As Long

    'PD Undo files attempt to save only the minimum amount of data needed to reconstruct an image at any given point in time.
    ' I call these files "diffs" because they contain only the image information that has changed as a result of that
    ' undo entry's actions.  When it comes time to Undo an action, we must scan backward through all available diffs, and find
    ' the nearest entry that contains a backup of the changes made.
    '
    'No matter what, the loop below will always terminate at position 0 if no available diffs are found.  This is why we must
    ' always create a full image copy after the image is loaded, as it serves as our baseline for all subsequent diffs.
        
    'Debug.Print "Searching for undo type: " & undoTypeCheck
    
    Dim i As Long
    For i = undoIndex - 1 To 0 Step -1
    
        'Diff files can cover multiple types of data.  For example, some actions - like flattening the image - make a copy
        ' of the full pdImage stack.  This can be used to recover the pdImage header, individual layer headers, or individual
        ' layer data, which means such an entry works for multiple types of Undo requests.
        '
        'This select case statement matches the requested undoIndex against any valid diff file for that type of Undo data.
        Select Case m_UndoStack(i).srcProcCall.pcUndoType
        
            'UNDO_EVERYTHING works for all possible Undo types, as it has a copy of everything we would ever need!
            Case UNDO_Everything
                FindNearestDiffIndex = i
                Exit For
            
            'UNDO_IMAGE contains relevant diff information for everything but selections
            Case UNDO_Image
                If (undoTypeCheck <> UNDO_Selection) Then
                    FindNearestDiffIndex = i
                    Exit For
                End If
                
            'UNDO_IMAGE_VECTORSAFE is identical to UNDO_IMAGE
            Case UNDO_Image_VectorSafe
                If (undoTypeCheck <> UNDO_Selection) Then
                    FindNearestDiffIndex = i
                    Exit For
                End If
            
            'UNDO_IMAGEHEADER contains relevant diff information for itself, and UNDO_LAYERHEADER only.
            Case UNDO_ImageHeader
                If (undoTypeCheck = UNDO_ImageHeader) Or (undoTypeCheck = UNDO_LayerHeader) Then
                    FindNearestDiffIndex = i
                    Exit For
                End If
            
            'UNDO_LAYER contains relevant diff information for matching UNDO_LAYER and UNDO_LAYER_HEADER actions,
            ' but *only if the layer ID matches the requested one!*
            Case UNDO_Layer
                If ((undoTypeCheck = UNDO_Layer) Or (undoTypeCheck = UNDO_Layer_VectorSafe) Or (undoTypeCheck = UNDO_LayerHeader)) Then
                    
                    If (relevantLayerID = m_UndoStack(i).undoLayerID) Then
                        FindNearestDiffIndex = i
                        Exit For
                    End If
                    
                End If
                
            'UNDO_LAYER_VECTORSAFE is identical to UNDO_LAYER
            Case UNDO_Layer_VectorSafe
                If ((undoTypeCheck = UNDO_Layer) Or (undoTypeCheck = UNDO_Layer_VectorSafe) Or (undoTypeCheck = UNDO_LayerHeader)) Then
                    
                    If (relevantLayerID = m_UndoStack(i).undoLayerID) Then
                        FindNearestDiffIndex = i
                        Exit For
                    End If
                    
                End If
                
            'UNDO_LAYERHEADER contains relevant diff information for matching UNDO_LAYER_HEADER actions,
            ' but *only if the layer ID matches the requested one!*
            Case UNDO_LayerHeader
                If (undoTypeCheck = UNDO_LayerHeader) And (relevantLayerID = m_UndoStack(i).undoLayerID) Then
                    FindNearestDiffIndex = i
                    Exit For
                End If
            
            'UNDO_SELECTION contains relevant diff information for selections only
            Case UNDO_Selection
                If (undoTypeCheck = UNDO_Selection) Then
                    FindNearestDiffIndex = i
                    Exit For
                End If
                
            'This Case Else statement will not be necessary when we have implemented all Diff types.  Until that happens,
            ' however, we can treat it the same as UNDO_IMAGE.
            Case Else
                PDDebug.LogAction "FindNearestDiffIndex hit an unexpected condition??"
                If (undoTypeCheck <> UNDO_Selection) Then
                    FindNearestDiffIndex = i
                    Exit For
                End If
        
        End Select
    
    Next i
    
    'The function value will have been correctly set within the For loop, so we can now exit immediately.

End Function

'Restore a redo entry, e.g. "the user just clicked Redo"
Friend Sub RestoreRedoData()
    
    m_UndoRedoActive = True
    
    'If we're not at the top of the Undo stack, advance the stack index by one
    If (m_UndoNum < m_UndoMax) Then
        
        m_UndoNum = m_UndoNum + 1
        
        'Allow saving contingent on the image's state
        parentPDImage.SetSaveState False, pdSE_AnySave
        If (m_UndoNum = m_UndoNumAtLastSave_PDI) Then parentPDImage.SetSaveState True, pdSE_SavePDI
        If (m_UndoNum = m_UndoNumAtLastSave_Flat) Then parentPDImage.SetSaveState True, pdSE_SaveFlat
    
    End If
    
    'Undoing is always allowed after the user has Redone something
    m_UndoState = True
    
    'If we're at the end of our action chain, disable any further Redo-ing
    m_RedoState = (m_UndoNum <> m_UndoMax)
    
    'Select the relevant tool for this action, if one is provided
    'NOTE: as of 7.0, I have disabled this capability.  It's not as useful as I thought it would be, and it's very
    'confusing to have the program auto-swapping tools without your input.
    'If GetUndoTool > -1 Then toolbar_Toolbox.SelectNewTool GetUndoTool
    
    'The great thing about Redo is that it chronogically moves in the same direction as PD's Undo diff files.  So unlike
    ' Undo - which requires us to traverse the undo timeline in reverse, looking for a relevant diff file - we can simply
    ' apply the next valid diff in line, no extra work required!
    ImageImporter.LoadUndo GenerateUndoFilename(m_UndoNum - 1), GetUndoProcessType, GetUndoProcessType, m_UndoStack(m_UndoNum - 1).undoLayerID
    
    'To inspect the current state of the Undo/Redo stack, uncomment this line.
    If UNDO_DEBUG_MODE Then DebugUndo "user clicked Redo"
    
    'New to version 6.4 is our persistent creation of Undo XML header files.  These are extraordinarily helpful
    ' for both regular Undo/Redo behavior, as well as AutoSave recreation.
    Me.WriteUndoDataToXML GenerateUndoHeaderFilename()
    
    m_UndoRedoActive = False
    
End Sub

'New, simplified interface for "FillDIBWithLastUndoCopy".  You can still call that function if need be, but this function will
' supply a faster result if you just need to know if there are one or more repeatable commands in the current Undo/Redo stack.
Friend Function DoesStackContainRepeatableCommand(Optional ByRef nameOfProcessorAction As String = vbNullString) As Boolean
    
    'If the stack is empty, there is nothing to repeat; return FALSE
    If (m_UndoNum <= 1) Then
        DoesStackContainRepeatableCommand = False
    Else
        
        Dim actionFound As Boolean
        actionFound = False
    
        'Search the stack, looking for a destructive action.  (Note that we do *not* go all the way back to stack position 0,
        ' as that is the original image!)
        Dim i As Long
        For i = m_UndoNum - 1 To 1 Step -1
        
            With m_UndoStack(i)
                
                If (.srcProcCall.pcUndoType <> UNDO_Nothing) Then
                    
                    'Certain actions are not repeatable (e.g. "Fade" and "Repeat" cannot be "faded" or "repeated",
                    ' some selection actions can be repeated - e.g. "grow selection" - but not faded); if we encounter
                    ' a non-repeatable action, continue searching backward until we arrive at a function that *can* be
                    ' faded or repeated.
                    If Actions.IsActionRepeatable(.srcProcCall.pcID, True) Then
                        
                        'If this action is repeatable, but it requires a selection, make sure a selection exists.
                        If (.srcProcCall.pcUndoType = UNDO_Selection) Then
                            If PDImages.IsImageActive() Then
                                If PDImages.GetActiveImage.IsSelectionActive() Then
                                    actionFound = True
                                    nameOfProcessorAction = .srcProcCall.pcID
                                End If
                            End If
                        Else
                            actionFound = True
                            nameOfProcessorAction = .srcProcCall.pcID
                        End If
                        
                        Exit For
                        
                    End If
                    
                End If
            
            End With
            
        Next i
        
        DoesStackContainRepeatableCommand = actionFound
        
    End If
    
End Function

'Find the last modified layer in the image, and fill a supplied DIB with the contents of that layer *prior to its current state*,
' (obviously, as the caller has access to the current state).  This is currently used by the "Edit > Fade..." feature.
'
'The relevant layer index is returned by this function because the user may switch to a new layer, but we still want Fade to be
' available.  (The Undo/Redo engine stores this data anyway, so it's not problematic to receive it.)  To help with menu captioning,
' this function will also return the name of the action responsible for the change.
'
'This function will return FALSE if no layer-specific changes have been made, and thus no Fade action would be relevant.
Friend Function FillDIBWithLastUndoCopy(ByRef dstDIB As pdDIB, ByRef affectedLayerID As Long, Optional ByRef nameOfProcessorAction As String = vbNullString, Optional ByVal checkStateOnly As Boolean = True) As Boolean
    
    FillDIBWithLastUndoCopy = False
    
    If (m_UndoNum <> 0) Then
    
        'First things first: start searching the Undo/Redo stack for a layer-specific edit.
        ' We don't care *which* edit, just that it's a layer-specific action.
        Dim actionFound As Boolean: actionFound = False
        Dim layerChangeIndex As Long: layerChangeIndex = -1
        
        Dim i As Long
        For i = m_UndoNum - 1 To 0 Step -1
        
            'A layer-specific edit has been found.  This is the last action the user committed that modified a layer, so it will contain
            ' a copy of the image as it appeared *after the change was made*.  What this function needs to do is find the state of this
            ' layer *prior to the change* - and now that we know which layer is affected, we can do that!
            'If (m_UndoStack(i).srcProcCall.pcUndoType = UNDO_Layer) Or (m_UndoStack(i).srcProcCall.pcUndoType = UNDO_Layer_VectorSafe) Then
            
            'New plan!  Use the action database to determine fadeability for us.
            If Actions.IsActionFadeable(m_UndoStack(i).srcProcCall.pcID, True) Then
            
                'Certain actions are irrelevant to this function; in particular, "Fade" and "Repeat" cannot be
                ' "faded" or "repeated".  In these instances, we want to continue searching backward, to find
                ' the original function that was faded or repeated.
                If Strings.StringsNotEqual(m_UndoStack(i).srcProcCall.pcID, "Fade", True) Then
                    If Strings.StringsNotEqual(m_UndoStack(i).srcProcCall.pcID, "Repeat", True) Then
                        actionFound = True
                        affectedLayerID = m_UndoStack(i).undoLayerID
                        nameOfProcessorAction = m_UndoStack(i).srcProcCall.pcID
                        layerChangeIndex = i
                        Exit For
                    End If
                End If
                
            End If
        
        Next i
        
        'If a relevant layer was found, actionFound will be set to TRUE.
        If actionFound Then
            
            FillDIBWithLastUndoCopy = True
            
            'Some functions, like SyncInterfaceToCurrentImage, only care if a recent Layer modification exists.
            ' They don't actually need us to return the relevant DIB.  If checkStateOnly = TRUE, exit now.
            If (Not checkStateOnly) Then
                
                'Armed with the knowledge of which layer is being affected, we now need to find the last relevant
                ' DIFF index for that layer.
                Dim relevantDiffIndex As Long
                relevantDiffIndex = FindNearestDiffIndex(layerChangeIndex, UNDO_Layer, affectedLayerID)
                
                'Check also for vector-safe actions, which are identical to UNDO_LAYER but for operations that
                ' support vector-safe modifications.
                Dim relevantVectorIndex As Long
                relevantVectorIndex = FindNearestDiffIndex(layerChangeIndex, UNDO_Layer_VectorSafe, affectedLayerID)
                
                'relevantDiffIndex will now contain the index in the m_undoStack array corresponding to the
                ' previous state of the affected layer.
                '
                'Load the relevant portion of that Undo file into the target DIB.
                Dim tmpLayer As pdLayer
                Set tmpLayer = New pdLayer
                
                If (relevantVectorIndex > relevantDiffIndex) Then
                    ImageImporter.LoadUndo GenerateUndoFilename(relevantVectorIndex), m_UndoStack(relevantVectorIndex).srcProcCall.pcUndoType, UNDO_Layer_VectorSafe, affectedLayerID, True, tmpLayer
                Else
                    ImageImporter.LoadUndo GenerateUndoFilename(relevantDiffIndex), m_UndoStack(relevantDiffIndex).srcProcCall.pcUndoType, UNDO_Layer, affectedLayerID, True, tmpLayer
                End If
                
                'tmpLayer will now contain a copy of the relevant DIB (inside .GetLayerDIB).  Copy it into the destination DIB, then exit.
                Set dstDIB = tmpLayer.GetLayerDIB
                Set tmpLayer = Nothing
                
            End If
            
        '/(If a relevant layer was not found, provide a default FALSE return and perform no further processing)
        End If
        
    End If
    
End Function

'Get the filename of the current Undo file, using the same formula as GenerateUndoFilename() below
Friend Function GetCurrentUndoFilename() As String
    GetCurrentUndoFilename = GenerateUndoFilename(m_UndoNum)
End Function

'Generate a unique Undo/Redo file name, including full temp path.
Private Function GenerateUndoFilename(ByVal uIndex As Long) As String
    
    'As a failsafe, all Undo/Redo files are erased when this class is unloaded.  The ClearUndos function requires this
    ' function to generate Undo filenames.  If our parent class was terminated properly, it will have already erased
    ' our reference to it, so the parentPDImage.GetUniqueID reference below will fail - to prevent that, we must check the
    ' parent reference in advance.
    If (Not parentPDImage Is Nothing) Then
        GenerateUndoFilename = UserPrefs.GetTempPath & "~cPDU_(" & parentPDImage.GetUniqueID & ")_" & uIndex & ".pdtmp"
    End If
    
End Function

'Friendly accessible mechanism for generating a unique Undo/Redo file name.  Any changes to the private function, above,
' should be mirrored here.
Friend Function GenerateUndoFilenameExternal(ByRef uniqueID As String, ByVal uIndex As Long) As String
    GenerateUndoFilenameExternal = UserPrefs.GetTempPath & "~cPDU_(" & uniqueID & ")_" & uIndex & ".pdtmp"
End Function

'Generate a unique Undo/Redo header filename, including full temp path.  Note that this function does *NOT* take an
' Undo number index.  The entire stack is written out to the header file, every time.
Private Function GenerateUndoHeaderFilename() As String
    
    If (LenB(m_HeaderFilename) = 0) Then
    
        'As a failsafe, all Undo/Redo files are erased when this class is unloaded.  The ClearUndos function requires this
        ' function to generate Undo filenames.  If our parent class was terminated properly, it will have already erased
        ' our reference to it, so the parentPDImage.GetUniqueID reference below will fail - to prevent that, we must check the
        ' parent reference in advance.
        If (Not parentPDImage Is Nothing) Then
            m_HeaderFilename = UserPrefs.GetTempPath() & "~PDU_StackSummary_(" & parentPDImage.GetUniqueID & ").pdtmp"
        End If
        
    End If
    
    GenerateUndoHeaderFilename = m_HeaderFilename
    
End Function

'Erase all undo files that this class has created
Friend Sub ClearUndos()
    
    'Remember that we have no Undo files for this image
    m_UndoState = False
    
    'If Undo data has already been cleared, exit
    If (m_UndoTotal = 0) Then Exit Sub
    
    '99.9% of time we now want to wipe all Undos associated with this class.  There's a single
    ' exception to this, however - if a system-initiated shutdown is occurring and the user wants
    ' us to auto-restore this session after a reboot.  If that happens, we *must* preserve Undo data
    ' since it's a key component of PD's Autosave system.
    Dim suspendClearUndos As Boolean: suspendClearUndos = False
    If g_ProgramShuttingDown Then
        If (Not g_ThunderMain Is Nothing) Then suspendClearUndos = g_ThunderMain.WasEndSessionReceived(True)
    End If
    
    If suspendClearUndos Then
        PDDebug.LogAction "Suspending Undo clear due to session restoration setting"
        Exit Sub
    End If
    
    'Destroy every potential Undo file
    Dim i As Long
    For i = 0 To m_UndoTotal
        
        'Check image data first...
        Files.FileDeleteIfExists GenerateUndoFilename(i)
        
        '...followed by layer data
        Files.FileDeleteIfExists GenerateUndoFilename(i) & ".layer"
        
        '...followed by selection data
        Files.FileDeleteIfExists GenerateUndoFilename(i) & ".selection"
        
    Next i
    
    'Destroy the header file and preview image, if any
    If (m_HeaderXMLHandle <> 0) Then m_FSO.FileCloseHandle m_HeaderXMLHandle
    Files.FileDeleteIfExists GenerateUndoHeaderFilename()
    If (m_ThumbnailFileHandle <> 0) Then m_FSO.FileCloseHandle m_ThumbnailFileHandle
    Files.FileDeleteIfExists GenerateUndoHeaderFilename() & ".pdasi"
    Erase m_UTF8
    
    'Set the current and maximum Undo number to 0
    m_UndoNum = 0
    m_UndoMax = 0
    m_UndoTotal = 0
    
End Sub

'Retrieve a copy of an original, unmodified layer using PhotoDemon's Undo engine.  This is currently used by
' the "Export > Color lookup" feature (which compares current image state to original image state, then tries to
' generate a matching 3D LUT defining any changes that have taken place).
'
'This function will return FALSE if no layer-specific changes have been made, and thus retrieval is not relevant.
Friend Function GetOriginalLayer_FromUndo(ByRef dstDIB As pdDIB, ByVal targetLayerID As Long) As Boolean
    
    GetOriginalLayer_FromUndo = False
    
    'Ensure the image has changed
    If (m_UndoNum <> 0) Then
            
        GetOriginalLayer_FromUndo = True
        
        'We know the original Undo file is of type UNDO_EVERYTHING (because a full copy of the base image
        ' was saved at the head of the Undo chain).  Because of this, we don't need to invoke the full
        ' Undo parser - instead, just directly call the retrieval function we know we need.
        Dim tmpLayer As pdLayer
        Set tmpLayer = New pdLayer
        GetOriginalLayer_FromUndo = ImageImporter.LoadPDI_SingleLayer(GenerateUndoFilename(0), tmpLayer, targetLayerID, False)
        
        'On success, tmpLayer will now contain a copy of the relevant DIB (inside .GetLayerDIB).
        ' Copy it into the destination DIB before exiting.
        If GetOriginalLayer_FromUndo Then
            Set dstDIB = tmpLayer.GetLayerDIB
            Set tmpLayer = Nothing
        End If
        
    End If
    
End Function

'Return the tool associated with the current Undo state
Friend Function GetUndoTool() As Long
    If (m_UndoNum > 0) Then GetUndoTool = m_UndoStack(m_UndoNum - 1).srcProcCall.pcTool
End Function

'Return the process type associated with the current Undo state
Friend Function GetUndoProcessType() As PD_UndoType
    If m_UndoState Then
        GetUndoProcessType = m_UndoStack(m_UndoNum - 1).srcProcCall.pcUndoType
    Else
        GetUndoProcessType = UNDO_Nothing
    End If
End Function

'Return the process type associated with the current Redo state
Friend Function GetRedoProcessType() As PD_UndoType
    If m_RedoState Then
        GetRedoProcessType = m_UndoStack(m_UndoNum).srcProcCall.pcUndoType
    Else
        GetRedoProcessType = UNDO_Nothing
    End If
End Function

'Return the processID associated with the current Undo state
Friend Function GetUndoProcessID() As String
    If m_UndoState Then
        GetUndoProcessID = g_Language.TranslateMessage(m_UndoStack(m_UndoNum - 1).srcProcCall.pcID)
    Else
        GetUndoProcessID = vbNullString
    End If
End Function

'Return the processID associated with the current Redo state
Friend Function GetRedoProcessID() As String
    If m_RedoState Then
        GetRedoProcessID = g_Language.TranslateMessage(m_UndoStack(m_UndoNum).srcProcCall.pcID)
    Else
        GetRedoProcessID = vbNullString
    End If
End Function

'Given a type of Undo action, find the last relevant parameter string.  This is used by PD's central processor to detect on-canvas
' changes to the image since the last Undo was created.
Friend Function GetLastParamString(ByVal requestedUndoType As PD_UndoType) As String

    GetLastParamString = vbNullString
    
    'If no Undos have been created, exit immediately
    If (m_UndoNum = 0) Then Exit Function
    
    'Loop backward through the Undo stack, looking for a matching Undo type.  If found, return the associated
    ' param string and exit immediately.
    Dim i As Long
    
    For i = m_UndoNum - 1 To 0 Step -1
        If (m_UndoStack(i).srcProcCall.pcUndoType = requestedUndoType) Then
            GetLastParamString = m_UndoStack(i).srcProcCall.pcParameters
            Exit For
        End If
    Next i
    
End Function

'Undo/redo data handling is complex.  Sometimes it is helpful to see how our Undo/Redo stack looks - use this sub
' to print out a full list to the debug window.
Friend Sub DebugUndo(ByVal sourceOfDebugRequest As String)
    
    Dim cString As pdString
    Set cString = New pdString
    
    Dim i As Long
    Debug.Print "---"
    Debug.Print " Current Undo stack:"
    Debug.Print "  (Displaying because " & sourceOfDebugRequest & ")"
    For i = 0 To UBound(m_UndoStack)
        cString.Reset
        If (i = m_UndoNum - 1) Then cString.Append "  * " Else cString.Append "  "
        cString.Append ": "
        cString.Append m_UndoStack(i).srcProcCall.pcID
        cString.Append ", T:"
        cString.Append CStr(m_UndoStack(i).srcProcCall.pcUndoType)
        cString.Append ", L:"
        cString.Append CStr(m_UndoStack(i).undoLayerID)
        Debug.Print cString.ToString()
    Next i
    Debug.Print "---"

End Sub

'Write the Undo stack out to file, in XML format, with an attached thumbnail image as well (for AutoSave purposes)
Friend Sub WriteUndoDataToXML(ByRef dstFilename As String)
    
    'The undo/redo engine is the source of a lot of processing time, particularly when painting.  In an attempt to
    ' better profile it, detailed timing measurements are currently reported in nightly builds.
    Dim startTime As Currency
    VBHacks.GetHighResTime startTime
    
    'Create a pdXML class, which will help us assemble the XML file
    Dim xmlEngine As pdXML
    Set xmlEngine = New pdXML
    xmlEngine.PrepareNewXML "Undo stack"
    
    'Write out the XML version we're using for this macro
    xmlEngine.WriteTag "pdUndoVersion", PD_UNDO_VERSION_ID
    xmlEngine.WriteBlankLine
    
    'Write out some basic information about our parent pdImage object.  In the event of a critical program crash,
    ' the AutoSave engine can use this information to reconstruct any missing images.
    xmlEngine.WriteTag "imageID", parentPDImage.GetUniqueID
    
    Dim friendlyName As String
    friendlyName = parentPDImage.ImgStorage.GetEntry_String("OriginalFileName", vbNullString)
    If (LenB(parentPDImage.ImgStorage.GetEntry_String("OriginalFileExtension", vbNullString)) > 0) Then friendlyName = friendlyName & "." & parentPDImage.ImgStorage.GetEntry_String("OriginalFileExtension", vbNullString)
    
    xmlEngine.WriteTag "friendlyName", friendlyName
    xmlEngine.WriteTag "originalPath", parentPDImage.ImgStorage.GetEntry_String("CurrentLocationOnDisk", vbNullString)
    xmlEngine.WriteTag "originalFormat", parentPDImage.GetOriginalFileFormat()
    xmlEngine.WriteTag "currentFormat", parentPDImage.GetCurrentFileFormat()
    xmlEngine.WriteBlankLine
    
    'Write out the size of the current Undo stack, including the Max and Total values
    xmlEngine.WriteTag "StackHeight", m_UndoMax
    xmlEngine.WriteTag "StackAbsoluteMaximum", m_UndoTotal
    xmlEngine.WriteTag "CurrentStackPointer", m_UndoNum
    
    'We probably don't need the m_UndoNumAtLastSave values, but it doesn't hurt to embed them, "just in case"
    xmlEngine.WriteTag "UndoNumAtLastSave_Flat", m_UndoNumAtLastSave_Flat
    xmlEngine.WriteTag "UndoNumAtLastSave_PDI", m_UndoNumAtLastSave_PDI
    xmlEngine.WriteBlankLine
    
    'Next, we're going to write out all Undo stack entries. If PD's "PD_UndoEntry" type is ever changed, note
    ' that this function will need to be updated to match!
    Dim i As Long
    For i = 0 To UBound(m_UndoStack)
        
        'Start each undo entry with a unique numeric identifier (that corresponds to its location in the stack)
        xmlEngine.WriteTagWithAttribute "UndoEntry", "index", i, vbNullString, True
        
        'Write out all the properties of this entry
        With m_UndoStack(i)
            xmlEngine.WriteTag "ProcessID", .srcProcCall.pcID
            xmlEngine.WriteTag "ProcessParameters", .srcProcCall.pcParameters
            xmlEngine.WriteTag "UndoType", .srcProcCall.pcUndoType
            xmlEngine.WriteTag "RelevantTool", .srcProcCall.pcTool
            xmlEngine.WriteTag "RaiseDialog", .srcProcCall.pcRaiseDialog
            xmlEngine.WriteTag "Recorded", .srcProcCall.pcRecorded
            xmlEngine.WriteTag "RelevantLayerID", .undoLayerID
        End With
        
        'Close this undo entry
        xmlEngine.CloseTag "UndoEntry"
        xmlEngine.WriteBlankLine
    
    Next i
    
    If UNDO_DETAILED_REPORTS Then
        PDDebug.LogAction "Undo engine XML generation took " & VBHacks.GetTimeDiffNowAsString(startTime)
        VBHacks.GetHighResTime startTime
    End If
    
    'With all Undo data successfully written, close the XML data and write it out to file.
    ' (Note that a custom, high-speed approach is used for this.)
    If (m_HeaderXMLHandle = 0) Then m_FSO.FileCreateHandle dstFilename, m_HeaderXMLHandle, True, True, OptimizeTempFile
    If (m_HeaderXMLHandle <> 0) Then

        'Convert the source string to UTF-8, and leave three open spaces for the BOM at the front
        Dim utf8Len As Long
        Strings.UTF8FromStrPtr xmlEngine.GetXMLStringPtr, -1, m_UTF8, utf8Len, 3
        
        'Add the required UTF-8 BOM.  (Required by PD, anyway - it makes text file reading much faster, as we don't have to
        ' guess at encodings.)
        m_UTF8(0) = &HEF: m_UTF8(1) = &HBB: m_UTF8(2) = &HBF
        
        'Overwrite our existing XML file.  Note that we use utf8Len + 2 instead of +3, because the UTF8FromStrPtr call above
        ' returns a length that *includes* the trailing null char (which we obviously don't want to write to file).
        m_FSO.FileSave_FastOverwrite m_HeaderXMLHandle, VarPtr(m_UTF8(0)), utf8Len + 2

    End If
    
    If UNDO_DETAILED_REPORTS Then
        PDDebug.LogAction "Undo engine XML file write took " & VBHacks.GetTimeDiffNowAsString(startTime)
        VBHacks.GetHighResTime startTime
    End If
    
    'As an added bonus (new to PD 6.4), also write out a thumbnail of the image *as it appears right now*.  The Autosave engine
    ' will use this as an image preview if something goes horribly wrong.
    If (m_UndoNum > 0) Then
        If (Not m_UndoStack(m_UndoNum - 1).thumbnailLarge Is Nothing) And (Not m_UndoStack(m_UndoNum - 1).srcProcCall.pcUndoType = UNDO_Selection) Then
            
            'For performance reasons, we use a persistent file handle
            If (m_ThumbnailFileHandle = 0) Then m_FSO.FileCreateHandle GenerateUndoHeaderFilename() & ".pdasi", m_ThumbnailFileHandle, True, True, OptimizeTempFile
            If (m_ThumbnailFileHandle <> 0) Then
                
                'If this image was created from an Autosave instance (e.g. after a program crash), thumbnails won't exist.
                ' That's okay - but we don't want to try and write a thumbnail file if this occurs.
                If (m_UndoStack(m_UndoNum - 1).thumbnailLarge.GetDIBWidth > 0) Then
                    m_FSO.FileMovePointer m_ThumbnailFileHandle, 0, FILE_BEGIN
                    If m_UndoStack(m_UndoNum - 1).thumbnailLarge.WriteToFile(vbNullString, cf_Lz4, True, m_ThumbnailFileHandle) Then m_FSO.FileSetEOF m_ThumbnailFileHandle
                End If
                
            End If
            
        End If
    End If
    
    If UNDO_DETAILED_REPORTS Then PDDebug.LogAction "Undo engine thumbnail write took " & VBHacks.GetTimeDiffNowAsString(startTime)
    
End Sub

'Given an XML string generated by another pdUndo instance, reconstruct this stack to match.  This is used by the Autosave engine
' to restore a full previous session.
Friend Function ReconstructStackFromExternalSource(ByRef srcXMLString As String) As Boolean

    'Initiate an XML parser
    Dim xmlEngine As pdXML
    Set xmlEngine = New pdXML
    
    'Validate the XML string
    If xmlEngine.LoadXMLFromString(srcXMLString) Then
    
        'Make sure the XML string actually contains Undo/Redo data
        If xmlEngine.IsPDDataType("Undo stack") And xmlEngine.ValidateLoadedXMLData("pdUndoVersion") Then
            
            'The file checks out!  Start by initiating a few module-level values to match the data in the file.
            m_UndoMax = xmlEngine.GetUniqueTag_Long("StackHeight")
            m_UndoTotal = xmlEngine.GetUniqueTag_Long("StackAbsoluteMaximum")
            m_UndoNum = xmlEngine.GetUniqueTag_Long("CurrentStackPointer")
            m_UndoNumAtLastSave_Flat = xmlEngine.GetUniqueTag_Long("UndoNumAtLastSave_Flat")
            m_UndoNumAtLastSave_PDI = xmlEngine.GetUniqueTag_Long("UndoNumAtLastSave_PDI")
            
            'Next, we must reconstruct the Undo stack itself.  Start by preparing the stack array.
            ReDim m_UndoStack(0 To m_UndoTotal) As PD_UndoEntry
            
            'Start retrieving individual stack entries from the file
            Dim i As Long
            For i = 0 To UBound(m_UndoStack)
                
                'Start by finding the location of the tag we want
                Dim tagPosition As Long
                tagPosition = xmlEngine.GetLocationOfTagPlusAttribute("UndoEntry", "index", i)
                
                'If a matching tag was found, retrieve its data and store it in the stack
                If (tagPosition > 0) Then
                
                    With m_UndoStack(i)
                        .srcProcCall.pcID = xmlEngine.GetUniqueTag_String("ProcessID", , tagPosition)
                        .srcProcCall.pcParameters = xmlEngine.GetUniqueTag_String("ProcessParameters", , tagPosition)
                        .srcProcCall.pcUndoType = xmlEngine.GetUniqueTag_Long("UndoType", , tagPosition)
                        .srcProcCall.pcTool = xmlEngine.GetUniqueTag_Long("RelevantTool", , tagPosition)
                        .srcProcCall.pcRaiseDialog = xmlEngine.GetUniqueTag_Boolean("RaiseDialog", , tagPosition)
                        .srcProcCall.pcRecorded = xmlEngine.GetUniqueTag_Boolean("Recorded", , tagPosition)
                        .undoLayerID = xmlEngine.GetUniqueTag_Long("RelevantLayerID", , tagPosition)
                    End With
                    
                    'Initialize the thumbnail DIB, even though it won't contain anything useful.
                    Set m_UndoStack(i).thumbnailLarge = New pdDIB
                
                End If
                
            Next i
            
            'The stack was reconstructed successfully.
            ReconstructStackFromExternalSource = True
            
        Else
            ReconstructStackFromExternalSource = False
        End If
    
    Else
        ReconstructStackFromExternalSource = False
    End If

End Function

'External functions can request a copy of the Undo stack by calling this function and supply the three required parameters
' (all passed ByRef, obviously):
' 1) Number of Undo entries
' 2) Current Undo entry
' 3) Array of type PD_UndoEntry, to receive the full contents of the stack
Friend Sub CopyUndoStack(ByRef numOfUndos As Long, ByRef curUndoNumber As Long, ByRef stackList() As PD_UndoEntry)

    numOfUndos = m_UndoMax
    curUndoNumber = m_UndoNum
    
    ReDim stackList(0 To m_UndoMax) As PD_UndoEntry
    
    Dim i As Long
    For i = 0 To m_UndoMax - 1
        
        With stackList(i)
            
            .srcProcCall = m_UndoStack(i).srcProcCall
            .undoLayerID = m_UndoStack(i).undoLayerID
            
            Set .thumbnailLarge = New pdDIB
            .thumbnailLarge.CreateFromExistingDIB m_UndoStack(i).thumbnailLarge
            
        End With
        
    Next i

End Sub

Private Sub Class_Initialize()
    
    'Initialize all relevant undo/redo tracking variables
    ReDim m_UndoStack(0) As PD_UndoEntry
    
    m_UndoNum = 0
    m_UndoMax = 0
    m_UndoNumAtLastSave_Flat = 0
    m_UndoNumAtLastSave_PDI = 0
    m_UndoState = False
    m_RedoState = False
    
    'A private FSO instance and UTF-8 buffer are now used, for improved performance.
    Set m_FSO = New pdFSO
    
    'We also track the export-time taken by Undo type.  If Undos are taking too long to generate, we crank up lz4 acceleration
    ' to compensate.
    ReDim m_UndoAcceleration(0 To UNDO_Selection) As Long
    
    'By default, different undo types save at different compression levels.  "-1" indicates default acceleration level.
    ' Higher numbers indicate faster-but-worse compression.  (Undo types that require more work, e.g. saving the full
    ' layer stack, as required before "Flattening", are compressed with higher acceleration levels by default - but note
    ' that PD's undo engine dynamically adjusts these levels during a given session to arrive at idealized save times.)
    Dim i As Long
    For i = 0 To UNDO_Selection
        
        'Start with a faster default compression speed for large undo types
        If (i = UNDO_Everything) Or (i = UNDO_Image) Then
            m_UndoAcceleration(i) = 100
        
        'For layer-based undo types (as used by things like paint strokes), use a slightly slower-but-better option
        ElseIf (i = UNDO_Layer) Then
            m_UndoAcceleration(i) = 50
        
        'Everything else gets maximum compression by default
        Else
            m_UndoAcceleration(i) = -1
        
        End If
        
    Next i
    
End Sub

'When this class is unloaded, we must make sure to delete any Undo files we created
Private Sub Class_Terminate()
    
    If (m_HeaderXMLHandle <> 0) Then
        m_FSO.FileCloseHandle m_HeaderXMLHandle
        m_HeaderXMLHandle = 0
    End If
    
    If (m_ThumbnailFileHandle <> 0) Then
        m_FSO.FileCloseHandle m_ThumbnailFileHandle
        m_ThumbnailFileHandle = 0
    End If
    
    Me.ClearUndos
    
End Sub
