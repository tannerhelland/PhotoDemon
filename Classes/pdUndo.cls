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
'Copyright 2001-2015 by Tanner Helland
'Created: 2/April/01
'Last updated: 15/October/14
'Last update: add new forceLastUndoDataToIncludeEverything() to work around some Crop tool issues.
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
' entry.  This allows for very tiny Undo/Redo files, at a trade-off with a large amount of new complexity in the
' Undo/Redo engine itself.  I believe this approach is the best one possible, so I'm hoping 2014 marks the end of
' massive Undo/Redo engine rewrites.  (Fingers crossed.)
'
'Undo/Redo data is still saved to the HDD, and will remain that way for the forseeable future.  Unlimited operations
' are allowed.  This class does not check HDD space availability, so errors may be thrown if the user runs out of free
' space.  Similarly, clearing the specified TEMP folder in the middle of program operation is a bad idea.  PD will not
' crash, but as expected, Undo/Redo functionality will cease to work.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Undo/Redo currently works as a stack, and entries can only be traversed sequentially.  In the future, I plan to
' rework this to allow the program to "jump" to any Undo entry arbitrarily (e.g. as part of an Undo History browser).
' To that end, I am implementing Undo/Redo tracking in a more comprehensive way than may seem necessary for the
' current stack-only implementation.

'In the Undo/Redo stack, which entry is this image currently on?
Private undoNum As Long

'How many Undo/Redo entries exist in our stack?
Private undoMax As Long
Private undoTotal As Long

'Are Undo/Redo operations allowed for this image?  (This is used to determine the state of undo/redo buttons and
' menu items.)
Private undoState As Boolean
Private redoState As Boolean

'Note that the UndoType struct is now declared publicly, so you can find it in the Public_Enums_and_Types module.
' (Declaration mirrored here for clarity)

'Information about each Undo entry is stored in an array; the array is dynamically resized as necessary when new
' Undos are created.  We track the ID of each action in preparation for a future History browser that allows the
' user to jump to any arbitrary Undo/Redo state.  (Also, to properly update the text of the Undo/Redo menu and
' buttons so the user knows which action they are undo/redoing.)

' Public Type undoEntry
'     processID As String             'Name of the associated action (e.g. "Gaussian blur")
'     processParamString As String    'Processor string supplied to the action
'     undoType As PD_UNDO_TYPE        'What type of Undo/Redo data was stored for this action (e.g. Image or Selection data)
'     undoLayerID As Long             'If the undoType is UNDO_LAYER, UNDO_LAYER_VECTORSAFE, or UNDO_LAYERHEADER, this value will note the ID (NOT THE INDEX) of the affected layer
'     relevantTool As Long            'If a tool was associated with this action, it can be set here.  This value is not currently used.
'     thumbnailSmall As pdDIB         'A small thumbnail associated with the current action.  In the future, this will be used by the Undo History window.
'     thumbnailLarge As pdDIB         'A large thumbnail associated with the current action.
' End Type

Private undoStack() As undoEntry

'If the user performs 10 actions, then presses Undo 5 times and saves the image, we need to disable the Save button
' at that Undo point, because it was the last point at which the image was saved.  Thus we must track the last point
' in the Undo/Redo stack where the image was saved.  Note that we separately track flat image saves (JPEG, PNG, etc),
' and layered image saves (PDI).  External functions can request either or both of these values as they need 'em.
Private undoNumAtLastSave_Flat As Long, undoNumAtLastSave_PDI As Long

'We need access to the parent class so we can request things like "write your current data out to file".
Public parentPDImage As pdImage

'As part of PD's Undo/Redo chain, we can write the Undo/Redo stack to an XML file.  These constants are used to help
' us ID and parse such data correctly.
Private Const pdUndoVersionID As Long = 64

'The size of the Undo/Redo Thumbnail
Private Const UNDO_THUMB_SMALL As Long = 48
Private Const UNDO_THUMB_LARGE As Long = 256

'When an image is saved, this class needs to be notified so that it can make a note of the current undo index
Public Sub notifyImageSaved(ByVal typeOfSave As PD_SAVE_EVENT)
    
    Dim saveIndex As Long
    
    If undoNum > 0 Then
        saveIndex = undoNum
    Else
        saveIndex = 1
    End If
    
    Select Case typeOfSave
        
        'This value should never be passed in!
        Case pdSE_AnySave
            Debug.Print "The pdSE_AnySave type is only for GETTING image save state, not SETTING it!  Fix this!"
            undoNumAtLastSave_Flat = saveIndex
            undoNumAtLastSave_PDI = saveIndex
                
        'Image has been saved to PDI format, meaning layers are intact
        Case pdSE_SavePDI
            undoNumAtLastSave_PDI = saveIndex
                
        'Image has been saved to some flat format (JPEG, PNG, etc).  Layers are not intact.
        Case pdSE_SaveFlat
            undoNumAtLastSave_Flat = saveIndex
        
    End Select
    
End Sub

Public Function getUndoState() As Boolean
    getUndoState = undoState
End Function

Public Function getRedoState() As Boolean
    getRedoState = redoState
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
Public Sub createUndoData(ByVal processID As String, ByVal pParamString As String, ByVal processType As PD_UNDO_TYPE, Optional ByVal targetLayerID As Long = -1, Optional ByVal curTool As Long = -1, Optional ByVal coalescenceAllowed As Boolean = False)
    
    'If the caller is okay with us coalescing this into the previous entry, test for coalescence acceptability
    If coalescenceAllowed And (undoNum > 0) Then
        
        'If this processID and layerID match the previous ones, we are okay to just overwrite that entry with this one.
        If (StrComp(processID, undoStack(undoNum - 1).processID, vbBinaryCompare) = 0) And (targetLayerID = undoStack(undoNum - 1).undoLayerID) Then
            
            'As an additional failsafe, make sure the Undo type is layer-safe, and vector-safe.  This guarantees that only
            ' non-destructive changes are allowed to coalesce.
            If (processType = UNDO_LAYER_VECTORSAFE) And (undoStack(undoNum - 1).undoType = UNDO_LAYER_VECTORSAFE) Then
                undoNum = undoNum - 1
            End If
            
        End If
        
    End If
    
    'Forward a save request to the actual saveUndoData function.  It handles the messy business of actually writing the correct
    ' filetypes (which it determines based on the current processType and targetLayer)
    Saving.saveUndoData parentPDImage, getCurrentUndoFilename(), processType, targetLayerID
    
    'Previously, we notified the user when undo/redo data was saved.  This was stupid, but it is sometimes helpful to
    ' display that info in the debug window.
    'Debug.Print "Saved " & getCurrentUndoFilename() & " as Undo file #" & undoNum & "."
    
    'Store all relevant information about this Undo action, so we can access it in the future if necessary
    If undoNum <> UBound(undoStack) Then ReDim Preserve undoStack(0 To undoNum) As undoEntry
    With undoStack(undoNum)
        .processID = processID
        .processParamString = pParamString
        .undoType = processType
        .undoLayerID = targetLayerID
        .relevantTool = curTool
    End With
        
    'If the present image is NOT the last image in the undo train (i.e., the user pressed Undo before running
    ' the current operation), treat this action as the new "last" Undo entry in the stack.  However, we still need
    ' to remember that Undo files exist past this point, so we can clean them up when we exit - hence the need for
    ' a separate undoTotal variable.
    undoMax = undoNum
    
    'Next, we will store a copy of the current image thumbnail in two sizes: 256x256, and 48x48
    
    'First, obtain a large copy of the thumbnail.  The source of the thumbnail changes depending on the undo type.
    Set undoStack(undoNum).thumbnailLarge = New pdDIB
    
    Select Case processType
    
        Case UNDO_NOTHING
            parentPDImage.requestThumbnail undoStack(undoNum).thumbnailLarge, FixDPI(UNDO_THUMB_LARGE)
        
        Case UNDO_EVERYTHING
            parentPDImage.requestThumbnail undoStack(undoNum).thumbnailLarge, FixDPI(UNDO_THUMB_LARGE)
            
        Case UNDO_IMAGE, UNDO_IMAGE_VECTORSAFE
            parentPDImage.requestThumbnail undoStack(undoNum).thumbnailLarge, FixDPI(UNDO_THUMB_LARGE)
        
        Case UNDO_IMAGEHEADER
            parentPDImage.requestThumbnail undoStack(undoNum).thumbnailLarge, FixDPI(UNDO_THUMB_LARGE)
        
        Case UNDO_LAYER, UNDO_LAYER_VECTORSAFE
            parentPDImage.getLayerByID(targetLayerID).requestThumbnail undoStack(undoNum).thumbnailLarge, FixDPI(UNDO_THUMB_LARGE)
        
        Case UNDO_LAYERHEADER
            parentPDImage.requestThumbnail undoStack(undoNum).thumbnailLarge, FixDPI(UNDO_THUMB_LARGE)
        
        Case UNDO_SELECTION
            parentPDImage.mainSelection.requestThumbnail undoStack(undoNum).thumbnailLarge, FixDPI(UNDO_THUMB_LARGE)
            
    End Select
    
    'Next, create a smaller copy of the thumbnail using FreeImage resize.  (This is much faster than regenerating
    ' the thumbnail from scratch, especially if an image has many layers.)
    Set undoStack(undoNum).thumbnailSmall = New pdDIB
    undoStack(undoNum).thumbnailSmall.createBlank FixDPI(UNDO_THUMB_SMALL), FixDPI(UNDO_THUMB_SMALL), 32, 0
    FreeImageResizeDIBFast undoStack(undoNum).thumbnailSmall, 0, 0, undoStack(undoNum).thumbnailSmall.getDIBWidth, undoStack(undoNum).thumbnailSmall.getDIBHeight, undoStack(undoNum).thumbnailLarge, FILTER_BILINEAR, True
    
    'Increase our Undo count
    undoNum = undoNum + 1
    undoMax = undoMax + 1
    
    'If the present image is *truly* the last one in the Undo train, note that we have more Undo files to delete
    ' when our parent is unloaded.
    If undoMax > undoTotal Then undoTotal = undoMax
    
    'If this is not the first Undo entry, make a note that the user is now allowed to initiate Undo actions.
    If undoNum > 1 Then undoState = True
    
    'Because we have just created Undo data, the user is not allowed to initiate Redo actions.
    redoState = False
    
    'If this action changed pixel data, mark the image as NOT being saved in its current state.
    If (processType <> UNDO_NOTHING) And (processType <> UNDO_SELECTION) And (undoNum > 1) Then
        parentPDImage.setSaveState False, pdSE_AnySave
    End If
    
    'New to version 6.4 is our persistent creation of Undo XML header files.  These are extraordinarily helpful
    ' for both regular Undo/Redo behavior, as well as AutoSave recreation.  We always create a copy of the Undo
    ' stack, in XML format, in the temporary folder.
    writeUndoDataToXML generateUndoHeaderFilename()
    
    'To inspect the current state of the Undo/Redo stack, uncomment this line.
    'debugUndo "new Undo entry created"
    
End Sub

'This (rather odd) function is required by the Crop tool.  Because the Crop tool forcibly removes the current selection - but only after
' making use of it - it causes the Undo/Redo engine's selection tracking to fall out of sync.  To work around this without adding
' unnecessary dummy entries to the Undo chain, the Crop command calls this function prior to processing.
'
'This function simply updates the previous Undo entry to include a full record of everything: the whole image, and any
' active selections.  This ensures that if the user uses Undo to return to a pre-Crop state, everything is restored properly.
' (Without this function, Undo-ing a crop action would cause the selection to be lost unless the user returned to the point when the
' selection was first created.)
Public Sub forceLastUndoDataToIncludeEverything()
    
    'Overwrite the previous action's Undo data with a full image backup.  Note that all the settings from the previous
    ' action are retained; the only thing we're doing differently is specifying the Undo type as UNDO_EVERYTHING.
    Saving.saveUndoData parentPDImage, generateUndoFilename(undoNum - 1), UNDO_EVERYTHING, -1
    
    'Update the stack entry at that point to reflect the UNDO_EVERYTHING usage
    With undoStack(undoNum - 1)
        .undoType = UNDO_EVERYTHING
        .undoLayerID = -1
    End With
    
    'Update the Undo XML header as well.
    writeUndoDataToXML generateUndoHeaderFilename()
        
End Sub

'This special command can be used to revert the image to its last saved state.  For images that have never been
' saved (e.g. those pasted from the clipboard, etc), it will undo to the image's original state.
Public Sub revertToLastSavedState()
    
    Dim targetUndoPoint As Long
    
    If undoNumAtLastSave_Flat > undoNumAtLastSave_PDI Then
        targetUndoPoint = undoNumAtLastSave_Flat
    Else
        targetUndoPoint = undoNumAtLastSave_PDI
    End If
    
    'Failsafe check to make sure the specified index is valid
    If targetUndoPoint < 1 Then targetUndoPoint = 1
    
    moveToSpecificUndoPoint targetUndoPoint
    
End Sub

'Move to a specific location in the Undo/Redo chain.  Note that this function may take some time if the requested
' point is quite far into the Undo/Redo chain.
Public Sub moveToSpecificUndoPoint(ByVal undoIndex As Long)
    
    'Some external actions must be deactivated while Undo/Redo processing occurs.  Such asynchronous functions can check
    ' this global variable to determine if they should fire, so we must set it before doing anything else.
    g_UndoRedoActive = True
    
    Message "Reconstructing image..."
    
    'Start by resetting the image to its default state, e.g. Undo 0.
    LoadUndo generateUndoFilename(0), UNDO_EVERYTHING, UNDO_EVERYTHING, , True
    
    'From this point, we now want to advance the Undo stack to the requested index by applying successive diff entries
    ' until the final index is reached.
    If undoIndex > 1 Then
    
        Dim i As Long
        For i = 1 To undoIndex - 1
            LoadUndo generateUndoFilename(i), undoStack(i).undoType, undoStack(i).undoType, undoStack(i).undoLayerID, True
        Next i
    
    End If
    
    'Our parent pdImage object has now been successively reconstructed.  Set all Undo variables to their proper values, then
    ' redraw the screen and exit.
    
    'undoNum is always one higher than the state of the image.  (This is easily remembered because the base Undo copy is created at
    ' point 0 right when an image is loaded, and because the undoNum is immediately advanced, it starts at value 1.)
    undoNum = undoIndex
    
    'Allow saving contingent on the image's state
    parentPDImage.setSaveState False, pdSE_AnySave
    If undoNum = undoNumAtLastSave_PDI Then parentPDImage.setSaveState True, pdSE_SavePDI
    If undoNum = undoNumAtLastSave_Flat Then parentPDImage.setSaveState True, pdSE_SaveFlat
    
    'See if Undoing is allowed
    If undoNum > 1 Then undoState = True Else undoState = False
    
    'See if redoing is allowed
    If undoNum = undoMax Then redoState = False Else redoState = True
    
    'Select the relevant tool for this action, if one is provided
    If getUndoTool > -1 Then toolbar_Toolbox.selectNewTool getUndoTool
    
    'To inspect the current state of the Undo/Redo stack, uncomment this line.
    'debugUndo "Undo pointer moved by request to a specific location"
    
    'New to version 6.4 is our persistent creation of Undo XML header files.  These are extraordinarily helpful
    ' for both regular Undo/Redo behavior, as well as AutoSave recreation.
    writeUndoDataToXML generateUndoHeaderFilename()
    
    'Now that all layers have been updated against their most recent diff, we can finally redraw the screen
    Viewport_Engine.Stage1_InitializeBuffer pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    
    Message "Image successfully restored."
    
    'Notify the program that undo/redo data is no longer being processed, so asynchronous functions can resume their behavior.
    g_UndoRedoActive = False
    
End Sub

'Restore the last undo entry, e.g. "the user just clicked Undo".
Public Sub RestoreUndoData()
    
    'Some external actions must be deactivated while Undo/Redo processing occurs.  Such asynchronous functions can check
    ' this global variable to determine if they should fire, so we must set it before doing anything else.
    g_UndoRedoActive = True
        
    'Decrement the Undo stack pointer
    undoNum = undoNum - 1
    If undoNum < 0 Then undoNum = 0
    
    'If this is not the last Undo state, still allow the user to Undo
    If undoNum > 1 Then
        undoState = True
    
    'If we are at the bottom of the Undo stack (e.g. the image is in its original state), disable further Undo-ing
    Else
        undoState = False
    End If
    
    'If we are not at the top of the Undo stack, also enable Redo-ing
    If undoNum < undoMax Then redoState = True
        
    'If this image initially existed on the user's hard drive (e.g. it was not imported from the clipboard or
    ' some other source), note that the image has now BEEN SAVED (because it has not changed at all from its original
    ' appearance).
    If (undoNum = undoNumAtLastSave_Flat) Or (undoNum = undoNumAtLastSave_PDI) Then
    
        'In the special case of UndoNumAtLastSave = 0, check to see if this image exists on the hard drive
        If Len(parentPDImage.locationOnDisk) <> 0 Then
            
            If parentPDImage.currentFileFormat = FIF_PDI Then
                parentPDImage.setSaveState True, pdSE_SavePDI
            Else
                parentPDImage.setSaveState True, pdSE_SaveFlat
            End If
            
        Else
            parentPDImage.setSaveState False, pdSE_AnySave
        End If
        
    'If the image was imported (e.g. from a scanner or the clipboard), note that it has NOT BEEN SAVED just because
    ' the user Undo'd all the way back to its original state.
    Else
        parentPDImage.setSaveState False, pdSE_AnySave
    End If
    
    'Select the relevant tool for this action, if we were supplied one previously
    If getUndoTool > -1 Then
        toolbar_Toolbox.selectNewTool getUndoTool
    End If
    
    'With all prep work complete, we now need to figure out the most recent diff file that possesses data we need.  We handle
    ' this in two steps.
    
    'First, we ask our search routine to find the nearest relevant diff for this Undo type.  Some parameters (like the undoLayerID)
    ' will be ignored if this undo entry isn't layer-specific.
    Dim nearestDiffIndex As Long
    nearestDiffIndex = findNearestDiffIndex(undoNum, undoStack(undoNum).undoType, undoStack(undoNum).undoLayerID)
    
    'Debug.Print "Nearest diff index: " & nearestDiffIndex & "; undo type: " & undoStack(nearestDiffIndex).undoType
    
    'We now have to handle some tricky diff-related business.  If the Undo type is a full pdImage reconstruction (meaning the entire
    ' layer stack is being recreated as part of the Undo process), we have to re-apply all diffs in two steps: first, we have to
    ' reconstruct the layer stack by finding a diff that contains full layer data, e.g. a file of type UNDO_IMAGE (or in some special
    ' cases, UNDO_IMAGEHEADER).  Once the layer stack has been properly recreated, we have to update each layer and layer header
    ' against their most recent diff patch(es).  To prevent the screen from flickering like crazy while we do all this background
    ' processing, we forcibly suspend all screen redraws until our work is complete.
    '
    'Similarly, if we are undoing a single layer action (UNDO_LAYER, UNDO_LAYER_VECTORSAFE), we may need to apply an optional
    ' UNDO_LAYERHEADER diff after the fact, to make sure the layer's header is up-to-date against any header-only changes that occurred
    ' since the latest full layer DIB write-to-file.
    Dim suspendRedraws As Boolean
    If (undoStack(undoNum).undoType = UNDO_EVERYTHING) Or (undoStack(undoNum).undoType = UNDO_IMAGE) Or (undoStack(undoNum).undoType = UNDO_IMAGE_VECTORSAFE) Or (undoStack(undoNum).undoType = UNDO_IMAGEHEADER) Or (undoStack(undoNum).undoType = UNDO_LAYER) Or (undoStack(undoNum).undoType = UNDO_LAYER_VECTORSAFE) Then
        
        'If we are reverting the image to its original state, we do not need to process any extra layer data, as we already know
        ' the the source file will be accurate (because a full copy of the image is always created at Undo point 0).
        If undoNum > 0 Then
            suspendRedraws = True
        Else
            suspendRedraws = False
        End If
        
    Else
        suspendRedraws = False
    End If
    
    'With the most recent relevant diff file identified, we can now launch the undo loading routine for that diff.  It will
    ' automatically apply any changes against the relevant parts of the current pdImage object.
    LoadUndo generateUndoFilename(nearestDiffIndex), undoStack(nearestDiffIndex).undoType, undoStack(undoNum).undoType, undoStack(undoNum).undoLayerID, suspendRedraws
    
    'If redraws have been suspended, we know that we need to reconstruct one or more layers in the image against their
    ' most recent diff.
    If suspendRedraws Then
    
        Dim nearestLayerDiffDIB As Long, nearestLayerDiffHeader As Long, nearestImageDiffHeader As Long, curLayerID As Long
        Dim i As Long
        
        'The amount of diffs we need to reconstruct varies by Undo type.
        
        'If this is a full image Undo, all layers need to be checked against their most recent diffs.
        If (undoStack(undoNum).undoType = UNDO_EVERYTHING) Or (undoStack(undoNum).undoType = UNDO_IMAGE) Or (undoStack(undoNum).undoType = UNDO_IMAGE_VECTORSAFE) Or (undoStack(undoNum).undoType = UNDO_IMAGEHEADER) Then
        
            'For UNDO_EVERYTHING, UNDO_IMAGE, and UNDO_IMAGE_VECTORSAFE entries, we may need to reconstruct the layer stack
            ' before any further processing occurs.
            If (undoStack(undoNum).undoType = UNDO_EVERYTHING) Or (undoStack(undoNum).undoType = UNDO_IMAGE) Or (undoStack(undoNum).undoType = UNDO_IMAGE_VECTORSAFE) Then
            
                'Search for the nearest relevant pdImage header diff, as it may have changed independent of the rest
                ' of the pdImage stack (for example, if the user has rearranged layer order)
                nearestImageDiffHeader = findNearestDiffIndex(undoNum, UNDO_IMAGEHEADER)
                
                'If a more recent layer header was found, apply it now.
                If (nearestImageDiffHeader > nearestDiffIndex) Then
                    LoadUndo generateUndoFilename(nearestImageDiffHeader), undoStack(nearestImageDiffHeader).undoType, UNDO_IMAGEHEADER, , suspendRedraws
                End If
                
                'Also reset the nearestDiffIndex value to match the header's location; subsequent layer checks
                ' will use this as their baseline for determining whether they need to load additional data for that layer.
                nearestDiffIndex = nearestImageDiffHeader
            
            End If
                        
            'Now we know that the layer stack has the correct number of layers, in the correct order.
            ' Loop through each layer in the image, searching for more recent diffs for either the layer itself or the
            ' layer header (relevant if the user has made on-canvas changes that don't necessarily affect the layer DIB).
            For i = 0 To pdImages(g_CurrentImage).getNumOfLayers - 1
            
                'Grab the current layer ID (all undo functions operate on layer ID, not index, as index may change unexpectedly)
                curLayerID = pdImages(g_CurrentImage).getLayerByIndex(i).getLayerID
            
                'Start by reconstructing a layer DIB using the default UNDO_LAYER type
                nearestLayerDiffDIB = findNearestDiffIndex(undoNum, UNDO_LAYER, curLayerID)
                
                'Vector layers can also be modified by the UNDO_LAYER_VECTORSAFE type, which simply signifies layer changes
                ' that don't require rasterization.  If a newer version of the undo for this layer exists under this type,
                ' take it instead.
                'Dim vectorCheck As Long
                'vectorCheck = findNearestDiffIndex(undoNum, UNDO_LAYER_VECTORSAFE, curLayerID)
                
                'If vectorCheck > nearestLayerDiffDIB Then nearestLayerDiffDIB = vectorCheck
                
                'If the nearest diff for this layer is *not* the same as the nearest diff for the full image, apply this diff
                ' to the layer in question.
                If (nearestLayerDiffDIB > nearestDiffIndex) Then
                
                    'If vectorCheck = nearestLayerDiffDIB Then
                    '    LoadUndo generateUndoFilename(nearestLayerDiffDIB), undoStack(nearestLayerDiffDIB).undoType, UNDO_LAYER_VECTORSAFE, curLayerID, suspendRedraws
                    'Else
                        LoadUndo generateUndoFilename(nearestLayerDiffDIB), undoStack(nearestLayerDiffDIB).undoType, UNDO_LAYER, curLayerID, suspendRedraws
                    'End If
                    
                End If
                
                'Next, we must reconstruct the layer header, as it too may have changed (for example, if the user moved the layer
                ' since the last action)
                nearestLayerDiffHeader = findNearestDiffIndex(undoNum, UNDO_LAYERHEADER, curLayerID)
                
                'If a more recent layer header was found, apply it now.
                If (nearestLayerDiffHeader > nearestDiffIndex) And (nearestLayerDiffHeader > nearestLayerDiffDIB) Then
                    LoadUndo generateUndoFilename(nearestLayerDiffHeader), undoStack(nearestLayerDiffHeader).undoType, UNDO_LAYERHEADER, curLayerID, suspendRedraws
                End If
                
            Next i
        
        'If this is a just a layer Undo, we simply need to verify the layer header against any diffs more recent than the last
        ' full DIB reconstruction for this layer.
        ElseIf (undoStack(undoNum).undoType = UNDO_LAYER) Or (undoStack(undoNum).undoType = UNDO_LAYER_VECTORSAFE) Then
        
            'Grab the current layer ID (all undo functions operate on layer ID, not index, as index may change unexpectedly)
            curLayerID = undoStack(undoNum).undoLayerID
            
            'Next, look for the most recent layer header diff, as it may have changed independent of the layer DIB
            ' (for example, if the user moves a layer after painting on it)
            nearestLayerDiffHeader = findNearestDiffIndex(undoNum, UNDO_LAYERHEADER, curLayerID)
                
            'If a more recent layer header was found, apply it now.
            If (nearestLayerDiffHeader > nearestDiffIndex) Then
                LoadUndo generateUndoFilename(nearestLayerDiffHeader), undoStack(nearestLayerDiffHeader).undoType, UNDO_LAYERHEADER, curLayerID, suspendRedraws
            End If
        
        End If
        
        'Now that all layers have been updated against their most recent diff, we can finally redraw the screen
        Viewport_Engine.Stage1_InitializeBuffer pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    
    End If
    
    'To inspect the current state of the Undo/Redo stack, uncomment this line.
    'debugUndo "user clicked Undo"
    
    'New to version 6.4 is our persistent creation of Undo XML header files.  These are extraordinarily helpful
    ' for both regular Undo/Redo behavior, as well as AutoSave recreation.
    writeUndoDataToXML generateUndoHeaderFilename()
    
    'Notify the program that undo/redo data is no longer being processed, so asynchronous functions can resume their behavior.
    g_UndoRedoActive = False
    
End Sub

'Given an undo Index, find the nearest Diff file that contains the information we need to reconstruct the image's state
' at that point in time.
Private Function findNearestDiffIndex(ByVal undoIndex As Long, ByVal undoTypeCheck As PD_UNDO_TYPE, Optional ByVal relevantLayerID As Long = -1) As Long

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
        Select Case undoStack(i).undoType
        
            'UNDO_EVERYTHING works for all possible Undo types, as it has a copy of everything we would ever need!
            Case UNDO_EVERYTHING
                findNearestDiffIndex = i
                Exit For
            
            'UNDO_IMAGE contains relevant diff information for everything but selections
            Case UNDO_IMAGE
                If (undoTypeCheck <> UNDO_SELECTION) Then
                    findNearestDiffIndex = i
                    Exit For
                End If
                
            'UNDO_IMAGE_VECTORSAFE is identical to UNDO_IMAGE
            Case UNDO_IMAGE_VECTORSAFE
                If (undoTypeCheck <> UNDO_SELECTION) Then
                    findNearestDiffIndex = i
                    Exit For
                End If
            
            'UNDO_IMAGEHEADER contains relevant diff information for itself, and UNDO_LAYERHEADER only.
            Case UNDO_IMAGEHEADER
                If (undoTypeCheck = UNDO_IMAGEHEADER) Or (undoTypeCheck = UNDO_LAYERHEADER) Then
                    findNearestDiffIndex = i
                    Exit For
                End If
            
            'UNDO_LAYER contains relevant diff information for matching UNDO_LAYER and UNDO_LAYER_HEADER actions,
            ' but *only if the layer ID matches the requested one!*
            Case UNDO_LAYER
                If ((undoTypeCheck = UNDO_LAYER) Or (undoTypeCheck = UNDO_LAYER_VECTORSAFE) Or (undoTypeCheck = UNDO_LAYERHEADER)) Then
                    
                    If (relevantLayerID = undoStack(i).undoLayerID) Then
                        findNearestDiffIndex = i
                        Exit For
                    End If
                    
                End If
                
            'UNDO_LAYER_VECTORSAFE is identical to UNDO_LAYER
            Case UNDO_LAYER_VECTORSAFE
                If ((undoTypeCheck = UNDO_LAYER) Or (undoTypeCheck = UNDO_LAYER_VECTORSAFE) Or (undoTypeCheck = UNDO_LAYERHEADER)) Then
                    
                    If (relevantLayerID = undoStack(i).undoLayerID) Then
                        findNearestDiffIndex = i
                        Exit For
                    End If
                    
                End If
                
            'UNDO_LAYERHEADER contains relevant diff information for matching UNDO_LAYER_HEADER actions,
            ' but *only if the layer ID matches the requested one!*
            Case UNDO_LAYERHEADER
                If (undoTypeCheck = UNDO_LAYERHEADER) And (relevantLayerID = undoStack(i).undoLayerID) Then
                    findNearestDiffIndex = i
                    Exit For
                End If
            
            'UNDO_SELECTION contains relevant diff information for selections only
            Case UNDO_SELECTION
                If (undoTypeCheck = UNDO_SELECTION) Then
                    findNearestDiffIndex = i
                    Exit For
                End If
                
            'This Case Else statement will not be necessary when we have implemented all Diff types.  Until that happens,
            ' however, we can treat it the same as UNDO_IMAGE.
            Case Else
                If (undoTypeCheck <> UNDO_SELECTION) Then
                    findNearestDiffIndex = i
                    Exit For
                End If
        
        End Select
    
    Next i
    
    'The function value will have been correctly set within the For loop, so we can now exit immediately.

End Function

'Restore a redo entry, e.g. "the user just clicked Redo"
Public Sub RestoreRedoData()
    
    'Some external actions must be deactivated while Undo/Redo processing occurs.  Such asynchronous functions can check
    ' this global variable to determine if they should fire, so we must set it before doing anything else.
    g_UndoRedoActive = True
    
    'If we're not at the top of the Undo stack, advance the stack index by one
    If undoNum < undoMax Then
        
        undoNum = undoNum + 1
        
        'Allow saving contingent on the image's state
        parentPDImage.setSaveState False, pdSE_AnySave
        If undoNum = undoNumAtLastSave_PDI Then parentPDImage.setSaveState True, pdSE_SavePDI
        If undoNum = undoNumAtLastSave_Flat Then parentPDImage.setSaveState True, pdSE_SaveFlat
    
    End If
    
    'Undoing is always allowed after the user has Redone something
    undoState = True
    
    'If we're at the end of our action chain, disable any further Redo-ing
    If undoNum = undoMax Then redoState = False Else redoState = True
    
    'Select the relevant tool for this action, if one is provided
    If getUndoTool > -1 Then toolbar_Toolbox.selectNewTool getUndoTool
    
    'The great thing about Redo is that it chronogically moves in the same direction as PD's Undo diff files.  So unlike
    ' Undo - which requires us to traverse the undo timeline in reverse, looking for a relevant diff file - we can simply
    ' apply the next valid diff in line, no extra work required!
    LoadUndo generateUndoFilename(undoNum - 1), getUndoProcessType, getUndoProcessType, undoStack(undoNum - 1).undoLayerID
    
    'To inspect the current state of the Undo/Redo stack, uncomment this line.
    'debugUndo "user clicked Redo"
    
    'New to version 6.4 is our persistent creation of Undo XML header files.  These are extraordinarily helpful
    ' for both regular Undo/Redo behavior, as well as AutoSave recreation.
    writeUndoDataToXML generateUndoHeaderFilename()
    
    'Notify the program that undo/redo data is no longer being processed, so asynchronous functions can resume their behavior.
    g_UndoRedoActive = False
    
End Sub

'Find the last modified layer in the image, and fill a supplied DIB with the contents of that layer *prior to its current state*,
' (obviously, as the caller has access to the current state).  This is currently used by the "Edit > Fade..." feature.
'
'The relevant layer index is returned by this function because the user may switch to a new layer, but we still want Fade to be
' available.  (The Undo/Redo engine stores this data anyway, so it's not problematic to receive it.)  To help with menu captioning,
' this function will also return the name of the action responsible for the change.
'
'This function will return FALSE if no layer-specific changes have been made, and thus no Fade action would be relevant.
Public Function fillDIBWithLastUndoCopy(ByRef dstDIB As pdDIB, ByRef affectedLayerID As Long, Optional ByRef nameOfProcessorAction As String, Optional ByVal checkStateOnly As Boolean = True) As Boolean
    
    If undoNum = 0 Then
        fillDIBWithLastUndoCopy = 0
        Exit Function
    End If
    
    'First things first: start searching the Undo/Redo stack for a layer-specific edit.  We don't care what edit, just that it's
    ' layer-specific.
    Dim actionFound As Boolean
    actionFound = False
    
    Dim layerChangeIndex As Long
    layerChangeIndex = -1
    
    Dim i As Long
    For i = undoNum - 1 To 0 Step -1
    
        'A layer-specific edit has been found.  This is the last action the user committed that modified a layer, so it will contain
        ' a copy of the image as it appeared *after the change was made*.  What this function needs to do is find the state of this
        ' layer *prior to the change* - and now that we know which layer is affected, we can do that!
        If (undoStack(i).undoType = UNDO_LAYER) Or (undoStack(i).undoType = UNDO_LAYER_VECTORSAFE) Then
            
            'Certain actions are irrelevant to this function; in particular, "Fade" and "Repeat" cannot be "faded" or "repeated".
            ' In these instances, we want to continue searching backward, to find the original function that was faded or repeated.
            If (StrComp(UCase$(undoStack(i).processID), "FADE", vbBinaryCompare) <> 0) And (StrComp(UCase$(undoStack(i).processID), "REPEAT", vbBinaryCompare) <> 0) Then
                actionFound = True
                affectedLayerID = undoStack(i).undoLayerID
                nameOfProcessorAction = undoStack(i).processID
                layerChangeIndex = i
                Exit For
            End If
            
        End If
    
    Next i
    
    'If a relevant layer was found, actionFound will be set to TRUE.  If a relevant layer was not found, return FALSE and exit.
    If actionFound Then
        
        'Some functions, like SyncInterfaceToCurrentImage, only care if a recent Layer modification exists.  They don't actually need
        ' us to return the relevant DIB.  If checkStateOnly = TRUE, exit now.
        If checkStateOnly Then
            fillDIBWithLastUndoCopy = True
            Exit Function
        End If
        
        'Armed with the knowledge of which layer is being affected, we now need to find the last relevant DIFF index for that layer.
        ' Use the built-in function to do that.
        Dim relevantDiffIndex As Long
        relevantDiffIndex = findNearestDiffIndex(layerChangeIndex, UNDO_LAYER, affectedLayerID)
        
        'Check also for vector-safe actions, which are identical to UNDO_LAYER but for operations that support vector-safe modifications.
        Dim relevantVectorIndex As Long
        relevantVectorIndex = findNearestDiffIndex(layerChangeIndex, UNDO_LAYER_VECTORSAFE, affectedLayerID)
        
        'relevantDiffIndex will now contain the index in the UndoStack array corresponding to the previous state of the affected layer.
        ' Load the relevant portion of that Undo file into the target DIB.
        Dim tmpLayer As pdLayer
        Set tmpLayer = New pdLayer
        
        If relevantVectorIndex > relevantDiffIndex Then
            LoadUndo generateUndoFilename(relevantVectorIndex), undoStack(relevantVectorIndex).undoType, UNDO_LAYER_VECTORSAFE, affectedLayerID, True, tmpLayer
        Else
            LoadUndo generateUndoFilename(relevantDiffIndex), undoStack(relevantDiffIndex).undoType, UNDO_LAYER, affectedLayerID, True, tmpLayer
        End If
        
        'tmpLayer will now contain a copy of the relevant DIB (inside .layerDIB).  Copy it into the destination DIB, then exit.
        Set dstDIB = New pdDIB
        dstDIB.createFromExistingDIB tmpLayer.layerDIB
        
        Set tmpLayer = Nothing
        fillDIBWithLastUndoCopy = True
        
    Else
        fillDIBWithLastUndoCopy = False
    End If
    
End Function

'Get the filename of the current Undo file, using the same formula as generateUndoFilename below
Public Function getCurrentUndoFilename()
    getCurrentUndoFilename = generateUndoFilename(undoNum)
End Function

'Generate a unique Undo/Redo file name, including full temp path.
Private Function generateUndoFilename(ByVal uIndex As Long) As String
    
    'As a failsafe, all Undo/Redo files are erased when this class is unloaded.  The clearUndos function requires this
    ' function to generate Undo filnames.  If our parent class was terminated properly, it will have already erased
    ' our reference to it, so the parentPDImage.imageID reference below will fail - to prevent that, we must check the
    ' parent reference in advance.
    If Not (parentPDImage Is Nothing) Then
        generateUndoFilename = g_UserPreferences.GetTempPath & "~cPDU_(" & g_SessionID & ")_" & parentPDImage.imageID & "_" & uIndex & "_.pdtmp"
    End If
    
End Function

'Publicly accessible mechanism for generating a unique Undo/Redo file name.  Any changes to the private function, above,
' should be mirrored here.
Public Function generateUndoFilenameExternal(ByVal imageID As Long, ByVal uIndex As Long, ByVal originalSessionID As String) As String
    generateUndoFilenameExternal = g_UserPreferences.GetTempPath & "~cPDU_(" & originalSessionID & ")_" & imageID & "_" & uIndex & "_.pdtmp"
End Function

'Generate a unique Undo/Redo header filename, including full temp path.  Note that this function does *NOT* take an
' Undo number index.  The entire stack is written out to the header file, every time.
Private Function generateUndoHeaderFilename() As String
    
    'As a failsafe, all Undo/Redo files are erased when this class is unloaded.  The clearUndos function requires this
    ' function to generate Undo filnames.  If our parent class was terminated properly, it will have already erased
    ' our reference to it, so the parentPDImage.imageID reference below will fail - to prevent that, we must check the
    ' parent reference in advance.
    If Not (parentPDImage Is Nothing) Then
        generateUndoHeaderFilename = g_UserPreferences.GetTempPath & "~PDU_StackSummary_(" & g_SessionID & ")_" & parentPDImage.imageID & "_.pdtmp"
    End If
    
End Function

'Erase all undo files that this class has created
Public Sub clearUndos()
    
    On Error Resume Next
    
    'Remember that we have no Undo files for this image
    undoState = False
    
    'If Undo data has already been cleared, exit
    If undoTotal = 0 Then Exit Sub
    
    Dim cFile As pdFSO
    Set cFile = New pdFSO
    
    'Destroy every potential Undo file
    Dim i As Long
    For i = 0 To undoTotal
        
        'Check image data first...
        If cFile.FileExist(generateUndoFilename(i)) Then cFile.KillFile generateUndoFilename(i)
        
        '...followed by layer data
        If cFile.FileExist(generateUndoFilename(i) & ".layer") Then cFile.KillFile generateUndoFilename(i) & ".layer"
        
        '...followed by selection data
        If cFile.FileExist(generateUndoFilename(i) & ".selection") Then cFile.KillFile generateUndoFilename(i) & ".selection"
        
    Next i
    
    'Destroy the header file and preview image, if any
    If cFile.FileExist(generateUndoHeaderFilename()) Then cFile.KillFile generateUndoHeaderFilename()
    If cFile.FileExist(generateUndoHeaderFilename() & ".asp") Then cFile.KillFile generateUndoHeaderFilename() & ".asp"
    
    'Set the current and maximum Undo number to 0
    undoNum = 0
    undoMax = 0
    undoTotal = 0
    
End Sub

'Return the tool associated with the current Undo state
Public Function getUndoTool() As Long
    If undoNum > 0 Then getUndoTool = undoStack(undoNum - 1).relevantTool
End Function

'Return the process type associated with the current Undo state
Public Function getUndoProcessType() As PD_UNDO_TYPE
    If undoState Then
        getUndoProcessType = undoStack(undoNum - 1).undoType
    Else
        getUndoProcessType = UNDO_NOTHING
    End If
End Function

'Return the process type associated with the current Redo state
Public Function getRedoProcessType() As PD_UNDO_TYPE
    If redoState Then
        getRedoProcessType = undoStack(undoNum).undoType
    Else
        getRedoProcessType = UNDO_NOTHING
    End If
End Function

'Return the processID associated with the current Undo state
Public Function getUndoProcessID() As String
    If undoState Then
        getUndoProcessID = g_Language.TranslateMessage(undoStack(undoNum - 1).processID)
    Else
        getUndoProcessID = ""
    End If
End Function

'Return the processID associated with the current Redo state
Public Function getRedoProcessID() As String
    If redoState Then
        getRedoProcessID = g_Language.TranslateMessage(undoStack(undoNum).processID)
    Else
        getRedoProcessID = ""
    End If
End Function

'Given a type of Undo action, find the last relevant parameter string.  This is used by PD's central processor to detect on-canvas
' changes to the image since the last Undo was created.
Public Function getLastParamString(ByVal requestedUndoType As PD_UNDO_TYPE) As String

    'If no Undos have been created, exit immediately
    If undoNum = 0 Then
        getLastParamString = ""
        Exit Function
    End If

    'Loop backward through the Undo stack, looking for a matching Undo type.  If found, return the associated
    ' param string and exit immediately.
    Dim i As Long
    
    For i = undoNum - 1 To 0 Step -1
        If undoStack(i).undoType = requestedUndoType Then
            getLastParamString = undoStack(i).processParamString
            Exit Function
        End If
    Next i
    
    'If we made it all the way here, no matching function was found.  Return nothing.
    getLastParamString = ""

End Function

'Undo/redo data handling is complex.  Sometimes it is helpful to see how our Undo/Redo stack looks - use this sub
' to print out a full list to the debug window.
Public Sub debugUndo(ByVal sourceOfDebugRequest As String)

    Dim i As Long
    Debug.Print "---"
    Debug.Print " Current Undo stack:"
    Debug.Print "  (Displaying because " & sourceOfDebugRequest & ")"
    For i = 0 To UBound(undoStack)
        If i = undoNum - 1 Then
            Debug.Print "  * " & i & ": " & undoStack(i).processID & ", T:" & undoStack(i).undoType & ", L:" & undoStack(i).undoLayerID
        Else
            Debug.Print "  " & i & ": " & undoStack(i).processID & ", T:" & undoStack(i).undoType & ", L:" & undoStack(i).undoLayerID
        End If
    Next i
    Debug.Print "---"

End Sub

Private Sub Class_Initialize()
    
    'Initialize all relevant undo/redo tracking variables
    ReDim undoStack(0) As undoEntry
    
    undoNum = 0
    undoMax = 0
    undoNumAtLastSave_Flat = 0
    undoNumAtLastSave_PDI = 0
    undoState = False
    redoState = False
    
End Sub

'When this class is unloaded, we must make sure to delete any Undo files we created
Private Sub Class_Terminate()
    clearUndos
End Sub

'Write the Undo stack out to file, in XML format
Public Sub writeUndoDataToXML(ByVal dstFilename As String)

    'Create a pdXML class, which will help us assemble the XML file
    Dim xmlEngine As pdXML
    Set xmlEngine = New pdXML
    xmlEngine.prepareNewXML "Undo stack"
    
    'Write out the XML version we're using for this macro
    xmlEngine.writeTag "pdUndoVersion", pdUndoVersionID
    xmlEngine.writeBlankLine
    
    'Write out some basic information about our parent pdImage object.  In the event of a critical program crash,
    ' the AutoSave engine can use this information to reconstruct any missing images.
    xmlEngine.writeTag "imageID", parentPDImage.imageID
    xmlEngine.writeTag "friendlyName", parentPDImage.originalFileNameAndExtension
    xmlEngine.writeTag "originalPath", parentPDImage.locationOnDisk
    xmlEngine.writeBlankLine
    
    'Write out the size of the current Undo stack, including the Max and Total values
    xmlEngine.writeTag "StackHeight", undoMax
    xmlEngine.writeTag "StackAbsoluteMaximum", undoTotal
    xmlEngine.writeTag "CurrentStackPointer", undoNum
    
    'We probably don't need the undoNumAtLastSave values, but it doesn't hurt to embed them, "just in case"
    xmlEngine.writeTag "UndoNumAtLastSave_Flat", undoNumAtLastSave_Flat
    xmlEngine.writeTag "UndoNumAtLastSave_PDI", undoNumAtLastSave_PDI
    
    'Finally, add the current session ID, a 16-character string unique to each running PD instance.  This ID is appended to all
    ' Undo files, which allows multiple PD sessions to run concurrently without trouble.
    xmlEngine.writeTag "OriginalSessionID", g_SessionID
    
    xmlEngine.writeBlankLine
    
    'Next, we're going to write out all Undo stack entries. If PD's "UndoEntry" type is ever changed, note
    ' that this function will need to be updated to match!
    Dim i As Long
    For i = 0 To UBound(undoStack)
        
        'Start each undo entry with a unique numeric identifier (that corresponds to its location in the stack)
        xmlEngine.writeTagWithAttribute "UndoEntry", "index", i, "", True
        
        'Write out all the properties of this entry
        xmlEngine.writeTag "ProcessID", undoStack(i).processID
        xmlEngine.writeTag "ProcessParameters", undoStack(i).processParamString
        xmlEngine.writeTag "UndoType", undoStack(i).undoType
        xmlEngine.writeTag "RelevantLayerID", undoStack(i).undoLayerID
        xmlEngine.writeTag "RelevantTool", undoStack(i).relevantTool
        
        'Close this undo entry
        xmlEngine.closeTag "UndoEntry"
        xmlEngine.writeBlankLine
    
    Next i
    
    'With all Undo data successfully written, close the XML data and write it out to file.
    xmlEngine.writeXMLToFile dstFilename, False
    
    'As an added bonus (new to PD 6.4), also write out a thumbnail of the image *as it appears right now*.  The Autosave engine
    ' will use this as an image preview if something goes horribly wrong.
    If (undoNum > 0) Then
        If (Not (undoStack(undoNum - 1).thumbnailLarge Is Nothing)) Then Saving.QuickSaveDIBAsPNG generateUndoHeaderFilename & ".asp", undoStack(undoNum - 1).thumbnailLarge
    End If
    
End Sub

'Given an XML string generated by another pdUndo instance, reconstruct this stack to match.  This is used by the Autosave engine
' to restore a full previous session.
Public Function reconstructStackFromExternalSource(ByRef srcXMLString As String) As Boolean

    'Initiate an XML parser
    Dim xmlEngine As pdXML
    Set xmlEngine = New pdXML
    
    'Validate the XML string
    If xmlEngine.loadXMLFromString(srcXMLString) Then
    
        'Make sure the XML string actually contains Undo/Redo data
        If xmlEngine.isPDDataType("Undo stack") And xmlEngine.validateLoadedXMLData("pdUndoVersion") Then
            
            'The file checks out!  Start by initiating a few module-level values to match the data in the file.
            undoMax = xmlEngine.getUniqueTag_Long("StackHeight")
            undoTotal = xmlEngine.getUniqueTag_Long("StackAbsoluteMaximum")
            undoNum = xmlEngine.getUniqueTag_Long("CurrentStackPointer")
            undoNumAtLastSave_Flat = xmlEngine.getUniqueTag_Long("UndoNumAtLastSave_Flat")
            undoNumAtLastSave_PDI = xmlEngine.getUniqueTag_Long("UndoNumAtLastSave_PDI")
            
            'Next, we must reconstruct the Undo stack itself.  Start by preparing the stack array.
            ReDim undoStack(0 To undoTotal) As undoEntry
            
            'Start retrieving individual stack entries from the file
            Dim i As Long
            For i = 0 To UBound(undoStack)
                
                'Start by finding the location of the tag we want
                Dim tagPosition As Long
                tagPosition = xmlEngine.getLocationOfTagPlusAttribute("UndoEntry", "index", i)
                
                'If a matching tag was found, retrieve its data and store it in the stack
                If tagPosition > 0 Then
                
                    With undoStack(i)
                    
                        .processID = xmlEngine.getUniqueTag_String("ProcessID", , tagPosition)
                        .processParamString = xmlEngine.getUniqueTag_String("ProcessParameters", , tagPosition)
                        .undoType = xmlEngine.getUniqueTag_Long("UndoType", , tagPosition)
                        .undoLayerID = xmlEngine.getUniqueTag_Long("RelevantLayerID", , tagPosition)
                        .relevantTool = xmlEngine.getUniqueTag_Long("RelevantTool", , tagPosition)
                        
                    End With
                    
                    'Initialize the two thumbnail DIBs, even though they won't contain anything useful.
                    Set undoStack(i).thumbnailSmall = New pdDIB
                    Set undoStack(i).thumbnailLarge = New pdDIB
                
                End If
                
            Next i
            
            'The stack was reconstructed successfully.
            reconstructStackFromExternalSource = True
            
        Else
            reconstructStackFromExternalSource = False
        End If
    
    Else
        reconstructStackFromExternalSource = False
    End If

End Function

'External functions can request a copy of the Undo stack by calling this function and supply the three required parameters
' (all passed ByRef, obviously):
' 1) Number of Undo entries
' 2) Current Undo entry
' 3) Array of type undoEntry, to receive the full contents of the stack
Friend Sub copyUndoStack(ByRef numOfUndos As Long, ByRef curUndoNumber As Long, ByRef stackList() As undoEntry)

    numOfUndos = undoMax
    curUndoNumber = undoNum
    
    ReDim stackList(0 To undoMax) As undoEntry
    
    Dim i As Long
    For i = 0 To undoMax - 1
        
        With stackList(i)
            
            .processID = undoStack(i).processID
            .processParamString = undoStack(i).processParamString
            .relevantTool = undoStack(i).relevantTool
            .undoLayerID = undoStack(i).undoLayerID
            .undoType = undoStack(i).undoType
            
            Set .thumbnailLarge = New pdDIB
            .thumbnailLarge.createFromExistingDIB undoStack(i).thumbnailLarge
            
            Set .thumbnailSmall = New pdDIB
            .thumbnailSmall.createFromExistingDIB undoStack(i).thumbnailSmall
            
        End With
        
    Next i

End Sub
