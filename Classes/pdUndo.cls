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
'Copyright ©2001-2014 by Tanner Helland
'Created: 2/April/01
'Last updated: 09/April/14
'Last update: for now, Undo / Redo is always going to make a full copy of the entire pdImage structure.  I know, I know,
'             this is extrordinarily wasteful - but cracking a lighter version of Undo/Redo is *tough*.  Sometimes a
'             full copy of the pdImage is necessary (e.g. Flatten Image), while at other times, something much lighter
'             would suffice.  That's another project for another day, however.
'
'PhotoDemon has had Undo/Redo capabilities for a long time.  The first incarnation (way back in 2001) stored a single
' byte-array copy of the image in memory, allowing for a lone Undo step if necessary.  Eventually this was converted
' to a more robust HDD-based solution, and further updates have tacked on features to (the very poor) original
' Undo/Redo framework.
'
'In 2013, I finally had enough of the program's shoddy Undo/Redo code, so I merged all the various Undo/Redo subs
' into a single class, and gave each pdImage a copy of said class.  This achieves a long-desired goal of making
' the Undo/Redo code completely self-contained, and should make future improvements much, much easier.  As part
' of this change, the Undo/Redo engine was largely rewritten, and all interface-related code was moved outside the
' class.  (Specifically, it is now handled by the separate syncInterfaceToCurrentImage function.)
'
'(NOTE: the following paragraph is currently irrelevant; I may revisit an active engine in the future, but it has
' been placed on hold due to some nasty interactions elsewhere in the project.)
'In 2014, as part of adding Layers support, the Undo/Redo engine was changed from a lazy implementation to an active
' one.  Basically, the previous "lazy" implementation did not create Undo/Redo data until absolutely necessary.  This
' had some benefits; for example, Undo data was not created until an image was actually edited, improving image load
' time, and Redo data was never created until Undo was pressed.  While good for conserving memory, this system
' ultimately proved untenable, as it required the storage of a lot of "what happened previously" data in order to
' handle Undo creation correctly for odd cases (like when an image edit is applied after a selection modification).
' To work around this, Undo/Redo data is now created immediately for the active function.
'
'As an added benefit, the Active system also works much better with PD's Autosave engine.
'
'Undo/Redo data is still saved to the HDD, and will remain that way for the forseeable future.  Unlimited operations
' are allowed.  This class does not check HDD space availability, so errors may be thrown if you run out of free
' space.  Similarly, clearing your TEMP folder in the middle of program operation could cause a lot of bad things,
' including hard crashes.  I should probably fix this at some point in the future, but as no one has ever complained,
' it remains a low priority for me.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Undo/Redo currently works as a stack, and entries can only be traversed sequentially.  In the future, we will be
' reworking this to allow the program to "jump" to any Undo entry arbitrarily.  To that end, I am implementing
' Undo/Redo tracking in a more comprehensive way than may seem necessary for the current implementation.

'In the Undo/Redo stack, which entry is this image currently on?
Private undoNum As Long

'How many Undo/Redo entries exist in our stack?
Private undoMax As Long
Private undoTotal As Long

'Are Undo/Redo operations allowed for this image?  (This is used to determine the state of undo/redo buttons and
' menu items.)
Private undoState As Boolean
Private redoState As Boolean

'Information about each Undo entry is stored in an array; the array is dynamically resized as necessary when new
' Undos are created.  We track the ID of each action in preparation for a future History browser that allows the
' user to jump to any arbitrary Undo/Redo state.  (Also, to properly update the text of the Undo/Redo menu and
' buttons so the user knows which action they are undo/redoing.)
Private Type undoEntry
    processID As String     'Name of the associated action (e.g. "Gaussian blur")
    undoType As Long        'What type of Undo/Redo data was stored for this action (e.g. Image or Selection data)
    relevantTool As Long    'If a tool was associated with this action, it can be set here.  This value is not currently used.
End Type

Private undoStack() As undoEntry

'If the user performs 10 actions, then presses Undo 5 times and saves the image, we need to disable the Save button
' at that Undo point, because it was the last point at which the image was saved.  Thus we must track the last point
' in the Undo/Redo stack where the image was saved.
Private undoNumAtLastSave As Long

'We need access to the parent class so we can request things like "write your current data out to file".
Public parentPDImage As pdImage

'When an image is saved, this class needs to be notified so that it can make a note of the current undo index
Public Sub notifyImageSaved()
    undoNumAtLastSave = undoNum
End Sub

Public Function getUndoState() As Boolean
    getUndoState = undoState
End Function

Public Function getRedoState() As Boolean
    getRedoState = redoState
End Function

'With the creation of PD's Autosave engine, it became helpful to write some Undo immediately, outside of the
' traditional Undo/Redo stack.  This function can be used to do so.
Public Sub writeOneOffUndoEntry()

    'Ask the parent PD image to write itself out to file
    Saving.SavePhotoDemonImage parentPDImage, getCurrentUndoFilename(), True, False, False, False
    parentPDImage.mainSelection.writeSelectionToFile getCurrentUndoFilename() & ".selection"
    
    Debug.Print "Saved " & getCurrentUndoFilename() & " at the request of PD's Autosave engine."

End Sub

'Create an Undo entry (e.g. save a copy of the present image or selection to the temp directory)
' Inputs:
'  1) the ID string of the process that called this action (e.g. "Gaussian blur")
'  2) optionally, the type of Undo that needs to be created.  By default, type 1 (image pixel undo) is assumed.
'  3) optionally, the tool associated with this action.  This value is not currently used, but in the future it could
'     be helpful for automatically selecting the tool associated with a given Undo/Redo point.
Public Sub createUndoData(ByVal processID As String, Optional ByVal processType As Long = 1, Optional ByVal curTool As Long = -1)
    
    'Depending on the type of process being applied, save out image or selection data to file.
    ' NOTE: this is currently disabled because I can't make it work :/.  For now, always attempt to create both
    '       selection and image undo data.  If a selection is not available, that function will simply exit.
    'Select Case processType
    
        'Pixel data
        'Case 1
            Saving.SavePhotoDemonImage parentPDImage, getCurrentUndoFilename(), True, False, False, False
            
        'Selection data
        'Case 2
            parentPDImage.mainSelection.writeSelectionToFile getCurrentUndoFilename() & ".selection"
        
    'End Select
    
    'Previously, we notified the user when undo/redo data was saved.  This was stupid, but it is helpful to
    ' display the info in the debug window.
    Debug.Print "Saved " & getCurrentUndoFilename() & " as Undo file #" & undoNum & "."
    
    'Store all relevant information about this Undo action, so we can access it in the future if necessary
    ReDim Preserve undoStack(0 To undoNum) As undoEntry
    With undoStack(undoNum)
        .processID = processID
        .undoType = processType
        .relevantTool = curTool
    End With
    
    'To inspect the current state of the Undo/Redo stack, uncomment this line.
    'debugUndo
    
    'If the present image is NOT the last image in the undo train (i.e., the user pressed Undo before running
    ' the current operation), treat this action as the new "last" Undo entry in the stack.  However, we still need
    ' to remember that Undo files exist past this point, so we can clean them up when we exit - hence the need for
    ' a separate undoTotal variable.
    undoMax = undoNum
    
    'Increase our Undo count
    undoNum = undoNum + 1
    undoMax = undoMax + 1
    
    'If the present image is *truly* the last one in the Undo train, note that we have more Undo files to delete
    ' when our parent is unloaded.
    If undoMax > undoTotal Then undoTotal = undoMax
    
    'Make a note that at this point, the user should be allowed to Undo, but not Redo
    undoState = True
    redoState = False
    
    'If this action changed pixel data, mark the image as NOT saved (because if we're creating an Undo file, the
    ' image has changed from its saved version).
    If processType = 1 Then parentPDImage.setSaveState False
    
    Debug.Print "Undo created successfully."
    
End Sub

'This special command can be used to revert the image to its last saved state.  For images that have never been
' saved (e.g. those pasted from the clipboard, etc), it will undo to the image's original state.
Public Sub revertToLastSavedState()
    If undoNumAtLastSave > 0 Then
        restoreUndoData undoNumAtLastSave
    Else
        restoreUndoData 0
    End If
End Sub

'Restore an undo entry, e.g. "the user just clicked Undo".
' Optionally, a specific position in the Undo/Redo chain can be specified if the user wants to revert to
' a specific state in an image's history.  (For example, the File -> Revert command calls this function
' with an undoPosition of 0.)
Public Sub restoreUndoData(Optional ByVal undoPosition As Long = -1)
    
    'Some external actions are deactivated while Undo/Redo processing occurs.  Such functions can check this global
    ' variable to determine if they should fire - so set it before doing anything else.
    g_UndoRedoActive = True
    
    'If we're at the top of the Undo stack, save the image's current state - this allows the user to Redo whichever
    ' action they're about to Undo.
    If undoNum = undoMax Then
        
        'Depending on the last action, write out a Redo file for that action
        'Select Case UndoTypes(UndoNum - 1)
        
            'Pixel undo data
            'Case 1
                Saving.SavePhotoDemonImage parentPDImage, getCurrentUndoFilename(), True, False, False, False
            
            'Selection undo data
            'Case 2
                parentPDImage.mainSelection.writeSelectionToFile getCurrentUndoFilename() & ".selection"
            
        'End Select
        
        'Previously, we notified the user when undo/redo data was saved.  This was stupid, but it is helpful to
        ' display the info in the debug window.
        Debug.Print "Undo initiated by user. Saved " & getCurrentUndoFilename() & " for Redo purposes."
        
        'Move neighboring values into the current Undo location (this makes Redo work properly)
        ReDim Preserve undoStack(0 To undoNum) As undoEntry
        With undoStack(undoNum)
            .processID = undoStack(undoNum - 1).processID
            .undoType = undoStack(undoNum - 1).undoType
            .relevantTool = undoStack(undoNum - 1).relevantTool
        End With
        
    End If
    
    'If the user supplied an undoPosition, apply it now.  Otherwise, decrement it from its current position.
    If undoPosition <> -1 Then
        undoNum = undoPosition
    Else
        undoNum = undoNum - 1
        If undoNum < 0 Then undoNum = 0
    End If
    
    'If this is not the last Undo state, still allow the user to Undo.
    If undoNum > 0 Then undoState = True
    
    'If we are not at the top of the Undo stack, also enable Redo-ing
    If undoNum < undoMax Then redoState = True
    
    'If we are at the bottom of the Undo stack (e.g. the image is in its original state), disable further Undo-ing
    If undoNum = 0 Then undoState = False
    
    'If this image initially existed on the user's hard drive (e.g. it was not imported from the clipboard or
    ' some other source), note that the image has now BEEN SAVED (because it has not changed at all from its original
    ' appearance).
    If undoNum = undoNumAtLastSave Then
    
        'In the special case of UndoNumAtLastSave = 0, check to see if this image exists on the hard drive
        If parentPDImage.locationOnDisk = "" Then parentPDImage.setSaveState False Else parentPDImage.setSaveState True
        
    'If the image was imported (e.g. from a scanner or the clipboard), note that it has NOT BEEN SAVED just because
    ' the user Undo'd all the way back to its original state.
    Else
        parentPDImage.setSaveState False
    End If
    
    'Select the relevant tool for this action, if we were supplied one previously
    If getUndoTool > -1 Then
        toolbar_Tools.selectNewTool getUndoTool
    End If
        
    'Launch the undo bitmap loading routine
    LoadUndo getCurrentUndoFilename(), getUndoProcessType
    
    'Notify the program that undo/redo data is no longer being processed.
    g_UndoRedoActive = False
    
End Sub

'When an action is canceled, use this to roll back any Undo data that was created for the canceled action
Public Sub rollBackLastUndo()

    'Some external actions are deactivated while Undo/Redo processing occurs.  Such functions can check this global
    ' variable to determine if they should fire - so set it before doing anything else.
    g_UndoRedoActive = True
    
    'Kill the Undo file created by the terminated action
    If FileExist(generateUndoFilename(undoMax)) Then Kill generateUndoFilename(undoMax)
    
    'Roll back the undo counter
    undoNum = undoNum - 1
    undoMax = undoMax - 1

    Debug.Print "Undo rolled back. No Redo information saved."
        
    'Rebuild the Undo array (not technically necessary, but it doesn't hurt)
    ReDim Preserve undoStack(0 To undoNum) As undoEntry
    
    'If Undos exist before this point, allow the user to access them, but if we are at the end of our Undo chain,
    ' disable further undo-ing.
    If undoNum > 0 Then undoState = True Else undoState = False
    
    'Regardless of Undo status, DO NOT ALLOW REDOING.  The Redo copy has been erased (because the action was terminated).
    redoState = False
    
    'If this image initially existed on the user's hard drive (e.g. it was not imported from the clipboard or
    ' some other source), note that the image has now BEEN SAVED (because it has not changed at all from its original
    ' appearance).
    If undoNum = undoNumAtLastSave Then
    
        'In the special case of UndoNumAtLastSave = 0, check to see if this image exists on the hard drive
        If parentPDImage.locationOnDisk = "" Then parentPDImage.setSaveState False Else parentPDImage.setSaveState True
        
    'If the image was imported (e.g. from a scanner or the clipboard), note that it has NOT BEEN SAVED just because
    ' the user Undo'd all the way back to its original state.
    Else
        parentPDImage.setSaveState False
    End If
    
    'Notify the program that undo/redo data is no longer being processed.
    g_UndoRedoActive = False

End Sub

'Restore an undo entry : "Redo"
Public Sub RestoreRedoData()
    
    'Some external actions are deactivated while Undo/Redo processing occurs.  Such functions can check this global
    ' variable to determine if they should fire - so set it before doing anything else.
    g_UndoRedoActive = True
    
    'If we're not at the top of the Undo stack, advance the stack index by one
    If undoNum < undoMax Then
    
        undoNum = undoNum + 1
        undoState = True
        
        'Allow saving contingent on the image's state
        If undoNum = undoNumAtLastSave Then parentPDImage.setSaveState True Else parentPDImage.setSaveState False
        
    End If
    
    'If we're at the end of our action chain, disable any further Redo-ing
    If undoNum = undoMax Then redoState = False Else redoState = True
    
    'Select the relevant tool for this action, if one is provided
    If getUndoTool > -1 Then toolbar_Tools.selectNewTool getUndoTool
        
    'Load the Redo bitmap file
    LoadUndo getCurrentUndoFilename(), getUndoProcessType, True
    
    'Notify the program that undo/redo data is no longer being processed.
    g_UndoRedoActive = False
    
End Sub

'Return the path of the last Undo file (used only for the "Fade last effect" function at present)
Public Function getLastUndoFile() As String
    getLastUndoFile = generateUndoFilename(undoNum - 2)
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
        generateUndoFilename = g_UserPreferences.getTempPath & "~cPDU_" & parentPDImage.imageID & "_" & uIndex & "_.pdtmp"
    End If
    
End Function

'Erase all undo files that this class has created
Public Sub clearUndos()

    'Remember that we have no Undo files for this image
    undoState = False
    
    'If Undo data has already been cleared, exit
    If undoTotal = 0 Then Exit Sub
    
    'Destroy every potential Undo file
    Dim i As Long
    For i = 0 To undoTotal
        
        'Check image data first...
        If FileExist(generateUndoFilename(i)) Then Kill generateUndoFilename(i)
        
        '...followed by selection data
        If FileExist(generateUndoFilename(i) & ".selection") Then Kill generateUndoFilename(i) & ".selection"
        
    Next i
    
    'Set the current and maximum Undo number to 0
    undoNum = 0
    undoMax = 0
    undoTotal = 0
    
End Sub

'Return the tool associated with the current Undo state
Public Function getUndoTool() As Long
    getUndoTool = undoStack(undoNum).relevantTool
End Function

'Return the process type associated with the current Undo state
Public Function getUndoProcessType() As Long
    If undoState Then
        getUndoProcessType = undoStack(undoNum - 1).undoType
    Else
        getUndoProcessType = -1
    End If
End Function

'Return the process type associated with the current Redo state
Public Function getRedoProcessType() As Long
    If redoState Then
        getRedoProcessType = undoStack(undoNum).undoType
    Else
        getRedoProcessType = -1
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

'Undo/redo data handling is complex.  Sometimes it is helpful to see how our Undo/Redo stack looks - use this sub
' to print out a full list to the debug window.
Public Sub debugUndo()

    Dim i As Long
    Debug.Print "---"
    For i = 0 To UBound(undoStack)
        If i = undoNum Then
            Debug.Print "* " & i & ": " & undoStack(i).processID & "," & undoStack(i).undoType
        Else
            Debug.Print i & ": " & undoStack(i).processID & "," & undoStack(i).undoType
        End If
    Next i

End Sub

Private Sub Class_Initialize()
    
    'Initialize all relevant undo/redo tracking variables
    ReDim undoStack(0) As undoEntry
    
    undoNum = 0
    undoMax = 0
    undoNumAtLastSave = 0
    undoState = False
    redoState = False
    
End Sub

'When this class is unloaded, we must make sure to delete any Undo files we created
Private Sub Class_Terminate()
    clearUndos
End Sub
