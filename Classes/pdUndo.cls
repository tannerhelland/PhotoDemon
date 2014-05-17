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
'Last updated: 16/May/14
'Last update: start work on a hideously complicated overhaul of the Undo / Redo engine.  The goal?  As much as
'             possible, make PD store only the relevant minimum amount of Undo data, as opposed to the full
'             pdImage stack.  No idea if this will work, but we're gonna try...
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
' class.  (Specifically, it is now handled by the separate syncInterfaceToCurrentImage function.)  As long as future
' updates supply valid data to the getUndoState and getRedoState functions, all interface stuff will be handled
' just fine by external interface functions.
'
'In 2014, as part of adding Layers support, the Undo/Redo engine was further modified from a lazy implementation to
' an active one.  What does this mean?  Basically, the previous "lazy" implementation did not create Undo/Redo data
' until absolutely necessary (e.g. after at least one image-changing action had been initiated, and right before a
' new actions was about to occur).  This ultra-conservative approach had some benefits; for example, Undo data was
' never created until an image was actually edited, improving image load time, and Redo data was never created
' unless Undo was actuall pressed.  While good for conserving memory, this system ultimately proved untenable, as it
' requires the storage of a lot of "what happened previously" data in order to handle Undo creation correctly for odd
' edge cases (like when an image edit is applied after a selection modification). To simplify the engine and make it
' possible to store Undo/Redo data other than just the full pdImage stack, Undo/Redo data is now created immediately
' for the active function, using the passed PD_UNDO_TYPE as its guide to what data should be stored.
'
'As an added benefit, the Active system also works much better with PD's Autosave engine.  It also simplifies things
' like the user cancelling an operation before it has completed, because instead of rolling back to some previous
' Undo state, the current action's changes simply don't get written to file.
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

'Information about each Undo entry is stored in an array; the array is dynamically resized as necessary when new
' Undos are created.  We track the ID of each action in preparation for a future History browser that allows the
' user to jump to any arbitrary Undo/Redo state.  (Also, to properly update the text of the Undo/Redo menu and
' buttons so the user knows which action they are undo/redoing.)
Private Type undoEntry
    processID As String             'Name of the associated action (e.g. "Gaussian blur")
    processParamString As String    'Processor string supplied to the action
    undoType As PD_UNDO_TYPE        'What type of Undo/Redo data was stored for this action (e.g. Image or Selection data)
    relevantTool As Long            'If a tool was associated with this action, it can be set here.  This value is not currently used.
End Type

Private undoStack() As undoEntry

'If the user performs 10 actions, then presses Undo 5 times and saves the image, we need to disable the Save button
' at that Undo point, because it was the last point at which the image was saved.  Thus we must track the last point
' in the Undo/Redo stack where the image was saved.
Private undoNumAtLastSave As Long

'We need access to the parent class so we can request things like "write your current data out to file".
Public parentPDImage As pdImage

'As part of PD's Undo/Redo chain, we can write the Undo/Redo stack to an XML file.  These constants are used to help
' us ID and parse such data correctly.
Private Const pdUndoVersionID As Long = 64

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

'Create an Undo entry (e.g. save a copy of the present image or selection to the temp directory)
' Inputs:
'  1) the ID string of the process that called this action (e.g. "Gaussian blur")
'  2) optionally, the type of Undo that needs to be created.  Make absolutely certain that this value is correct, because
'     it determines what type of Undo data is saved (e.g. full layers stack, single layer, selection, etc).
'  3) optionally, the tool associated with this action.  This value is not currently used, but in the future it could
'     be helpful for automatically selecting the tool associated with a given Undo/Redo point.
Public Sub createUndoData(ByVal processID As String, ByVal pParamString As String, ByVal processType As PD_UNDO_TYPE, Optional ByVal curTool As Long = -1)
    
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
    'Debug.Print "Saved " & getCurrentUndoFilename() & " as Undo file #" & undoNum & "."
    
    'Store all relevant information about this Undo action, so we can access it in the future if necessary
    ReDim Preserve undoStack(0 To undoNum) As undoEntry
    With undoStack(undoNum)
        .processID = processID
        .processParamString = pParamString
        .undoType = processType
        .relevantTool = curTool
    End With
        
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
    
    'If this is not the first Undo entry, make a note that the user is now allowed to initiate Undo actions.
    If undoNum > 1 Then undoState = True
    
    'Because we have just created Undo data, the user is not allowed to initiate Redo actions.
    redoState = False
    
    'If this action changed pixel data, mark the image as NOT being saved in its current state.
    If (processType <> UNDO_NOTHING) And (processType <> UNDO_SELECTION) Then parentPDImage.setSaveState False
    
    'New to version 6.4 is our persistent creation of Undo XML header files.  These are extraordinarily helpful
    ' for both regular Undo/Redo behavior, as well as AutoSave recreation.  We always create a copy of the Undo
    ' stack, in XML format, in the temporary folder.
    writeUndoDataToXML generateUndoHeaderFilename()
    
    'To inspect the current state of the Undo/Redo stack, uncomment this line.
    debugUndo "new Undo entry created"
    
    Debug.Print "Undo created successfully."
    
End Sub

'This special command can be used to revert the image to its last saved state.  For images that have never been
' saved (e.g. those pasted from the clipboard, etc), it will undo to the image's original state.
Public Sub revertToLastSavedState()
    If undoNumAtLastSave > 1 Then
        restoreUndoData undoNumAtLastSave
    Else
        restoreUndoData 1
    End If
End Sub

'Restore an undo entry, e.g. "the user just clicked Undo".
' Optionally, a specific position in the Undo/Redo chain can be specified if the user wants to revert to
' a specific state in an image's history.  (For example, the File -> Revert command calls this function
' with an undoPosition of 0.)
Public Sub restoreUndoData(Optional ByVal undoPosition As Long = -1)
    
    'Some external actions must be deactivated while Undo/Redo processing occurs.  Such asynchronous functions can check
    ' this global variable to determine if they should fire, so we must set it before doing anything else.
    g_UndoRedoActive = True
        
    'If the user supplied an undoPosition, apply it now.  Otherwise, decrement it from its current position.
    If undoPosition <> -1 Then
        undoNum = undoPosition
    Else
        undoNum = undoNum - 1
        If undoNum < 0 Then undoNum = 0
    End If
    
    'If this is not the last Undo state, still allow the user to Undo.
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
    If undoNum = undoNumAtLastSave Then
    
        'In the special case of UndoNumAtLastSave = 0, check to see if this image exists on the hard drive
        If Len(parentPDImage.locationOnDisk) > 0 Then parentPDImage.setSaveState True Else parentPDImage.setSaveState False
        
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
    LoadUndo generateUndoFilename(undoNum - 1), getUndoProcessType
    
    'To inspect the current state of the Undo/Redo stack, uncomment this line.
    debugUndo "user clicked Undo"
    
    'New to version 6.4 is our persistent creation of Undo XML header files.  These are extraordinarily helpful
    ' for both regular Undo/Redo behavior, as well as AutoSave recreation.
    writeUndoDataToXML generateUndoHeaderFilename()
    
    'Notify the program that undo/redo data is no longer being processed, so asynchronous functions can resume their behavior.
    g_UndoRedoActive = False
    
End Sub

'Restore a redo entry, e.g. "the user just clicked Redo"
Public Sub RestoreRedoData()
    
    'Some external actions must be deactivated while Undo/Redo processing occurs.  Such asynchronous functions can check
    ' this global variable to determine if they should fire, so we must set it before doing anything else.
    g_UndoRedoActive = True
    
    'If we're not at the top of the Undo stack, advance the stack index by one
    If undoNum < undoMax Then
        
        undoNum = undoNum + 1
        
        'Allow saving contingent on the image's state
        If undoNum = undoNumAtLastSave Then parentPDImage.setSaveState True Else parentPDImage.setSaveState False
    
    End If
    
    'Undoing is always allowed after the user has Redone something
    undoState = True
    
    'If we're at the end of our action chain, disable any further Redo-ing
    If undoNum = undoMax Then redoState = False Else redoState = True
    
    'Select the relevant tool for this action, if one is provided
    If getUndoTool > -1 Then toolbar_Tools.selectNewTool getUndoTool
    
    'Load the Redo bitmap file
    LoadUndo generateUndoFilename(undoNum - 1), getUndoProcessType, True
    
    'To inspect the current state of the Undo/Redo stack, uncomment this line.
    debugUndo "user clicked Redo"
    
    'New to version 6.4 is our persistent creation of Undo XML header files.  These are extraordinarily helpful
    ' for both regular Undo/Redo behavior, as well as AutoSave recreation.
    writeUndoDataToXML generateUndoHeaderFilename()
    
    'Notify the program that undo/redo data is no longer being processed, so asynchronous functions can resume their behavior.
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
        generateUndoFilename = g_UserPreferences.GetTempPath & "~cPDU_" & parentPDImage.imageID & "_" & uIndex & "_.pdtmp"
    End If
    
End Function

'Generate a unique Undo/Redo header filename, including full temp path.  Note that this function does *NOT* take an
' Undo number index.  The entire stack is written out to the header file, every time.
Private Function generateUndoHeaderFilename() As String
    
    'As a failsafe, all Undo/Redo files are erased when this class is unloaded.  The clearUndos function requires this
    ' function to generate Undo filnames.  If our parent class was terminated properly, it will have already erased
    ' our reference to it, so the parentPDImage.imageID reference below will fail - to prevent that, we must check the
    ' parent reference in advance.
    If Not (parentPDImage Is Nothing) Then
        generateUndoHeaderFilename = g_UserPreferences.GetTempPath & "~PDU_StackSummary_" & parentPDImage.imageID & "_.pdtmp"
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
    
    'Destroy the header file, if any
    If FileExist(generateUndoHeaderFilename()) Then Kill generateUndoHeaderFilename()
    
    'Set the current and maximum Undo number to 0
    undoNum = 0
    undoMax = 0
    undoTotal = 0
    
End Sub

'Return the tool associated with the current Undo state
Public Function getUndoTool() As Long
    getUndoTool = undoStack(undoNum - 1).relevantTool
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
            Debug.Print "  * " & i & ": " & undoStack(i).processID & ", " & undoStack(i).undoType
        Else
            Debug.Print "  " & i & ": " & undoStack(i).processID & ", " & undoStack(i).undoType
        End If
    Next i
    Debug.Print "---"

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

'Write the Undo stack out to file, in XML format
Public Sub writeUndoDataToXML(ByVal dstFilename As String)

    'Create a pdXML class, which will help us assemble the XML file
    Dim xmlEngine As pdXML
    Set xmlEngine = New pdXML
    xmlEngine.prepareNewXML "Undo stack"
    
    'Write out the XML version we're using for this macro
    xmlEngine.writeTag "pdUndoVersion", pdUndoVersionID
    
    'Write out the size of the current Undo stack, including the Max and Total values
    xmlEngine.writeTag "StackHeight", undoMax
    xmlEngine.writeTag "StackAbsoluteMaximum", undoTotal
    xmlEngine.writeTag "CurrentStackPointer", undoNum
    
    'We probably don't need the undoNumAtLastSave value, but it doesn't hurt to embed it, "just in case"
    xmlEngine.writeTag "UndoNumAtLastSave", undoNumAtLastSave
    
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
        xmlEngine.writeTag "RelevantTool", undoStack(i).relevantTool
        
        'Close this undo entry
        xmlEngine.closeTag "UndoEntry"
        xmlEngine.writeBlankLine
    
    Next i
    
    'With all Undo data successfully written, close the XML data and write it out to file.
    xmlEngine.writeXMLToFile dstFilename
    
End Sub

''With the creation of PD's Autosave engine, it became helpful to write some Undo immediately, outside of the
'' traditional Undo/Redo stack.  This function can be used to do so.
'Public Sub writeOneOffUndoEntry()
'
'    'Ask the parent PD image to write itself out to file
'    Saving.SavePhotoDemonImage parentPDImage, getCurrentUndoFilename(), True, False, False, False
'    parentPDImage.mainSelection.writeSelectionToFile getCurrentUndoFilename() & ".selection"
'
'    Debug.Print "Saved " & getCurrentUndoFilename() & " at the request of PD's Autosave engine."
'
'End Sub
