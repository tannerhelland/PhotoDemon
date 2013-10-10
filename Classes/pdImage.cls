VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Image class
'Copyright ©2006-2013 by Tanner Helland
'Created: sometime 2006
'Last updated: 21/August/13
'Last update: added the ability to roll back the Undo data for a canceled operation
'
'The pdImage class is used to store information on each image loaded by the user.  One copy of this
' class exists for each loaded image.  These copies are stored in the pdImages() array, which is declared
' in the MDIWindow module.
'
'The purpose of this class is to store a bunch of variables related to a given image - things like size, zoom,
' name, file path, format containing form, Undo/Redo tracking - as well as a reference to a pdLayer object
' called "mainLayer", which contains an in-memory copy of the relevant image's data.  In the future,
' additional layers may be utilized within each image.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************


Option Explicit


'This class is designed to store all the information about a particular image.
'It currently includes the following information, per image:

'Image ID (the value of NumOfImagesLoaded when the class was initialized - see the CreateNewImageForm routine)
'Is it active? (e.g., still associated with a form that has not been unloaded)
'Did the image load successfully? (this defaults to false, and must be set to True by the load function)
'Containing form (e.g. upon which FormImage does this pdImage appear?)
'Original file size
'Original file name (e.g. the file name used when the image was first loaded)
'Original file name + extension (e.g. original file name + .***)
'Location on disk (full path, from the drive all the way to the filename + extension)
'Original file format (a numerical value corresponding to the relevant FREE_IMAGE_FORMAT enum)
'Original color depth
'Undo information, specifically:
    'Current Undo state (e.g. have undo files been created?)
    'Current Redo state (e.g. has the user pressed the undo button, meaning it's possible to redo actions?)
    'Number of Undo files created
    'Position in the current Undo chain (the number moves down as the Undo button is pressed, up as Redo is pressed or new actions are performed)
    'Array containing the Process IDs of each action at each Undo value.  This is used to rebuild the Undo and Redo text with the name of the action performed.
'Current size
'Current zoom level
'Two save flags (e.g. which format the image was saved in, and any attached parameters - such as JPEG quality)

'Layer's image data (stored as a DIB inside the mainLayer object)

    Public imageID As Long

    'Is the form associated with this image still loaded?
    Public IsActive As Boolean

    'Was the image loaded successfully?
    Public loadedSuccessfully As Boolean

    'The ID of the MDI child form containing this particular image
    Public containingForm As Form
    
    'Image dimensions
    Public Width As Long
    Public Height As Long
    
    'Which Undo number this image is currently on
    Public UndoNum As Integer
    'How many Undo levels exist
    Public UndoMax As Integer
    Public UndoTotal As Integer
    'State of undo/redo buttons
    Public UndoState As Boolean
    Public RedoState As Boolean
    
    'Information about each Undo entry is stored in an array; the array is dynamically resized as necessary when new Undos are created
    Private Type undoEntry
        Id As String
        Type As Long
        relevantTool As Long
    End Type
    
    Private Undos() As undoEntry
    
    'Current index of the zoom combo box
    Public CurrentZoomValue As Long
    
    'Size of the file when loaded
    Public OriginalFileSize As Long
    
    'Original name of the image
    Public OriginalFileName As String
    
    'Original name of the image AND its extension
    Public OriginalFileNameAndExtension As String
    
    'Original path of the image (full path, including filename); blank if image does not exist on disk
    Public LocationOnDisk As String
    
    'Original image format.  This is derived from the equivalent FreeImage constant - see the enum for FREE_IMAGE_FORMAT for details
    Public OriginalFileFormat As Long
    
    'Current file format.  For example, if the user loads a .BMP file and uses "Save As" to save it as a .JPG, this variable will
    ' represent "JPG" - while OriginalFileFormat will still represent "BMP".  (OriginalFileFormat is only set once, at load-time.)
    Public CurrentFileFormat As Long
    
    'Original color depth (a BPP value, most commonly 24 or 32, may be 8)
    Public OriginalColorDepth As Long
    
    'saveParameters let us save the image using settings the user has already specified (so we don't have to pester
    ' them every time they save)
    Public saveParameters As String
    
    'For JPEGs specifically, has a quality value been set?  If it hasn't, the user needs to be prompted at least once.
    Public hasSeenJPEGPrompt As Boolean
    
    'For JPEG-2000 files, has a compression ratio been set?  If it hasn't, the user needs to be prompted at least once.
    Public hasSeenJP2Prompt As Boolean
    
    'For PNG files, if the original PNG file has a background color, it will be stored here.
    Public pngBackgroundColor As Long
    
    'Has this image been saved?
    Public HasBeenSaved As Boolean
    Public UndoNumAtLastSave As Integer
    
    'These two variables are only used at form load; they can be safely ignored once the image is visible.
    Public WindowLeft As Long
    Public WindowTop As Long
    
    'Track the current window state; if the image's containing form goes from maximized to normal, all images need their
    ' buffer refreshed (as scroll bars may then be needed)
    Public WindowState As Long
    
    'Variables related to drawing the buffer on-screen; these are set by PrepareViewport and read by ScrollViewport
    Public targetWidth As Long
    Public targetHeight As Long     'Formerly the width and height of the .FrontBuffer picture box
    Public targetTop As Long
    Public targetLeft As Long       'Formerly the left and top of the .FrontBuffer picture box
    
    'A copy of the form's current icon
    Public curFormIcon32 As Long, curFormIcon16 As Long
    
    'If this image is simply being used for internal PhotoDemon processing, hide it from the user
    Public forInternalUseOnly As Boolean
    
    'Image's actual image data (stored as a DIB inside the mainLayer object)
    Public mainLayer As pdLayer
    
    'Image's back buffer (also a DIB; this is used to hold the final composite before everything is rendered on-screen
    Public backBuffer As pdLayer
    
    'If this image is 32bpp then it requires a specialized compositing layer for rendering the transparency
    Public alphaFixLayer As pdLayer
    
    'Image's selection data (stored and handled by a mainSelection object)
    Public mainSelection As pdSelection
    Public selectionActive As Boolean
    
    'Metadata handler/storage class.  This class is automatically filled with an image file's metadata at load-time.
    ' NOTE: EXIFTOOL IS REQUIRED FOR METADATA HANDLING.
    Public imgMetadata As pdMetadata
    
    'This image's location in the window manager.  We cache this value so we can quickly access information about this image's window.
    Public indexInWindowManager As Long
    
'When this layer is no longer being used, we can deactivate it to save on resources.
Public Sub deactivateImage()
    
    'Erase the main layer
    If Not (mainLayer Is Nothing) Then
        mainLayer.eraseLayer
        Set mainLayer = Nothing
    End If
    
    'Erase the back buffer
    If Not (backBuffer Is Nothing) Then
        backBuffer.eraseLayer
        Set backBuffer = Nothing
    End If
    
    'If the alpha adjustment layer exists, erase it too
    If Not (alphaFixLayer Is Nothing) Then
        alphaFixLayer.eraseLayer
        Set alphaFixLayer = Nothing
    End If
    
    'If a selection exists, wipe it
    If Not (mainSelection Is Nothing) Then Set mainSelection = Nothing
    
    'Remove the reference to the parent form so VB can wipe it
    If Not (containingForm Is Nothing) Then Set containingForm = Nothing
    
    'It doesn't make a big difference, but we can also empty out this image's String-type variables to save a bit of space.
    OriginalFileName = ""
    OriginalFileNameAndExtension = ""
    LocationOnDisk = ""
    
    'Mark this image as inactive
    IsActive = False
    
End Sub
    
'Outside actions (such as saving) can affect the HasBeenSaved variable.  However, because we need to do additional
' processing based on the state of this variable, we provide this interface
Public Sub UpdateSaveState(ByVal newSaveState As Boolean)

    If newSaveState = True Then
        HasBeenSaved = True
        
        'Remember the undo value at this juncture; if the user performs additional actions, but "Undos" to this point,
        ' we still want the save button disabled
        UndoNumAtLastSave = UndoNum
    Else
        HasBeenSaved = False
    End If

    metaToggle tSave, Not HasBeenSaved

End Sub
    
'UNDO routines:

'Create an undo entry (a copy of the present image); also required is the ID of the process that required this action.
' This is used to generate a string of the undo action's name (for example, "Undo Blur")
Public Sub BuildUndo(ByVal processID As String, ByVal processType As Long, ByVal curTool As Long)
    
    'Depending on the type of process being applied, save out image or selection data to file
    'Select Case processType
    
        'Pixel data
        'Case 1
            mainLayer.writeToFile GenerateUndoFile(UndoNum), False
            
        'Selection data
        'Case 2
            mainSelection.writeSelectionToFile GenerateUndoFile(UndoNum) & ".selection"
        
    'End Select
    
    Debug.Print "Saved " & GenerateUndoFile(UndoNum) & " as Undo file #" & UndoNum & "."
    
    'Remember this processID for future use
    ReDim Preserve Undos(0 To UndoNum) As undoEntry
    With Undos(UndoNum)
        .Id = processID
        .Type = processType
        .relevantTool = curTool
    End With
    debugUndo
    'If the present image is NOT the last image in the undo train (i.e., an Undo
    'was fired before running the current operation) treat this as the new "last"
    UndoMax = UndoNum
    
    'Increase the number of Undos we recognize as existant
    UndoNum = UndoNum + 1
    UndoMax = UndoMax + 1
    
    'If the present image is legitimately the last one in the Undo train, note that we have more Undo files
    ' to delete when the program terminates.
    If UndoMax > UndoTotal Then UndoTotal = UndoMax
    
    'Allow the user to Undo, but no longer Redo
    UndoState = True
    RedoState = False
    
    'If this action changed pixel data, mark the image as NOT saved (because if we're creating an Undo file, the image has changed
    ' from its saved version)
    If processType = 1 Then
        HasBeenSaved = False
        metaToggle tSave, True
    End If

    Debug.Print "Undo created successfully."

End Sub
    
'Restore an undo entry, i.e. pushing the Undo button (formerly RestoreImage)
Public Sub Undo()
    
    'If we're at the end of actions, save this image for Redo potential
    If UndoNum = UndoMax Then
        
        'Depending on the last action, write out a Redo file for that action
        'Select Case UndoTypes(UndoNum - 1)
        
            'Pixel undo data
            'Case 1
                mainLayer.writeToFile GenerateUndoFile(UndoNum), False
            
            'Selection undo data
            'Case 2
                mainSelection.writeSelectionToFile GenerateUndoFile(UndoNum) & ".selection"
            
        'End Select
        
        Debug.Print "Undo initiated by user. Saved " & GenerateUndoFile(UndoNum) & " for Redo purposes."
        
        'Move neighboring values into the current Undo location (this makes Redo work properly)
        ReDim Preserve Undos(0 To UndoNum) As undoEntry
        With Undos(UndoNum)
            .Id = Undos(UndoNum - 1).Id
            .Type = Undos(UndoNum - 1).Type
            .relevantTool = Undos(UndoNum - 1).relevantTool
        End With
        
    End If
    
    'If this is not the last Undo state, decrease the number of Undos available
    If UndoNum > 0 Then
        UndoNum = UndoNum - 1
        UndoState = True
        RedoState = True
    End If
    
    'If we're now at the end of our Undo chain, disable further undo-ing
    If UndoNum = 0 Then UndoState = False
    
            
    'Also, if this image initially existed on the user's hard drive (e.g. it was not imported from the clipboard or
    ' some other source), note that the image has now BEEN SAVED (because it has not been changed at all). If the image
    ' was imported from somewhere, it has NOT BEEN SAVED just because the user Undo'd all the way back to its
    ' original state.
    If UndoNum = UndoNumAtLastSave Then
    
        'In the special case of UndoNumAtLastSave = 0, check to see if this image exists on the hard drive
        If LocationOnDisk = "" Then HasBeenSaved = False Else HasBeenSaved = True
        metaToggle tSave, Not HasBeenSaved
    Else
        HasBeenSaved = False
        metaToggle tSave, Not HasBeenSaved
    End If
    
End Sub

'When the user cancels an action, it is necessary to rollback the previous Undo.  This is slightly different from
' a normal Undo action, so it requires its own function.
Public Sub rollBackUndo()

    'Kill the Undo file created by the terminated action
    If FileExist(GenerateUndoFile(UndoMax)) Then Kill GenerateUndoFile(UndoMax)
    
    'Roll back the undo counter
    UndoNum = UndoNum - 1
    UndoMax = UndoMax - 1

    Debug.Print "Undo rolled back. No Redo information saved."
        
    'Rebuild the Undo array (not technically necessary, but it doesn't hurt)
    ReDim Preserve Undos(0 To UndoNum) As undoEntry
    
    'If Undos exist before this point, allow them
    If UndoNum > 0 Then UndoState = True
    
    'If we're now at the end of our Undo chain, disable further undo-ing
    If UndoNum = 0 Then UndoState = False
    
    'Regardless of Undo status, DO NOT ALLOW REDOING.  The Redo copy has been erased (because the action was terminated)
    RedoState = False
    
    'Similar to regular Undo, if this image initially existed on the user's hard drive (e.g. it was not imported from
    ' the clipboard or some other source), note that the image has now BEEN SAVED (because it has not been changed).
    ' If the image was imported from somewhere, it has still NOT BEEN SAVED, because the file doesn't exist on the
    ' hard drive (that PD knows of, anyway).
    If UndoNum = UndoNumAtLastSave Then
    
        'In the special case of UndoNumAtLastSave = 0, check to see if this image exists on the hard drive
        If LocationOnDisk = "" Then HasBeenSaved = False Else HasBeenSaved = True
        metaToggle tSave, Not HasBeenSaved
    Else
        HasBeenSaved = False
        metaToggle tSave, Not HasBeenSaved
    End If
    
End Sub

'Restore an undo entry : "Redo"
Public Sub Redo()
    
    'If we're not at the end of our action chain, advance us by 1
    If UndoNum < UndoMax Then
        UndoNum = UndoNum + 1
        UndoState = True
        
        'Allow saving
        If UndoNum = UndoNumAtLastSave Then
            HasBeenSaved = True
            metaToggle tSave, False
        Else
            HasBeenSaved = False
            metaToggle tSave, True
        End If
        
    End If
    
    'If we're at the end of our action chain, disable any further Redo-ing
    If UndoNum = UndoMax Then
        RedoState = False
    Else
        RedoState = True
    End If
    
End Sub

'Clear this image's Undo files
Public Sub ClearUndos()

    'Remember that we have no Undo files for this image
    UndoState = False
    
    'Using a temporary string, destroy every possible Undo file
    Dim tempFile As String
    
    Dim x As Long
    For x = 0 To UndoTotal
        tempFile = GenerateUndoFile(x)
        If FileExist(tempFile) Then Kill tempFile
        tempFile = GenerateUndoFile(x) & ".selection"
        If FileExist(tempFile) Then Kill tempFile
    Next x
    
    'Set the current and maximum Undo number to 0
    UndoNum = 0
    UndoMax = 0
    UndoTotal = 0
    
End Sub

'Return the tool associated with the current Undo state
Public Function getUndoTool() As Long
    getUndoTool = Undos(UndoNum).relevantTool
End Function

'Return the process type associated with the current Undo state
Public Function getUndoProcessType() As Long
    If UndoState = False Then
        getUndoProcessType = 0
    Else
        getUndoProcessType = Undos(UndoNum).Type
    End If
End Function

'Return the process type associated with the current Redo state
Public Function getRedoProcessType() As Long

    'If we don't have a relevant Redo state, return nothing.  Otherwise return the proper process ID.
    If RedoState = False Then
        getRedoProcessType = 0
    Else
        getRedoProcessType = Undos(UndoNum).Type
    End If

End Function

'Return the processID associated with the current Undo state
Public Function getUndoProcessID() As String
    
    'If we don't have a relevant Undo state, return -1.  Otherwise return the proper process ID.
    If UndoState = False Then
        getUndoProcessID = ""
    Else
        getUndoProcessID = g_Language.TranslateMessage(Undos(UndoNum - 1).Id)
    End If
        
End Function

'Return the processID associated with the current Redo state
Public Function getRedoProcessID() As String

    'If we don't have a relevant Redo state, return -1.  Otherwise return the proper process ID.
    If RedoState = False Then
        getRedoProcessID = ""
    Else
        getRedoProcessID = g_Language.TranslateMessage(Undos(UndoNum).Id)
    End If

End Function

'Subroutine for generating an Undo/Redo filename
Private Function GenerateUndoFile(ByVal uIndex As Long) As String
    GenerateUndoFile = g_UserPreferences.getTempPath & "~cPDU" & imageID & "_" & uIndex & ".tmp"
End Function

'Public routine for getting the filename of the current Undo file
Public Function GetUndoFile()
    GetUndoFile = g_UserPreferences.getTempPath & "~cPDU" & imageID & "_" & UndoNum & ".tmp"
End Function

'If the class's size has somehow changed, this can be called to update it.  At present, it simply copies over the size of
' the main layer, but in the future it may need to take into account multiple layers.
Public Function updateSize()
    Me.Width = mainLayer.getLayerWidth
    Me.Height = mainLayer.getLayerHeight
End Function

'INITIALIZE class
Private Sub Class_Initialize()

    'Initially, mark the image as *not* having been saved
    IsActive = False
    loadedSuccessfully = False
    HasBeenSaved = False
    UndoNumAtLastSave = 0
    forInternalUseOnly = False
    hasSeenJPEGPrompt = False
    pngBackgroundColor = -1
    
    'Set up the process ID array
    ReDim Undos(0) As undoEntry
    
    'Initialize the main layer (currently, just the image data - layers aren't actually implemented)
    Set mainLayer = New pdLayer
    
    'Initialize the back buffer
    Set backBuffer = New pdLayer
    
    'Initialize the alpha composite layer
    Set alphaFixLayer = New pdLayer
    
    'Initialize the main selection
    Set mainSelection = New pdSelection
    selectionActive = False
    mainSelection.setSelectionShape sRectangle
    Set mainSelection.containingPDImage = Me
    
    'Initialize the metadata object (which may not get used, but this prevents errors if other functions try to access metadata)
    Set imgMetadata = New pdMetadata
        
End Sub

'Return a layer that contains the currently selected area, fully processed according to the selection mask
Public Function retrieveProcessedSelection(ByRef dstLayer As pdLayer) As Boolean

    'If this image does not contain an active selection, exit now.
    If (Not IsActive) Or (Not selectionActive) Then
        retrieveProcessedSelection = False
        Exit Function
    End If
    
    'Selections can be one of several types.  Right now, we don't give special handling to simple rectangular selections - all selections
    ' are fully processed according to the contents of the mask.  Also, all selections are currently created as 32bpp layers.
    
    'Start by initializing the destination layer to the size of the active selection
    Set dstLayer = New pdLayer
    dstLayer.createBlank mainSelection.boundWidth, mainSelection.boundHeight, 32
    
    'We now need pointers to three separate sets of image data: destination layer, source layer, and selection mask.
    Dim srcImageData() As Byte
    Dim srcSA As SAFEARRAY2D
    prepSafeArray srcSA, mainLayer
    CopyMemory ByVal VarPtrArray(srcImageData()), VarPtr(srcSA), 4
    
    Dim selData() As Byte
    Dim selSA As SAFEARRAY2D
    prepSafeArray selSA, mainSelection.selMask
    CopyMemory ByVal VarPtrArray(selData()), VarPtr(selSA), 4
            
    Dim dstImageData() As Byte
    Dim dstSA As SAFEARRAY2D
    prepSafeArray dstSA, dstLayer
    CopyMemory ByVal VarPtrArray(dstImageData()), VarPtr(dstSA), 4
    
    Dim leftOffset As Long, topOffset As Long
    leftOffset = mainSelection.boundLeft
    topOffset = mainSelection.boundTop
    
    Dim x As Long, y As Long
    Dim thisAlpha As Long
    Dim blendAlpha As Double
    
    Dim srcLayerCD As Long, srcQuickX As Long, srcQuickY As Long
    srcLayerCD = mainLayer.getLayerColorDepth \ 8
    
    Dim dstLayerCD As Long, dstQuickX As Long
    dstLayerCD = dstLayer.getLayerColorDepth \ 8
            
    For x = 0 To dstLayer.getLayerWidth - 1
        dstQuickX = x * dstLayerCD
        srcQuickX = (leftOffset + x) * srcLayerCD
    For y = 0 To dstLayer.getLayerHeight - 1
    
        srcQuickY = topOffset + y
        thisAlpha = selData((leftOffset + x) * 3, srcQuickY)
        
        dstImageData(dstQuickX, y) = srcImageData(srcQuickX, srcQuickY)
        dstImageData(dstQuickX + 1, y) = srcImageData(srcQuickX + 1, srcQuickY)
        dstImageData(dstQuickX + 2, y) = srcImageData(srcQuickX + 2, srcQuickY)
        If srcLayerCD = 4 Then
            blendAlpha = thisAlpha / 255
            dstImageData(dstQuickX + 3, y) = srcImageData(srcQuickX + 3, srcQuickY) * blendAlpha
        Else
            dstImageData(dstQuickX + 3, y) = thisAlpha
        End If
        
    Next y
    Next x
    
    'Clear all array references
    CopyMemory ByVal VarPtrArray(srcImageData), 0&, 4
    CopyMemory ByVal VarPtrArray(dstImageData), 0&, 4
    CopyMemory ByVal VarPtrArray(selData), 0&, 4
    
    Erase srcImageData
    Erase dstImageData
    Erase selData
    
    retrieveProcessedSelection = True

End Function


'Temporary sub for debugging undo data
Public Sub debugUndo()

    Dim i As Long
    Debug.Print "---"
    For i = 0 To UBound(Undos) '- 1
        If i = UndoNum Then
            Debug.Print "* " & i & ": " & Undos(i).Id & "," & Undos(i).Type
        Else
            Debug.Print i & ": " & Undos(i).Id & "," & Undos(i).Type
        End If
    Next i

End Sub

