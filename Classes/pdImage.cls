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
'Copyright ©2000-2013 by Tanner Helland
'Created: sometime 2006
'Last updated: 28/September/12
'Last update: added a Selection object (mainSelection)
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
    'IDs of the processes associated with each undo file
    Private UndoIDs() As Long
    
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
    
    'Save flags let us save the image using settings the user has already specified (so we don't have to pester
    ' them every time they save)
    Private saveFlags(0 To 9) As Long
    
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
    
'Because arrays in a class module cannot be declared as Public, save flags must be read/stored via these two routines
Public Function getSaveFlag(ByVal fIndex As Long) As Long
    getSaveFlag = saveFlags(fIndex)
End Function

Public Sub setSaveFlag(ByVal fIndex As Long, ByVal fValue As Long)
    saveFlags(fIndex) = fValue
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

    tInit tSave, Not HasBeenSaved

End Sub
    
'UNDO routines:

'Create an undo entry (a copy of the present image); also required is the ID of the process that required this action.
' This is used to generate a string of the undo action's name (for example, "Undo Blur")
Public Sub BuildUndo(ByVal processID As Long)
    
    'Easy enough - just store a 24-bit bitmap of the current image
    mainLayer.writeToFile GenerateUndoFile(UndoNum)
    Debug.Print "Saved " & GenerateUndoFile(UndoNum) & " as Undo file #" & UndoNum & "."
    
    'Remember this processID for future use
    ReDim Preserve UndoIDs(0 To UndoNum) As Long
    UndoIDs(UndoNum) = processID
    
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
    
    'Mark this image as NOT saved (because if we're creating an Undo file, the image has changed from its saved version)
    HasBeenSaved = False
    tInit tSave, True
    
    Debug.Print "Undo created successfully."

End Sub
    
'Restore an undo entry, i.e. pushing the Undo button (formerly RestoreImage)
Public Sub Undo()
    
    'If we're at the end of actions, save this image for Redo potential
    If UndoNum = UndoMax Then
        mainLayer.writeToFile GenerateUndoFile(UndoNum)
        Debug.Print "Undo initiated by user. Saved " & GenerateUndoFile(UndoNum) & " for Redo purposes."
        'Mirror the current processID
        ReDim Preserve UndoIDs(0 To UndoNum) As Long
        UndoIDs(UndoNum) = UndoIDs(UndoNum - 1)
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
        tInit tSave, Not HasBeenSaved
    Else
        HasBeenSaved = False
        tInit tSave, Not HasBeenSaved
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
            tInit tSave, False
        Else
            HasBeenSaved = False
            tInit tSave, True
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
    For x = 0 To UndoTotal
        tempFile = GenerateUndoFile(x)
        If FileExist(tempFile) Then Kill tempFile
    Next x
    
    'Set the current and maximum Undo number to 0
    UndoNum = 0
    UndoMax = 0
    UndoTotal = 0
    
End Sub

'Return the processID associated with the current Undo state
Public Function getUndoProcessID() As Long
    
    'If we don't have a relevant Undo state, return -1.  Otherwise return the proper process ID.
    If UndoState = False Then
        getUndoProcessID = -1
    Else
        getUndoProcessID = UndoIDs(UndoNum - 1)
    End If
        
End Function

'Return the processID associated with the current Redo state
Public Function getRedoProcessID() As Long

    'If we don't have a relevant Redo state, return -1.  Otherwise return the proper process ID.
    If RedoState = False Then
        getRedoProcessID = -1
    Else
        getRedoProcessID = UndoIDs(UndoNum)
    End If

End Function

'Subroutine for generating an Undo/Redo filename
Private Function GenerateUndoFile(ByVal uIndex As Integer) As String
    GenerateUndoFile = userPreferences.getTempPath & "~cPDU" & imageID & "_" & uIndex & ".tmp"
End Function

'Public routine for getting the filename of the current Undo file
Public Function GetUndoFile()
    GetUndoFile = userPreferences.getTempPath & "~cPDU" & imageID & "_" & UndoNum & ".tmp"
End Function

'If the class's size has somehow changed, this can be called to update it.  At present, it simply copies over the size of
' the main layer, but in the future it may need to take into account multiple layers.
Public Function updateSize()
    Me.Width = mainLayer.getLayerWidth
    Me.Height = mainLayer.getLayerHeight
End Function

'See if a file exists or not
Private Function FileExist(fName As String) As Boolean
    On Error Resume Next
    Dim Temp As Long
    Temp = GetAttr(fName)
    FileExist = Not CBool(Err)
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
    ReDim UndoIDs(0 To 0) As Long
    
    'Initialize the main layer (currently, just the image data - layers aren't actually implemented)
    Set mainLayer = New pdLayer
    
    'Initialize the back buffer
    Set backBuffer = New pdLayer
    
    'Initialize the alpha composite layer
    Set alphaFixLayer = New pdLayer
    
    'Initialize the main selection
    Set mainSelection = New pdSelection
    selectionActive = False
    mainSelection.setSelectionType sRectangle
    Set mainSelection.containingPDImage = Me
        
End Sub
