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
'©2000-2012 Tanner Helland
'Created: sometime 2006
'Last updated: 07/July/12
'Last update: implemented HasBeenSaved tracking
'
'The pdImage class is used to store information on each image loaded by the user.  One copy of this
' class exists for each loaded image.  These copies are stored in the pdImages() array, which is declared
' in the MDIWindow module.
'
'At present, the main purpose of this class is to hold a bunch of variables related to a given image -
' things like size, zoom, name, file path, containing form, Undo/Redo tracking.  In the future, I'd like
' for this class to reference an in-memory copy of an image's data, which would finally break PhotoDemon's
' reliance on picture boxes for storing image data (and would allow things like deep color support).
'
'***************************************************************************


Option Explicit


'This class is designed to store all the information about a particular image
'It currently includes the following information, per image:

'Image ID (the value of NumOfImagesLoaded when the class was initialized - see the CreateNewImageForm routine)
'Whether or not it is active (e.g., still tied to a form that has not been unloaded)
'Containing form (e.g. upon which FormImage does this pdImage appear?)
'Original file size
'Original file name (e.g. the file name used when the image was first loaded)
'Original file name + extension (e.g. original file name + .***)
'Location on disk (full path, from the drive all the way to the filename + extension)
'Undo information, specifically:
    'Current Undo state (e.g. have undo files been created?)
    'Current Redo state (e.g. has the user pressed the undo button, meaning it's possible to redo actions?)
    'Number of Undo files created
    'Position in the current Undo chain (the number moves down as the Undo button is pressed)
'Current size
'Current zoom level
'Two save flags (e.g. which format the image was saved in, and any attached parameters - such as JPEG quality)

'Eventually, it might be nice to add the following information:
' Full image data (instead of storing it in a picture box, as we do now)

    Public ImageID As Long

    'Is the form associated with this image still loaded?
    Public IsActive As Boolean

    'The ID of the MDI child form containing this particular image
    Public containingForm As Form
    
    'Image dimensions
    Public PicWidth As Long
    Public PicHeight As Long
    
    'Which Undo number this image is currently on
    Public UndoNum As Integer
    'How many Undo levels exist
    Public UndoMax As Integer
    Public UndoTotal As Integer
    'State of undo/redo buttons
    Public UndoState As Boolean
    Public RedoState As Boolean
    
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
    
    'Save flags let us save the image using settings the user has already specified (so we don't have to pester
    ' them every time they save)
    Public saveFlag0 As Long
    Public saveFlag1 As Long
    
    'Has this image been saved?
    Public HasBeenSaved As Boolean
    Public UndoNumAtLastSave As Integer
    
    'These two variables are only used at form load; they can be safely ignored once the image is visible.
    Public WindowLeft As Long
    Public WindowTop As Long
    
    'Variables related to drawing the buffer on-screen; these are set by PrepareViewport and read by ScrollViewport
    Public targetWidth As Long
    Public targetHeight As Long       'Formerly the width and height of the .FrontBuffer picture box
    Public targetTop As Long
    Public targetLeft As Long               'Formerly the left and top of the .FrontBuffer picture box
    
    'If this image is simply being used for internal PhotoDemon processing, hide it from the user
    Public forInternalUseOnly As Boolean
    
'Outside actions (such as saving) can affect the HasBeenSaved variable.  However, because we need to do additional
' processing based on the state of this variable, we provide this interface
Public Sub UpdateSaveState(ByVal newSaveState As Boolean)

    If newSaveState = True Then
        HasBeenSaved = True
        UndoNumAtLastSave = UndoNum '- 1
    Else
        HasBeenSaved = False
    End If

    tInit tSave, Not HasBeenSaved

End Sub
    
'UNDO routines:

'Create an undo entry (a copy of the present image) (formerly BuildImageRestore)
Public Sub BuildUndo()
    
    'Easy enough - just store a 24-bit bitmap of the current image
    SavePicture FormMain.ActiveForm.BackBuffer.Picture, GenerateUndoFile(UndoNum)
    Debug.Print "Saved " & GenerateUndoFile(UndoNum) & " as Undo file #" & UndoNum & "."
    
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
    
    'Mark this image as NOT saved (because if we're creating an Undo file, the image is being changed somehow)
    HasBeenSaved = False
    tInit tSave, True
    
    Debug.Print "Undo created successfully."

    'MsgBox UndoNum & "," & UndoMax & "," & UndoTotal & "," & vbCrLf & UndoState & "," & RedoState

End Sub
    
'Restore an undo entry, i.e. pushing the Undo button (formerly RestoreImage)
Public Sub Undo()
    
    'If we're at the end of actions, save this image for Redo potential
    If UndoNum = UndoMax Then
        SavePicture FormMain.ActiveForm.BackBuffer.Picture, GenerateUndoFile(UndoNum)
        Debug.Print "Undo initiated by user. Saved " & GenerateUndoFile(UndoNum) & " for Redo purposes."
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
    ' was imported from somewhere, it HAS NOT BEEN SAVED just because the user Undo'd all the way back to its
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

'Subroutine for generating an Undo/Redo filename
Private Function GenerateUndoFile(ByVal uIndex As Integer) As String
    GenerateUndoFile = TempPath & "~cPDU" & ImageID & "_" & uIndex & ".tmp"
End Function

'Public routine for getting the filename of the current Undo file
Public Function GetUndoFile()
    GetUndoFile = TempPath & "~cPDU" & ImageID & "_" & UndoNum & ".tmp"
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
    HasBeenSaved = False
    UndoNumAtLastSave = 0
    forInternalUseOnly = False
    
End Sub
