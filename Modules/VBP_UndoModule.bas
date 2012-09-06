Attribute VB_Name = "Undo_Handler"
'***************************************************************************
'Undo/Redo Handler
'Copyright ©2000-2012 by Tanner Helland
'Created: 2/April/01
'Last updated: 12/August/12
'Last update: BuildImageRestore requests are now required to supply the ID of the process requesting
'             Undo generation.  This is used to generate text related to undos (e.g. "Undo Blur")
'
'Handles all "Undo"/"Redo" operations.  I currently have it programmed to use the hard
' drive for all backups in order to free up RAM; this could be changed with in-memory images,
' but the speed delay is so insignificant that I opted to use the hard drive.
'
'IMPORTANT NOTE: the pdImages() array (of type pdImage) is declared in the
'                MDIWindow module.
'
'***************************************************************************

Option Explicit


'Create an Undo entry (save a copy of the present image to the temp directory)
' Required: the ID of the process that called this action
Public Sub CreateUndoFile(ByVal processID As Long)
    
    'All undo work is handled internally in the pdImage class
    Message "Saving Undo data..."
    pdImages(CurrentImage).BuildUndo processID
    
    'Since an undo exists, enable the Undo button and disable the Redo button
    tInit tUndo, pdImages(CurrentImage).UndoState
    tInit tRedo, pdImages(CurrentImage).RedoState
    FormMain.MnuFadeLastEffect.Enabled = pdImages(CurrentImage).UndoState

End Sub

'Restore an undo entry : "Undo"
Public Sub RestoreImage()
    
    'Let the internal pdImage Undo handler take care of any changes
    Message "Restoring Undo data..."
    pdImages(CurrentImage).Undo
    
    'Set the undo, redo, fade last effect buttons to their proper state
    tInit tUndo, pdImages(CurrentImage).UndoState
    tInit tRedo, pdImages(CurrentImage).RedoState
    FormMain.MnuFadeLastEffect.Enabled = pdImages(CurrentImage).UndoState
    
    'Launch the undo bitmap loading routine
    LoadUndo pdImages(CurrentImage).GetUndoFile
    
End Sub

'Erase every undo file for every open image
Public Sub ClearALLUndo()
    
    'Temporary value for tracking forms
    Dim CurWindow As Long
    
    'Loop through every form
    Dim tForm As Form
    For Each tForm In VB.Forms
        'If it's a valid image form...
        If tForm.Name = "FormImage" Then
            'Strip tag out
            CurWindow = val(tForm.Tag)
            'Clear the undos internally
            pdImages(CurWindow).ClearUndos
        End If
    Next

End Sub

'Clear all undo images for a single image
Public Sub ClearUndo(ByVal imageID As Long)

    'Tell this pdImage class to destroy all its Undo files
    pdImages(imageID).ClearUndos
    
    'If the active form is requesting the clear, adjust the Undo button/menu to match
    If imageID = CurrentImage Then
        tInit tUndo, pdImages(CurrentImage).UndoState
    
        'Also, disable fading any previous effects on this image (since there is no long an image to use for the function)
        FormMain.MnuFadeLastEffect.Enabled = pdImages(CurrentImage).UndoState
    End If

End Sub

'Restore an undo entry : "Redo"
Public Sub RedoImageRestore()
    
    'Let pdImage handle the Redo by itself
    Message "Restoring Redo data..."
    pdImages(CurrentImage).Redo
    
    'Set the undo, redo, fade last effect buttons to their proper state
    tInit tUndo, pdImages(CurrentImage).UndoState
    tInit tRedo, pdImages(CurrentImage).RedoState
    FormMain.MnuFadeLastEffect.Enabled = pdImages(CurrentImage).UndoState
        
    'Load the Redo bitmap file
    LoadUndo pdImages(CurrentImage).GetUndoFile
    
End Sub

'Subroutine for generating an Undo filename
Private Function GenerateUndoFile(ByVal uIndex As Integer) As String
    GenerateUndoFile = TempPath & "~cPDU" & CurrentImage & "_" & uIndex & ".tmp"
End Function

'Subroutine for returning the path of the last Undo file (used for fading last effect)
Public Function GetLastUndoFile() As String
    'Launch the undo loading routine
    GetLastUndoFile = GenerateUndoFile(pdImages(CurrentImage).UndoNum - 2)
End Function
