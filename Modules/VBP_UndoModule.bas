Attribute VB_Name = "UndoFunctions"
'***************************************************************************
'Undo/Redo Handler
'©2000-2012 Tanner Helland
'Created: 2/April/01
'Last updated: 20/June/12
'Last update: ClearUndo now requires an image ID parameter.  Because that can by triggered by a non-active form
'             (for example, when closing the program while many images are open - VB will not correctly fire the
'             Activate event as each child form is unloaded in turn), it is necessary to specify which image will
'             have its Undo cleared out.  This now guarantees that PhotoDemon won't leave temp files behind.
'
'Handles all "Undo"/"Redo" operations.  I currently have it programmed to use the hard
'drive for all backups in order to free up RAM; this could be changed with in-memory images,
'but the speed delay is so insignificant that I opted to use the hard drive.
'
'IMPORTANT NOTE: the pdImages() array (of type pdImage) is declared in the
'                MDIWindow module.
'
'***************************************************************************

Option Explicit


'Create an Undo entry (save a copy of the present image to the temp directory)
Public Sub BuildImageRestore()
    
    'All undo work is handled internally in the pdImage class
    Message "Creating Undo data..."
    pdImages(CurrentImage).BuildUndo
    
    'Since an undo exists, enable the Undo button and disable the Redo button
    tInit tUndo, pdImages(CurrentImage).UndoState
    tInit tRedo, pdImages(CurrentImage).RedoState
    FormMain.MnuFadeLastEffect.Enabled = pdImages(CurrentImage).UndoState

End Sub

'Restore an undo entry (i.e. ctrl+z)
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

'Run through every form and wipe every undo file we can find
Public Sub ClearALLUndo()
    
    'Temporary value for tracking forms
    Dim CurWindow As Long
    
    'Loop through every form
    Dim tForm As Form
    For Each tForm In VB.Forms
        'If it's a valid image form...
        If tForm.Name = "FormImage" Then
            'Strip tag out
            CurWindow = Val(tForm.Tag)
            'Clear the undos internally
            pdImages(CurWindow).ClearUndos
        End If
    Next

End Sub

'Clear all known undo images from the temporary directory for the current image
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

'Subroutine for returning the path of the last Undo file (used for fading)
Public Function GetLastUndoFile() As String
    'Launch the undo loading routine
    GetLastUndoFile = GenerateUndoFile(pdImages(CurrentImage).UndoNum - 2)
End Function
