Attribute VB_Name = "Undo_Handler"
'***************************************************************************
'Undo/Redo Handler
'Copyright ©2001-2013 by Tanner Helland
'Created: 2/April/01
'Last updated: 21/August/13
'Last update: new function to roll back the last created Undo.  This is necessary when the user cancels an action mid-processing.
'
'Handles all "Undo"/"Redo" operations.  I currently have it programmed to use the hard
' drive for all backups in order to free up RAM; this could be changed with in-memory images,
' but the speed delay is so insignificant that I opted to use the hard drive.
'
'IMPORTANT NOTE: the pdImages() array (of type pdImage) is declared in the
'                MDIWindow module.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit


'Create an Undo entry (save a copy of the present image or tool to the temp directory)
' Inputs:
'  1) the ID string of the process that called this action (e.g. something like "Gaussian blur")
'  2) optionally, the type of Undo that needs to be created.  By default, type 1 (image pixel undo) is assumed.
Public Sub CreateUndoData(ByVal processID As String, Optional ByVal undoType As Long = 1, Optional ByVal relevantTool As Long = -1)
    
    'All undo work is handled internally in the pdImage class
    Message "Saving Undo data..."
    pdImages(CurrentImage).BuildUndo processID, undoType, relevantTool
    
    'Since an undo exists, enable the Undo button and disable the Redo button
    metaToggle tUndo, pdImages(CurrentImage).UndoState
    metaToggle tRedo, pdImages(CurrentImage).RedoState
    
    '"Fade last effect" is reserved for filters and effects only
    If (undoType = 0) Or (undoType = 1) Then FormMain.MnuFadeLastEffect.Enabled = True Else FormMain.MnuFadeLastEffect.Enabled = False

End Sub

'Restore an undo entry : "Undo"
Public Sub RestoreUndoData()
    
    g_UndoRedoActive = True
    
    'Let the internal pdImage Undo handler take care of any changes
    Message "Restoring Undo data..."
    pdImages(CurrentImage).Undo
    
    'Select the relevant tool for this action, if relevant
    If pdImages(CurrentImage).getUndoTool > -1 Then
        FormMain.selectNewTool pdImages(CurrentImage).getUndoTool
    End If
    
    'Set the undo, redo, Fade last effect buttons to their proper state
    metaToggle tUndo, pdImages(CurrentImage).UndoState
    metaToggle tRedo, pdImages(CurrentImage).RedoState
    FormMain.MnuFadeLastEffect.Enabled = pdImages(CurrentImage).UndoState
        
    'Launch the undo bitmap loading routine
    LoadUndo pdImages(CurrentImage).GetUndoFile, pdImages(CurrentImage).getUndoProcessType
    
    'Check the Undo image's color depth, and check/uncheck the matching Image Mode setting
    If pdImages(CurrentImage).mainLayer.getLayerColorDepth() = 32 Then metaToggle tImgMode32bpp, True Else metaToggle tImgMode32bpp, False
    
    g_UndoRedoActive = False
    
End Sub

'When an action is canceled, we must rollback the undo data that was created (and is no longer needed)
Public Sub rollBackLastUndo()

    g_UndoRedoActive = True
    
    'Let the internal pdImage Undo handler take care of any changes
    Message "Removing unneeded undo data..."
    pdImages(CurrentImage).rollBackUndo
    
    'Set the undo, redo, Fade last effect buttons to their proper state
    metaToggle tUndo, pdImages(CurrentImage).UndoState
    metaToggle tRedo, pdImages(CurrentImage).RedoState
    FormMain.MnuFadeLastEffect.Enabled = pdImages(CurrentImage).UndoState
    
    'Check the Undo image's color depth, and check/uncheck the matching Image Mode setting
    If pdImages(CurrentImage).mainLayer.getLayerColorDepth() = 32 Then metaToggle tImgMode32bpp, True Else metaToggle tImgMode32bpp, False
    
    g_UndoRedoActive = False

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
            CurWindow = Val(tForm.Tag)
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
        metaToggle tUndo, pdImages(CurrentImage).UndoState
    
        'Also, disable fading any previous effects on this image (since there is no longer an image to use for the function)
        FormMain.MnuFadeLastEffect.Enabled = pdImages(CurrentImage).UndoState
    End If

End Sub

'Restore an undo entry : "Redo"
Public Sub RestoreRedoData()
    
    g_UndoRedoActive = True
    
    'Let pdImage handle the Redo by itself
    Message "Restoring Redo data..."
    pdImages(CurrentImage).Redo
    
    'Select the relevant tool for this action, if relevant
    If pdImages(CurrentImage).getUndoTool > -1 Then
        FormMain.selectNewTool pdImages(CurrentImage).getUndoTool
    End If
    
    'Set the undo, redo, Fade last effect buttons to their proper state
    metaToggle tUndo, pdImages(CurrentImage).UndoState
    metaToggle tRedo, pdImages(CurrentImage).RedoState
    FormMain.MnuFadeLastEffect.Enabled = pdImages(CurrentImage).UndoState
        
    'Load the Redo bitmap file
    LoadUndo pdImages(CurrentImage).GetUndoFile, pdImages(CurrentImage).getUndoProcessType
    
    'Finally, check the Redo image's color depth, and check/uncheck the matching Image Mode setting
    If pdImages(CurrentImage).mainLayer.getLayerColorDepth() = 32 Then metaToggle tImgMode32bpp, True Else metaToggle tImgMode32bpp, False
    
    g_UndoRedoActive = False
    
End Sub

'Subroutine for generating an Undo filename
Private Function GenerateUndoFile(ByVal uIndex As Integer) As String
    GenerateUndoFile = g_UserPreferences.getTempPath & "~cPDU" & CurrentImage & "_" & uIndex & ".tmp"
End Function

'Subroutine for returning the path of the last Undo file (used for fading last effect)
Public Function GetLastUndoFile() As String
    'Launch the undo loading routine
    GetLastUndoFile = GenerateUndoFile(pdImages(CurrentImage).UndoNum - 2)
End Function
