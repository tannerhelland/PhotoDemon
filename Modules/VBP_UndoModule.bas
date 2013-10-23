Attribute VB_Name = "Undo_Handler"
'***************************************************************************
'Undo/Redo Handler
'Copyright ©2001-2013 by Tanner Helland
'Created: 2/April/01
'Last updated: 02/October/13
'Last update: do not display messages when saving/restoring Undo data.
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

'Erase every undo file for every open image
Public Sub ClearALLUndo()
    
    'Temporary value for tracking forms
    Dim curWindow As Long
    
    'Loop through every form
    Dim tForm As Form
    For Each tForm In VB.Forms
        'If it's a valid image form...
        If tForm.Name = "FormImage" Then
            'Strip tag out
            curWindow = Val(tForm.Tag)
            'Clear the undos internally
            pdImages(curWindow).undoManager.clearUndos
        End If
    Next

End Sub




