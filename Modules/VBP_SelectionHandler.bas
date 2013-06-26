Attribute VB_Name = "Selection_Handler"
'***************************************************************************
'Selection Interface
'Copyright ©2012-2013 by Tanner Helland
'Created: 21/June/13
'Last updated: 26/June/13
'Last update: user-facing implementations of load/save selections to file
'
'Selection tools have existed in PhotoDemon for awhile, but this module is the first to support Process varieties of
' selection operations - e.g. internal actions like "Process "Create Selection"".  Selection commands must be passed
' through the Process module so they can be recorded as macros, and as part of the program's Undo/Redo chain.  This
' module provides all selection-related functions that the Process module can call.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://www.tannerhelland.com/photodemon/#license
'
'***************************************************************************

Option Explicit

'Create a new selection using the settings stored in a pdParamString-compatible string
Public Sub CreateNewSelection(ByVal paramString As String)
    
    'Use the passed parameter string to initialize the selection
    pdImages(CurrentImage).mainSelection.initFromParamString paramString
    pdImages(CurrentImage).mainSelection.lockIn pdImages(CurrentImage).containingForm
    pdImages(CurrentImage).selectionActive = True
    
    'Change the selection-related menu items to match
    tInit tSelection, True
    
    'Draw the new selection to the screen
    RenderViewport pdImages(CurrentImage).containingForm

End Sub

'Create a new selection using the settings stored in a pdParamString-compatible string
Public Sub RemoveCurrentSelection(Optional ByVal paramString As String)
    
    'Use the passed parameter string to initialize the selection
    pdImages(CurrentImage).mainSelection.lockRelease
    pdImages(CurrentImage).selectionActive = False
    
    'Change the selection-related menu items to match
    tInit tSelection, False
    
    'Redraw the image (with selection removed)
    RenderViewport pdImages(CurrentImage).containingForm

End Sub

'Create a new selection using the settings stored in a pdParamString-compatible string
Public Sub SelectWholeImage()
    
    'Unselect any existing selection
    pdImages(CurrentImage).mainSelection.lockRelease
    pdImages(CurrentImage).selectionActive = False
    
    'Select the rectangular selection tool
    g_PreviousTool = g_CurrentTool
    g_CurrentTool = 0
    FormMain.selectNewTool 0
        
    'Create a new selection at the size of the image
    pdImages(CurrentImage).mainSelection.selectAll
    
    'Lock in this selection
    pdImages(CurrentImage).mainSelection.lockIn pdImages(CurrentImage).containingForm
    pdImages(CurrentImage).selectionActive = True
    
    'Change the selection-related menu items to match
    tInit tSelection, True
    
    'Draw the new selection to the screen
    RenderViewport pdImages(CurrentImage).containingForm

End Sub

'Load a previously saved selection.  Note that this function also handles creation and display of the relevant common dialog.
Public Sub LoadSelectionFromFile()

    'Simple open dialog
    Dim CC As cCommonDialog
        
    Dim sFile As String
    Set CC = New cCommonDialog
    
    Dim cdFilter As String
    cdFilter = PROGRAMNAME & " " & g_Language.TranslateMessage("Selection") & " (." & SELECTION_EXT & ")|*." & SELECTION_EXT & "|"
    cdFilter = cdFilter & g_Language.TranslateMessage("All files") & "|*.*"
    
    Dim cdTitle As String
    cdTitle = g_Language.TranslateMessage("Load a previously saved selection")
    
    If CC.VBGetOpenFileName(sFile, , , , , True, cdFilter, , g_UserPreferences.getSelectionPath, cdTitle, , FormMain.hWnd, 0) Then
        
        Dim tmpSelection As pdSelection
        Set tmpSelection = New pdSelection
        Set tmpSelection.containingPDImage = pdImages(CurrentImage)
        
        If tmpSelection.readSelectionFromFile(sFile, True) Then
            
            'Save the new directory as the default path for future usage
            g_UserPreferences.setSelectionPath sFile
            
            'Activate and redraw the seletion
            Process "Load selection", False, tmpSelection.getSelectionParamString, 2
            
        Else
            pdMsgBox "An error occurred while attempting to load %1.  Please verify that the file is a valid PhotoDemon selection file.", vbOKOnly + vbExclamation + vbApplicationModal, "Selection Error", sFile
        End If
        
        'Release the temporary selection object
        Set tmpSelection.containingPDImage = Nothing
        Set tmpSelection = Nothing
        
    End If
    
End Sub

'Save the current selection to file.  Note that this function also handles creation and display of the relevant common dialog.
Public Sub SaveSelectionToFile()

    'Simple save dialog
    Dim CC As cCommonDialog
        
    Dim sFile As String
    Set CC = New cCommonDialog
    
    Dim cdFilter As String
    cdFilter = PROGRAMNAME & " " & g_Language.TranslateMessage("Selection") & " (." & SELECTION_EXT & ")|*." & SELECTION_EXT
    
    Dim cdTitle As String
    cdTitle = g_Language.TranslateMessage("Save the current selection")
    
    If CC.VBGetSaveFileName(sFile, , True, cdFilter, , g_UserPreferences.getSelectionPath, cdTitle, "." & SELECTION_EXT, FormMain.hWnd, 0) Then
        
        'Save the new directory as the default path for future usage
        g_UserPreferences.setSelectionPath sFile
        
        'Write out the selection file
        If pdImages(CurrentImage).mainSelection.writeSelectionToFile(sFile) Then
            Message "Selection saved."
        Else
            Message "Unknown error occurred.  Selection was not saved.  Please try again."
        End If
        
    End If
    
End Sub
