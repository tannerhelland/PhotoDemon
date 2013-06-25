Attribute VB_Name = "Selection_Handler"
'***************************************************************************
'Selection Interface
'Copyright ©2012-2013 by Tanner Helland
'Created: 21/June/13
'Last updated: 21/June/13
'Last update: initial build
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
