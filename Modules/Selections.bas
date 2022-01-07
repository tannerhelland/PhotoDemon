Attribute VB_Name = "Selections"
'***************************************************************************
'Selection Interface
'Copyright 2013-2022 by Tanner Helland
'Created: 21/June/13
'Last updated: 05/December/21
'Last update: new interactions with the selections toolpanel to auto-drop (or auto-hide) relevant flyout panels
'             while creating or transforming a selection
'
'Selection tools have existed in PhotoDemon for awhile, but this module is the first to support Process varieties of
' selection operations - e.g. internal actions like "Process "Create Selection"".  Selection commands must be passed
' through the Process module so they can be recorded as macros, and as part of the program's Undo/Redo chain.  This
' module provides all selection-related functions that the Process module can call.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

Public Enum PD_SelectionCombine
    pdsm_Replace = 0
    pdsm_Add = 1
    pdsm_Subtract = 2
    pdsm_Intersect = 3
End Enum

#If False Then
    Private Const pdsm_Replace = 0, pdsm_Add = 1, pdsm_Subtract = 2, pdsm_Intersect = 3
#End If

'Create a new selection using the settings stored in a pdSerialize-compatible string
Public Sub CreateNewSelection(ByRef paramString As String)
    
    'Use the passed parameter string to initialize the selection
    PDImages.GetActiveImage.MainSelection.InitFromXML paramString
    PDImages.GetActiveImage.MainSelection.LockIn
    PDImages.GetActiveImage.SetSelectionActive True
    
    'For lasso selections, mark the lasso as closed if the selection is being created anew
    If (PDImages.GetActiveImage.MainSelection.GetSelectionShape() = ss_Lasso) Then PDImages.GetActiveImage.MainSelection.SetLassoClosedState True
    
    'Synchronize all user-facing controls to match
    SelectionUI.SyncTextToCurrentSelection PDImages.GetActiveImageID()
    
End Sub

'Remove the current selection
Public Sub RemoveCurrentSelection(Optional ByVal updateUIToo As Boolean = True)
    
    'Release the selection object and mark it as inactive
    PDImages.GetActiveImage.MainSelection.LockRelease
    PDImages.GetActiveImage.SetSelectionActive False
    
    'Reset any internal selection state trackers
    PDImages.GetActiveImage.MainSelection.EraseCustomTrackers
    
    'Free as many unneeded caches as we can
    PDImages.GetActiveImage.MainSelection.FreeNonEssentialResources
    
    'Synchronize all user-facing controls to match
    If updateUIToo Then SelectionUI.SyncTextToCurrentSelection PDImages.GetActiveImageID()
    
End Sub

'"Select all"
Public Sub SelectWholeImage()
    
    'Unselect any existing selection
    PDImages.GetActiveImage.MainSelection.LockRelease
    PDImages.GetActiveImage.SetSelectionActive False
    
    'Create a new selection at the size of the image
    PDImages.GetActiveImage.MainSelection.SelectAll
    
    'Lock in this selection
    PDImages.GetActiveImage.MainSelection.LockIn
    PDImages.GetActiveImage.SetSelectionActive True
    
    'Synchronize all user-facing controls to match
    SyncTextToCurrentSelection PDImages.GetActiveImageID()
    
End Sub

'Erase the currently selected area (LAYER ONLY!).  Note that this will not modify the current selection in any way;
' only the layer's pixel contents will be affected.
Public Sub EraseSelectedArea(ByVal targetLayerIndex As Long)
    PDImages.GetActiveImage.EraseProcessedSelection targetLayerIndex
    Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), FormMain.MainCanvas(0)
End Sub

'Selections can be initiated several different ways.  To cut down on duplicated code, all new selection instances are referred
' to this function.  Initial X/Y values are required.
Public Sub InitSelectionByPoint(ByVal x As Double, ByVal y As Double)
    
    'Reset any existing selection properties
    PDImages.GetActiveImage.MainSelection.EraseCustomTrackers
    
    'Activate the attached image's primary selection
    PDImages.GetActiveImage.SetSelectionActive True
    PDImages.GetActiveImage.MainSelection.LockRelease
    
    'Reflect all current selection tool settings to the active selection object
    Dim curShape As PD_SelectionShape
    curShape = SelectionUI.GetSelectionShapeFromCurrentTool()
    
    With PDImages.GetActiveImage.MainSelection
        .SetSelectionShape curShape
        If (curShape <> ss_Wand) Then .SetSelectionProperty sp_Area, toolpanel_Selections.cboSelArea(SelectionUI.GetSelectionSubPanelFromCurrentTool).ListIndex Else .SetSelectionProperty sp_Area, sa_Interior
        .SetSelectionProperty sp_Smoothing, toolpanel_Selections.cboSelSmoothing.ListIndex
        .SetSelectionProperty sp_FeatheringRadius, toolpanel_Selections.sltSelectionFeathering.Value
        If (curShape <> ss_Wand) Then .SetSelectionProperty sp_BorderWidth, toolpanel_Selections.sltSelectionBorder(SelectionUI.GetSelectionSubPanelFromCurrentTool).Value
        .SetSelectionProperty sp_RoundedCornerRadius, toolpanel_Selections.sltCornerRounding.Value
        If (curShape = ss_Polygon) Then .SetSelectionProperty sp_PolygonCurvature, toolpanel_Selections.sltPolygonCurvature.Value
        If (curShape = ss_Lasso) Then .SetSelectionProperty sp_SmoothStroke, toolpanel_Selections.sltSmoothStroke.Value
        If (curShape = ss_Wand) Then
            .SetSelectionProperty sp_WandTolerance, toolpanel_Selections.sltWandTolerance.Value
            .SetSelectionProperty sp_WandSampleMerged, toolpanel_Selections.btsWandMerge.ListIndex
            .SetSelectionProperty sp_WandSearchMode, toolpanel_Selections.btsWandArea.ListIndex
            .SetSelectionProperty sp_WandCompareMethod, toolpanel_Selections.cboWandCompare.ListIndex
        End If
    End With
    
    'Set the first two coordinates of this selection to this mouseclick's location
    PDImages.GetActiveImage.MainSelection.SetInitialCoordinates x, y
    SyncTextToCurrentSelection PDImages.GetActiveImageID()
    PDImages.GetActiveImage.MainSelection.RequestNewMask
    
    'Make the selection tools visible
    SetUIGroupState PDUI_Selections, True
    SetUIGroupState PDUI_SelectionTransforms, True
    
    'Ask the selection toolbar to display a flyout with (potentially) useful information.  Note that we
    ' need to pass the current (x, y) coordinates of the mouse - translated into screen coordinate space -
    ' so that the flyout is automatically hidden if the mouse is inside the flyout area.
    Dim screenX As Long, screenY As Long
    Drawing.ConvertImageCoordsToScreenCoords FormMain.MainCanvas(0), PDImages.GetActiveImage, x, y, screenX, screenY, False
    
    'TODO: may need to change selection state bools based on current selection type
    toolpanel_Selections.RequestDefaultFlyout screenX, screenY, True, True, False
    
    'Redraw the screen
    Viewport.Stage3_CompositeCanvas PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    
End Sub

'Are selections currently allowed?  Program states like "no open images" prevent selections from being created,
' and individual functions can use this function to determine that state.  Passing TRUE for the
' "transformableMatters" param will add a check for an existing, transformable-type selection (squares, etc)
' to the evaluation list.  (These have their own unique UI requirements.)
Public Function SelectionsAllowed(ByVal transformableMatters As Boolean) As Boolean
    
    SelectionsAllowed = False
    
    If PDImages.IsImageActive() Then
        If PDImages.GetActiveImage.IsSelectionActive And (Not PDImages.GetActiveImage.MainSelection Is Nothing) Then
            If (Not PDImages.GetActiveImage.MainSelection.GetAutoRefreshSuspend()) Then
                If transformableMatters Then
                    SelectionsAllowed = PDImages.GetActiveImage.MainSelection.IsTransformable
                Else
                    SelectionsAllowed = True
                End If
            End If
        End If
    End If
    
End Function
