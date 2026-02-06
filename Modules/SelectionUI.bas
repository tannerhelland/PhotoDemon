Attribute VB_Name = "SelectionUI"
'***************************************************************************
'Selection Tools: UI
'Copyright 2013-2026 by Tanner Helland
'Created: 21/June/13
'Last updated: 11/March/22
'Last update: fix combine mode accidentally "resetting" after certain actions
'
'This module should only contain UI code related to selection filters (e.g. key and mouse input,
' synchronizing UI elements and internal values, etc).
'
'Selection features have intense UI requirements, owing to their complexity and ubiquity.  This module
' is quite large, but PD also supports many selection features that its competitors do not...
' so hopefully all this extra code is worth it?
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

Public Enum PD_SelectionRenderSetting
    pdsr_Animate
    pdsr_InteriorFillMode
    pdsr_InteriorFillColor
    pdsr_InteriorFillOpacity
    pdsr_ExteriorFillMode
    pdsr_ExteriorFillColor
    pdsr_ExteriorFillOpacity
End Enum

#If False Then
    Private Const pdsr_Animate = 0, pdsr_InteriorFillMode = 0, pdsr_InteriorFillColor = 0, pdsr_InteriorFillOpacity = 0, pdsr_ExteriorFillMode = 0, pdsr_ExteriorFillColor = 0, pdsr_ExteriorFillOpacity = 0
#End If

'Rendering the interior of the current selection with a "fill" style supports three modes:
' always fill, fill only when combining selections (default), never fill
Public Enum PD_SelectionRenderMode
    pdsrm_Always
    pdsrm_Sometimes
    pdsrm_Never
End Enum

#If False Then
    Private Const pdsrm_Always = 0, pdsrm_Sometimes = 0, pdsrm_Never = 0
#End If

'This module caches a number of UI-related selection details.  We cache these here because these
' are tied to program preferences (and not to specific selection instances).  These preferences
' are also highly perf-sensitive, so we do not want to retrieve them on-demand from file or UI elements.
Private m_AnimateOutline As Boolean
Private m_InteriorFillMode As PD_SelectionRenderMode, m_InteriorFillColor As Long, m_InteriorFillOpacity As Single
Private m_ExteriorFillMode As PD_SelectionRenderMode, m_ExteriorFillColor As Long, m_ExteriorFillOpacity As Single

'A double-click event can be used to close the current polygon selection.  Unfortunately, this can
' have the (funny?) side-effect of removing the active selection, because the first click of the
' double-click causes a point to be created, but the second click causes that point to be removed
' and instead the polygon gets closed.  HOWEVER, on the subsequent _MouseUp, the click detector
' notices the _MouseUp potentially occurring *not* over the selection, and it erases the current
' selection accordingly.
'
'To avoid this debacle, we set a flag on the double-click event, and free it on the subsequent
' _MouseUp.
Private m_DblClickOccurred As Boolean

'Similarly, to avoid problematic mouse interactions, we halt processing at the start of _MouseUp,
' then resume processing after any actions triggered by _MouseUp finish.
Private m_IgnoreUserInput As Boolean

'The selection engine can query us about MouseDown and Move state.
' (This changes the way certain selection elements are rendered.)
Private m_MouseDown As Boolean, m_HasMouseMoved As Boolean

'Hotkeys can be used to temporarily trigger a switch to "add" or "subtract" selection mode.
' When these hotkeys are released, we restore the user's original combine mode.
Private m_OriginalCombineMode As PD_SelectionCombine, m_CurrentShiftState As ShiftConstants
Private m_RestoreCombineMode As Boolean

'The shift key (and possibly ctrl in the future, but for a different purpose) can also be used
' to constrain rectangular and elliptical selections to square proportions.  This only works
' when shift is pressed AFTER the mouse has been pressed down, so we track it separately.
Private m_ShiftForConstrain As Boolean

Public Function GetUISetting_Animate() As Boolean
    GetUISetting_Animate = m_AnimateOutline
End Function

Public Function GetUISetting_InteriorFillMode() As PD_SelectionRenderMode
    GetUISetting_InteriorFillMode = m_InteriorFillMode
End Function

Public Function GetUISetting_InteriorFillColor() As Long
    GetUISetting_InteriorFillColor = m_InteriorFillColor
End Function

Public Function GetUISetting_InteriorFillOpacity() As Single
    GetUISetting_InteriorFillOpacity = m_InteriorFillOpacity
End Function

Public Function GetUISetting_ExteriorFillMode() As PD_SelectionRenderMode
    GetUISetting_ExteriorFillMode = m_ExteriorFillMode
End Function

Public Function GetUISetting_ExteriorFillColor() As Long
    GetUISetting_ExteriorFillColor = m_ExteriorFillColor
End Function

Public Function GetUISetting_ExteriorFillOpacity() As Single
    GetUISetting_ExteriorFillOpacity = m_ExteriorFillOpacity
End Function

'The selection engine integrates closely with tool selection (as it needs to know what kind of selection is being
' created/edited at any given time).  This function is called whenever the selection engine needs to correlate the
' current tool with a selection shape.  This allows us to easily switch between a rectangle and circle selection,
' for example, without forcing the user to recreate the selection from scratch.
Public Function GetSelectionShapeFromCurrentTool() As PD_SelectionShape

    Select Case g_CurrentTool
    
        Case SELECT_RECT
            GetSelectionShapeFromCurrentTool = ss_Rectangle
            
        Case SELECT_CIRC
            GetSelectionShapeFromCurrentTool = ss_Circle
        
        Case SELECT_POLYGON
            GetSelectionShapeFromCurrentTool = ss_Polygon
            
        Case SELECT_LASSO
            GetSelectionShapeFromCurrentTool = ss_Lasso
            
        Case SELECT_WAND
            GetSelectionShapeFromCurrentTool = ss_Wand
            
        Case Else
            GetSelectionShapeFromCurrentTool = -1
    
    End Select
    
End Function

'The inverse of "getSelectionShapeFromCurrentTool", above
Public Function GetRelevantToolFromSelectShape() As PDTools

    If PDImages.IsImageActive() Then

        If (Not PDImages.GetActiveImage.MainSelection Is Nothing) Then

            Select Case PDImages.GetActiveImage.MainSelection.GetSelectionShape
            
                Case ss_Rectangle
                    GetRelevantToolFromSelectShape = SELECT_RECT
                    
                Case ss_Circle
                    GetRelevantToolFromSelectShape = SELECT_CIRC
                
                Case ss_Polygon
                    GetRelevantToolFromSelectShape = SELECT_POLYGON
                    
                Case ss_Lasso
                    GetRelevantToolFromSelectShape = SELECT_LASSO
                    
                Case ss_Wand
                    GetRelevantToolFromSelectShape = SELECT_WAND
                
                Case Else
                    GetRelevantToolFromSelectShape = -1
            
            End Select
            
        Else
            GetRelevantToolFromSelectShape = -1
        End If
            
    Else
        GetRelevantToolFromSelectShape = -1
    End If

End Function

'All selection tools share the same main panel on the options toolbox, but they have different subpanels that contain their
' specific parameters.  Use this function to correlate the two.
Public Function GetSelectionSubPanelFromCurrentTool() As Long

    Select Case g_CurrentTool
    
        Case SELECT_RECT
            GetSelectionSubPanelFromCurrentTool = 0
            
        Case SELECT_CIRC
            GetSelectionSubPanelFromCurrentTool = 1
        
        Case SELECT_POLYGON
            GetSelectionSubPanelFromCurrentTool = 2
            
        Case SELECT_LASSO
            GetSelectionSubPanelFromCurrentTool = 3
            
        Case SELECT_WAND
            GetSelectionSubPanelFromCurrentTool = 4
        
        Case Else
            GetSelectionSubPanelFromCurrentTool = -1
    
    End Select
    
End Function

Public Function GetSelectionSubPanelFromSelectionShape(ByRef srcImage As pdImage) As Long

    Select Case srcImage.MainSelection.GetSelectionShape
    
        Case ss_Rectangle
            GetSelectionSubPanelFromSelectionShape = 0
            
        Case ss_Circle
            GetSelectionSubPanelFromSelectionShape = 1
        
        Case ss_Polygon
            GetSelectionSubPanelFromSelectionShape = 2
            
        Case ss_Lasso
            GetSelectionSubPanelFromSelectionShape = 3
            
        Case ss_Wand
            GetSelectionSubPanelFromSelectionShape = 4
        
        Case Else
            GetSelectionSubPanelFromSelectionShape = -1
    
    End Select
    
End Function

Public Function GetSelectionUI_ShiftState() As Boolean
    GetSelectionUI_ShiftState = (m_CurrentShiftState <> 0)
End Function

'Call at program startup.
' At present, all this function does is cache the current user preferences for selection rendering settings.
' This ensures the settings are up-to-date, even if the user does not activate a specific selection tool.
' (Why does this matter? Selections can be loaded directly from file, without ever invoking a tool, so we
' need to ensure rendering settings are up-to-date when the program starts.)
Public Sub InitializeSelectionRendering()

    If UserPrefs.IsReady Then
        
        'Animate marching ants
        m_AnimateOutline = UserPrefs.GetPref_Boolean("Tools", "SelectionAnimateOutline", True)
        
        'Interior and exterior fill settings
        m_InteriorFillMode = UserPrefs.GetPref_Long("Tools", "SelectionInteriorFillMode", pdsrm_Sometimes)
        m_InteriorFillColor = Colors.GetRGBLongFromHex(UserPrefs.GetPref_String("Tools", "SelectionInteriorFillColor", "#6EE6FF"))
        m_InteriorFillOpacity = UserPrefs.GetPref_Float("Tools", "SelectionInteriorFillOpacity", 50!)
        
        m_ExteriorFillMode = UserPrefs.GetPref_Long("Tools", "SelectionExteriorFillMode", pdsrm_Never)
        m_ExteriorFillColor = Colors.GetRGBLongFromHex(UserPrefs.GetPref_String("Tools", "SelectionExteriorFillColor", "#FF3C50"))
        m_ExteriorFillOpacity = UserPrefs.GetPref_Float("Tools", "SelectionExteriorFillOpacity", 50!)
        
    End If

End Sub

'Whenever a selection render setting changes (like switching between outline and highlight mode), you must call this function
' so that we can cache the new render settings.
Public Sub NotifySelectionRenderChange(ByVal settingType As PD_SelectionRenderSetting, ByVal newValue As Variant)
    
    Select Case settingType
        Case pdsr_Animate
            m_AnimateOutline = newValue
            If Selections.SelectionsAllowed(False) Then PDImages.GetActiveImage.MainSelection.NotifyAnimationsAllowed newValue
            
            'Selection rendering settings are cached in PD's main preferences file.  This allows outside functions to access
            ' them correctly, even if selection tools have not been loaded this session.  (This can happen if the user runs
            ' the program, loads an image, then loads a selection directly from file, without invoking a specific tool.)
            If UserPrefs.IsReady Then UserPrefs.WritePreference "Tools", "SelectionAnimateOutline", Trim$(Str$(m_AnimateOutline))
            
        Case pdsr_InteriorFillMode
            m_InteriorFillMode = newValue
            If UserPrefs.IsReady Then UserPrefs.WritePreference "Tools", "SelectionInteriorFillMode", m_InteriorFillMode
            
        Case pdsr_InteriorFillColor
            m_InteriorFillColor = newValue
            If UserPrefs.IsReady Then UserPrefs.WritePreference "Tools", "SelectionInteriorFillColor", Colors.GetHexStringFromRGB(m_InteriorFillColor)
            
        Case pdsr_InteriorFillOpacity
            m_InteriorFillOpacity = newValue
            If UserPrefs.IsReady Then UserPrefs.WritePreference "Tools", "SelectionInteriorFillOpacity", m_InteriorFillOpacity
            
        Case pdsr_ExteriorFillMode
            m_ExteriorFillMode = newValue
            If UserPrefs.IsReady Then UserPrefs.WritePreference "Tools", "SelectionExteriorFillMode", m_ExteriorFillMode
            
        Case pdsr_ExteriorFillColor
            m_ExteriorFillColor = newValue
            If UserPrefs.IsReady Then UserPrefs.WritePreference "Tools", "SelectionExteriorFillColor", Colors.GetHexStringFromRGB(m_ExteriorFillColor)
            
        Case pdsr_ExteriorFillOpacity
            m_ExteriorFillOpacity = newValue
            If UserPrefs.IsReady Then UserPrefs.WritePreference "Tools", "SelectionExteriorFillOpacity", m_ExteriorFillOpacity
            
    End Select
    
End Sub

Public Function HasMouseMoved() As Boolean
    HasMouseMoved = m_HasMouseMoved
End Function

'Given an (x, y) pair in IMAGE coordinate space (not screen or canvas space), return a constant if the point is a valid
' "point of interest" for the active selection.  Standard UI mouse distances are allowed (meaning zoom is factored into the
' algorithm).
'
'The result of this function is typically passed to something like pdSelection.SetActiveSelectionPOI(), which will cache
' the point of interest and use it to interpret subsequent mouse events (e.g. click-dragging a selection to a new position).
'
'Note that only certain POIs are hard-coded.  Some selections (e.g. polygons) can return other values outside the enum,
' typically indices into an internal selection point array.
'
'This sub will return a constant correlating to the nearest selection point.  See the relevant enum for details.
Public Function IsCoordSelectionPOI(ByVal imgX As Double, ByVal imgY As Double, ByRef srcImage As pdImage) As PD_PointOfInterest
    
    IsCoordSelectionPOI = poi_Undefined
    
    'If the current selection is...
    ' 1) raster-type, or...
    ' 2) inactive...
    '...disallow POIs entirely.  (These types of selections do not support on-canvas interactions.)
    If (srcImage.MainSelection.GetSelectionShape = ss_Raster) Or (Not srcImage.IsSelectionActive) Then Exit Function
    
    'Similarly, POIs are only enabled if the current selection tool matches the current selection shape.
    ' (If a new selection shape has been selected, the user is definitely not modifying the existing selection.)
    If (g_CurrentTool <> SelectionUI.GetRelevantToolFromSelectShape()) Then IsCoordSelectionPOI = poi_Undefined
    
    'We're now going to compare the passed coordinate against a hard-coded list of "points of interest."  These POIs
    ' differ by selection type, as different selections allow for different levels of interaction.  (For example, a polygon
    ' selection behaves differently when a point is dragged, vs a rectangular selection.)
    
    'Regardless of selection type, start by establishing boundaries for the current selection.
    'Calculate points of interest for the current selection.  Individual selection types define what is considered a POI,
    ' but in most cases, corners or interior clicks tend to allow some kind of user interaction.
    Dim tmpRectF As RectF
    If (srcImage.MainSelection.GetSelectionShape = ss_Rectangle) Or (srcImage.MainSelection.GetSelectionShape = ss_Circle) Then
        tmpRectF = srcImage.MainSelection.GetCornersLockedRect()
    Else
        tmpRectF = srcImage.MainSelection.GetCompositeBoundaryRect()
    End If
    
    'Adjust the mouseAccuracy value based on the current zoom value
    Dim mouseAccuracy As Double
    mouseAccuracy = Drawing.ConvertCanvasSizeToImageSize(Interface.GetStandardInteractionDistance(), srcImage)
        
    'Find the smallest distance for this mouse position
    Dim minDistance As Double
    minDistance = mouseAccuracy
    
    Dim closestPoint As Long
    closestPoint = poi_Undefined
    
    'Some selection types (lasso, polygon) must use a more complicated region for hit-testing.  GDI+ will be used for this.
    Dim complexRegion As pd2DRegion
    
    'Other selection types will use a generic list of points (like the corners of the current selection)
    Dim poiListFloat() As PointFloat
    
    'If we made it here, this mouse location is worth evaluating.  How we evaluate it depends on the shape of the current selection.
    Select Case srcImage.MainSelection.GetSelectionShape
    
        'Rectangular and elliptical selections have identical POIs: the corners, edges, and interior of the selection
        Case ss_Rectangle, ss_Circle
    
            'Corners get preference, so check them first.
            ReDim poiListFloat(0 To 3) As PointFloat
            
            With tmpRectF
                poiListFloat(0).x = .Left
                poiListFloat(0).y = .Top
                poiListFloat(1).x = .Left + .Width
                poiListFloat(1).y = .Top
                poiListFloat(2).x = .Left + .Width
                poiListFloat(2).y = .Top + .Height
                poiListFloat(3).x = .Left
                poiListFloat(3).y = .Top + .Height
            End With
            
            'Used the generalized point comparison function to see if one of the points matches
            closestPoint = FindClosestPointInFloatArray(imgX, imgY, minDistance, poiListFloat)
            
            'Did one of the corner points match?  If so, map it to a valid constant and return.
            If (closestPoint <> poi_Undefined) Then
                
                If (closestPoint = 0) Then
                    IsCoordSelectionPOI = poi_CornerNW
                ElseIf (closestPoint = 1) Then
                    IsCoordSelectionPOI = poi_CornerNE
                ElseIf (closestPoint = 2) Then
                    IsCoordSelectionPOI = poi_CornerSE
                ElseIf (closestPoint = 3) Then
                    IsCoordSelectionPOI = poi_CornerSW
                
                'Failsafe only
                Else
                    IsCoordSelectionPOI = poi_Undefined
                End If
                
            Else
        
                'If we're at this line of code, a closest corner was not found.  Check edges next.
                ' (Unfortunately, we don't yet have a generalized function for edge checking, so this must be done manually.)
                '
                'Note that edge checks are a little weird currently, because we check one-dimensional distance between each
                ' side, and if that's a hit, we see if the point also lies between the bounds in the *other* direction.
                ' This allows the user to use the entire selection side to perform a stretch.
                Dim nDist As Double, eDist As Double, sDist As Double, wDist As Double
                
                With tmpRectF
                    nDist = DistanceOneDimension(imgY, .Top)
                    eDist = DistanceOneDimension(imgX, .Left + .Width)
                    sDist = DistanceOneDimension(imgY, .Top + .Height)
                    wDist = DistanceOneDimension(imgX, .Left)
                    
                    If (nDist <= minDistance) Then
                        If (imgX > (.Left - minDistance)) And (imgX < (.Left + .Width + minDistance)) Then
                            minDistance = nDist
                            closestPoint = poi_EdgeN
                        End If
                    End If
                    
                    If (eDist <= minDistance) Then
                        If (imgY > (.Top - minDistance)) And (imgY < (.Top + .Height + minDistance)) Then
                            minDistance = eDist
                            closestPoint = poi_EdgeE
                        End If
                    End If
                    
                    If (sDist <= minDistance) Then
                        If (imgX > (.Left - minDistance)) And (imgX < (.Left + .Width + minDistance)) Then
                            minDistance = sDist
                            closestPoint = poi_EdgeS
                        End If
                    End If
                    
                    If (wDist <= minDistance) Then
                        If (imgY > (.Top - minDistance)) And (imgY < (.Top + .Height + minDistance)) Then
                            minDistance = wDist
                            closestPoint = poi_EdgeW
                        End If
                    End If
                
                End With
                
                'Was a close point found? If yes, then return that value.
                If (closestPoint <> poi_Undefined) Then
                    IsCoordSelectionPOI = closestPoint
                Else
            
                    'If we're at this line of code, a closest edge was not found. Perform one final check to ensure that the mouse is within the
                    ' image's boundaries, and if it is, return the "move selection" ID, then exit.
                    If PDMath.IsPointInRectF(imgX, imgY, tmpRectF) Then
                        IsCoordSelectionPOI = poi_Interior
                    Else
                        IsCoordSelectionPOI = poi_Undefined
                    End If
                    
                End If
                
            End If
            
        Case ss_Polygon
            
            If (PDImages.GetActiveImage.MainSelection.GetNumOfPolygonPoints() > 0) Then
                
                'First, we want to check all polygon points for a hit.
                PDImages.GetActiveImage.MainSelection.GetPolygonPoints poiListFloat()
                
                'Used the generalized point comparison function to see if one of the points matches
                closestPoint = FindClosestPointInFloatArray(imgX, imgY, minDistance, poiListFloat)
                
                'Was a close point found? If yes, then return that value
                If (closestPoint <> poi_Undefined) Then
                    IsCoordSelectionPOI = closestPoint
                    
                'If no polygon point was a hit, our final check is to see if the mouse lies within the polygon itself.
                ' This will trigger a move transformation.
                Else
                    
                    'Use a region object for hit-detection
                    Set complexRegion = PDImages.GetActiveImage.MainSelection.GetSelectionAsRegion()
                    If (Not complexRegion Is Nothing) Then
                        If complexRegion.IsPointInRegion(imgX, imgY) Then IsCoordSelectionPOI = poi_Interior Else IsCoordSelectionPOI = poi_Undefined
                    Else
                        IsCoordSelectionPOI = poi_Undefined
                    End If
                    
                End If
                
            Else
                IsCoordSelectionPOI = poi_Undefined
            End If
            
        Case ss_Lasso
        
            'Use a region object for hit-detection
            Set complexRegion = PDImages.GetActiveImage.MainSelection.GetSelectionAsRegion()
            If (Not complexRegion Is Nothing) Then
                If complexRegion.IsPointInRegion(imgX, imgY) Then IsCoordSelectionPOI = poi_Interior Else IsCoordSelectionPOI = poi_Undefined
            Else
                IsCoordSelectionPOI = poi_Undefined
            End If
                
        Case ss_Wand
            
            'Wand selections do actually support a single point of interest - the wand's "clicked" location
            srcImage.MainSelection.GetCurrentPOIList poiListFloat
            
            'Used the generalized point comparison function to see if one of the points matches
            IsCoordSelectionPOI = FindClosestPointInFloatArray(imgX, imgY, minDistance, poiListFloat)
            
        Case Else
            IsCoordSelectionPOI = poi_Undefined
            Exit Function
            
    End Select

End Function

Public Function IsMouseDown() As Boolean
    IsMouseDown = m_MouseDown
End Function

'Keypresses on a source canvas are passed here.  The caller doesn't need pass anything except relevant keycodes, and a reference
' to itself (so we can relay canvas modifications).
Public Sub NotifySelectionKeyDown(ByRef srcCanvas As pdCanvas, ByVal Shift As ShiftConstants, ByVal vkCode As Long, ByRef markEventHandled As Boolean)
    
    'Handle arrow keys first
    If (vkCode = VK_UP) Or (vkCode = VK_DOWN) Or (vkCode = VK_LEFT) Or (vkCode = VK_RIGHT) Then

        'If a selection is active, nudge it using the arrow keys
        If (PDImages.GetActiveImage.IsSelectionActive And (PDImages.GetActiveImage.MainSelection.GetSelectionShape <> ss_Raster)) Then
            
            Dim canvasUpdateRequired As Boolean
            canvasUpdateRequired = False
            
            'Suspend automatic redraws until all arrow keys have been processed
            srcCanvas.SetRedrawSuspension True
            
            'If scrollbars are visible, nudge the canvas in the direction of the arrows.
            If srcCanvas.GetScrollVisibility(pdo_Vertical) Then
                If (vkCode = VK_UP) Or (vkCode = VK_DOWN) Then canvasUpdateRequired = True
                If (vkCode = VK_UP) Then srcCanvas.SetScrollValue pdo_Vertical, srcCanvas.GetScrollValue(pdo_Vertical) - 1
                If (vkCode = VK_DOWN) Then srcCanvas.SetScrollValue pdo_Vertical, srcCanvas.GetScrollValue(pdo_Vertical) + 1
            End If
            
            If srcCanvas.GetScrollVisibility(pdo_Horizontal) Then
                If (vkCode = VK_LEFT) Or (vkCode = VK_RIGHT) Then canvasUpdateRequired = True
                If (vkCode = VK_LEFT) Then srcCanvas.SetScrollValue pdo_Horizontal, srcCanvas.GetScrollValue(pdo_Horizontal) - 1
                If (vkCode = VK_RIGHT) Then srcCanvas.SetScrollValue pdo_Horizontal, srcCanvas.GetScrollValue(pdo_Horizontal) + 1
            End If
            
            'Re-enable automatic redraws
            srcCanvas.SetRedrawSuspension False
            
            'Redraw the viewport if necessary
            If canvasUpdateRequired Then
                markEventHandled = True
                Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), srcCanvas
            End If
            
        End If
    
    'Handle non-arrow keys here.  (Note: most non-arrow keys are not meant to work with key-repeating,
    ' so they are handled in the KeyUp event instead.)
    Else
        
        'If the mouse is *not* down, the user can use Shift and Alt keys to change combine mode.
        If (Not m_MouseDown) Then
        
            'If this is the first keypress during this round, save the user's current combine mode
            If (m_CurrentShiftState = 0) Then m_OriginalCombineMode = toolpanel_Selections.btsCombine.ListIndex
            
            'Add this state to the running tracker
            If (vkCode = VK_SHIFT) Then m_CurrentShiftState = m_CurrentShiftState Or vbShiftMask
            If (vkCode = VK_CONTROL) Then m_CurrentShiftState = m_CurrentShiftState Or vbCtrlMask
            If (vkCode = VK_ALT) Then m_CurrentShiftState = m_CurrentShiftState Or vbAltMask
            
            'The actual synchronizing between hotkey and UI/selection object is handled elsewhere
            SyncCombineModeToHotkeys
            
        Else
            If (vkCode = VK_SHIFT) And ((m_CurrentShiftState And vbShiftMask) = 0) Then
                m_ShiftForConstrain = True
            End If
        End If
        
    End If
    
End Sub

Public Sub NotifySelectionKeyUp(ByRef srcCanvas As pdCanvas, ByVal Shift As ShiftConstants, ByVal vkCode As Long, ByRef markEventHandled As Boolean)
    
    'Ctrl/Alt/Shift modifiers can change combine mode
    If (m_CurrentShiftState <> 0) Then
        
        If (vkCode = VK_SHIFT) Or (vkCode = VK_CONTROL) Or (vkCode = VK_ALT) Then
            
            'Update all flags
            If (vkCode = VK_SHIFT) Then m_CurrentShiftState = m_CurrentShiftState And (Not vbShiftMask)
            If (vkCode = VK_CONTROL) Then m_CurrentShiftState = m_CurrentShiftState And (Not vbCtrlMask)
            If (vkCode = VK_ALT) Then m_CurrentShiftState = m_CurrentShiftState And (Not vbAltMask)
            
            'If all modifier keys have been released, restore the user's original combine mode
            If (m_CurrentShiftState = 0) Then
                
                'If the mouse is *still* down, don't make any changes now - instead, set a flag and
                ' we'll restore the preferred combine mode in _MouseUp
                If m_MouseDown Then
                    m_RestoreCombineMode = True
                Else
                    m_RestoreCombineMode = False
                    toolpanel_Selections.btsCombine.ListIndex = m_OriginalCombineMode
                End If
                
            'If at least one modifier is still down, and the mouse is *not* down, switch to a new combine mode
            Else
                If (Not m_MouseDown) Then SyncCombineModeToHotkeys
            End If
            
        End If
    
    End If
    
    'Shift for constrain (works only during _MouseDown; see top of module for additional comments)
    If (vkCode = VK_SHIFT) Then m_ShiftForConstrain = False
    
    'Delete key: if a selection is active, erase the selected area
    If (vkCode = VK_DELETE) And PDImages.GetActiveImage.IsSelectionActive Then
        markEventHandled = True
        Process "Erase selected area", False, BuildParamList("targetlayer", PDImages.GetActiveImage.GetActiveLayerIndex), UNDO_Layer
    End If
    
    'Escape key: if a selection is active, clear it
    If (vkCode = VK_ESCAPE) And PDImages.GetActiveImage.IsSelectionActive Then
        markEventHandled = True
        Process "Remove selection", , , UNDO_Selection
    End If
    
    'Enter/return keys: for polygon selections, this will close the current selection
    If ((vkCode = VK_RETURN) Or (vkCode = VK_SPACE)) And (g_CurrentTool = SELECT_POLYGON) Then
        
        'A selection must be in-progress
        If PDImages.GetActiveImage.IsSelectionActive Then
        
            'The selection must *not* be closed yet, but there must be enough points to successfully close it
            If (Not PDImages.GetActiveImage.MainSelection.GetPolygonClosedState) And (PDImages.GetActiveImage.MainSelection.GetNumOfPolygonPoints > 2) Then
            
                'Close the selection
                PDImages.GetActiveImage.MainSelection.SetPolygonClosedState True
                PDImages.GetActiveImage.MainSelection.SetActiveSelectionPOI 0
                
                'Fully process the selection (important when recording macros!)
                Process "Create selection", , PDImages.GetActiveImage.MainSelection.GetSelectionAsXML, UNDO_Selection, g_CurrentTool
                
                'Redraw the viewport
                Viewport.Stage3_CompositeCanvas PDImages.GetActiveImage(), srcCanvas
                
            End If
        
        End If
        
    End If
    
    'Backspace key: for lasso and polygon selections, retreat back one or more coordinates, giving the user a chance to
    ' correct any potential mistakes.
    If (vkCode = VK_BACK) And ((g_CurrentTool = SELECT_LASSO) Or (g_CurrentTool = SELECT_POLYGON)) And PDImages.GetActiveImage.IsSelectionActive Then
        
        If (Not PDImages.GetActiveImage.MainSelection.IsLockedIn) Then
            
            markEventHandled = True
            
            'Polygons: do not allow point removal if the polygon has already been successfully closed.
            If (g_CurrentTool = SELECT_POLYGON) Then
                If (Not PDImages.GetActiveImage.MainSelection.GetPolygonClosedState) Then PDImages.GetActiveImage.MainSelection.RemoveLastPolygonPoint
            
            'Lassos: do not allow point removal if the lasso has already been successfully closed.
            Else
            
                If (Not PDImages.GetActiveImage.MainSelection.GetLassoClosedState) Then
            
                    'Ask the selection object to retreat its position
                    Dim newImageX As Double, newImageY As Double
                    PDImages.GetActiveImage.MainSelection.RetreatLassoPosition newImageX, newImageY
                    
                    'The returned coordinates will be in image coordinates.  Convert them to viewport coordinates.
                    Dim newCanvasX As Double, newCanvasY As Double
                    Drawing.ConvertImageCoordsToCanvasCoords srcCanvas, PDImages.GetActiveImage(), newImageX, newImageY, newCanvasX, newCanvasY
                    
                    'Finally, convert the canvas coordinates to screen coordinates, and move the cursor accordingly
                    srcCanvas.SetCursorToCanvasPosition newCanvasX, newCanvasY
                    
                End If
                
            End If
            
            'Redraw the screen to reflect this new change.
            Viewport.Stage3_CompositeCanvas PDImages.GetActiveImage(), srcCanvas
            
        End If
        
    End If
    
End Sub

Public Sub NotifySelectionMouseDown(ByRef srcCanvas As pdCanvas, ByVal imgX As Single, ByVal imgY As Single)
    
    If m_IgnoreUserInput Then Exit Sub
    
    m_MouseDown = True
    m_HasMouseMoved = False
    
    'Check to see if a selection is already active.  If it is, see if the user is clicking on a POI
    ' (and initiating a transform) or clicking somewhere else (initiating a new selection).
    If PDImages.GetActiveImage.IsSelectionActive Then
        
        'Check the mouse coordinates of this click.
        Dim sCheck As PD_PointOfInterest
        sCheck = SelectionUI.IsCoordSelectionPOI(imgX, imgY, PDImages.GetActiveImage())
        
        'TODO: potentially deal with raster selections here?  Right now PD doesn't allow any transforms
        ' on raster selections, but hypothetically the user could be allowed to click-drag to move them...
        ' I'll see if this is feasible in the future.
        
        'Polygon selections require special handling, because they don't operate on the same
        ' "mouse up = finished selection" assumption.  They are marked as complete under
        ' special circumstances (when the user re-clicks the first point or double-clicks).
        ' Any clicks prior to this are treated as an instruction to add a new point to the shape.
        If (g_CurrentTool = SELECT_POLYGON) Then
            
            'If a point of interest was clicked, initiate a transform event (to allow modification
            ' of the *already existing* selection).
            If (sCheck <> poi_Undefined) And (PDImages.GetActiveImage.MainSelection.GetSelectionShape <> ss_Raster) And PDImages.GetActiveImage.MainSelection.GetPolygonClosedState() Then
                PDImages.GetActiveImage.MainSelection.SetActiveSelectionPOI sCheck
                PDImages.GetActiveImage.MainSelection.SetInitialTransformCoordinates imgX, imgY
            
            Else
                
                'First, see if the current polygon is "locked in" (i.e. finished).
                ' If it is, treat this as starting a new selection.
                If PDImages.GetActiveImage.MainSelection.IsLockedIn Then
                    
                    Selections.NotifyNewSelectionStarting
                    Selections.InitSelectionByPoint imgX, imgY
                    
                    'Start transformation mode, using the index of the new point as the transform ID.
                    ' (This allows the user to click-drag this initial point.)
                    PDImages.GetActiveImage.MainSelection.SetInitialTransformCoordinates imgX, imgY
                    PDImages.GetActiveImage.MainSelection.SetActiveSelectionPOI 0
                    PDImages.GetActiveImage.MainSelection.OverrideTransformMode True
                    
                    'Redraw the screen
                    Viewport.Stage3_CompositeCanvas PDImages.GetActiveImage(), srcCanvas
                    
                'If the polygon is *not* locked in, the user is still constructing it.
                Else
                    
                    'If the user clicked on the initial polygon point, attempt to close the polygon
                    If (sCheck = 0) And (PDImages.GetActiveImage.MainSelection.GetNumOfPolygonPoints > 2) Then
                        
                        'Set appropriate closed flags, and activate the first point as a transform target
                        PDImages.GetActiveImage.MainSelection.SetPolygonClosedState True
                        PDImages.GetActiveImage.MainSelection.SetActiveSelectionPOI 0
                        
                    'The user did not click the initial polygon point, meaning we should add this coordinate as a new polygon point.
                    Else
                        
                        'Remove the current transformation mode (if any)
                        PDImages.GetActiveImage.MainSelection.SetActiveSelectionPOI poi_Undefined
                        PDImages.GetActiveImage.MainSelection.OverrideTransformMode False
                        
                        'Add the new point
                        If (PDImages.GetActiveImage.MainSelection.GetNumOfPolygonPoints = 0) Then
                            Selections.NotifyNewSelectionStarting
                            Selections.InitSelectionByPoint imgX, imgY
                        Else
                            
                            If (sCheck = poi_Undefined) Or (sCheck = poi_Interior) Then
                                PDImages.GetActiveImage.MainSelection.SetAdditionalCoordinates imgX, imgY
                                PDImages.GetActiveImage.MainSelection.SetActiveSelectionPOI PDImages.GetActiveImage.MainSelection.GetNumOfPolygonPoints - 1
                            Else
                                PDImages.GetActiveImage.MainSelection.SetActiveSelectionPOI sCheck
                            End If
                            
                        End If
                        
                        'Reinstate transformation mode, using the index of the new point as the transform ID
                        PDImages.GetActiveImage.MainSelection.SetInitialTransformCoordinates imgX, imgY
                        PDImages.GetActiveImage.MainSelection.OverrideTransformMode True
                        
                        'Redraw the screen
                        Viewport.Stage3_CompositeCanvas PDImages.GetActiveImage(), srcCanvas
                        
                    End If
                
                End If
                
            End If
            
        'All other selection types are *much* simpler
        Else
            
            'If a point of interest was clicked, initiate a transform event (to allow modification
            ' of the *already existing* selection).
            If (sCheck <> poi_Undefined) And (PDImages.GetActiveImage.MainSelection.GetSelectionShape <> ss_Raster) Then
                PDImages.GetActiveImage.MainSelection.SetActiveSelectionPOI sCheck
                PDImages.GetActiveImage.MainSelection.SetInitialTransformCoordinates imgX, imgY
                
            'If a point of interest was *not* clicked, start a new selection at the clicked location
            Else
                Selections.NotifyNewSelectionStarting
                Selections.InitSelectionByPoint imgX, imgY
            End If
            
        End If
        
    'If a selection is not already active, start a new one.
    Else
        
        Selections.NotifyNewSelectionStarting
        Selections.InitSelectionByPoint imgX, imgY
        
        'Polygon selections require special handling.  After creating the initial point,
        ' we want to immediately initiate "transform mode" (which allows the user to drag
        ' the mouse to move the newly created point).
        If (g_CurrentTool = SELECT_POLYGON) Then
            PDImages.GetActiveImage.MainSelection.SetActiveSelectionPOI PDImages.GetActiveImage.MainSelection.GetNumOfPolygonPoints - 1
            PDImages.GetActiveImage.MainSelection.OverrideTransformMode True
        End If
        
    End If
    
End Sub

'The only selection tool that responds to double-click events is the polygon selection tool.
' Photoshop convention (mirrored by GIMP, Krita) is to close the polygon on a double-click.
Public Sub NotifySelectionMouseDblClick(ByRef srcCanvas As pdCanvas, ByVal imgX As Single, ByVal imgY As Single)
    
    'Polygon selections only
    If (g_CurrentTool = SELECT_POLYGON) Then
    
        'A selection must be in-progress
        If PDImages.GetActiveImage.IsSelectionActive Then
        
            'The selection must *not* be closed yet
            If (Not PDImages.GetActiveImage.MainSelection.GetPolygonClosedState) And (PDImages.GetActiveImage.MainSelection.GetNumOfPolygonPoints > 2) Then
                
                'Set a flag to note that a double-click just occurred.  (See notes at the
                ' top of this module for details.)
                m_DblClickOccurred = True
                
                'Remove the last point (the point created by the first click of this
                ' double-click event), but *only* if there are enough valid points
                ' to create a polygon selection without it!
                If (PDImages.GetActiveImage.MainSelection.GetNumOfPolygonPoints > 3) Then PDImages.GetActiveImage.MainSelection.RemoveLastPolygonPoint
                
                'Close the selection and make the first point the active one
                PDImages.GetActiveImage.MainSelection.SetPolygonClosedState True
                PDImages.GetActiveImage.MainSelection.SetActiveSelectionPOI 0
                
                'Fully process the selection (important when recording macros!)
                Process "Create selection", , PDImages.GetActiveImage.MainSelection.GetSelectionAsXML, UNDO_Selection, g_CurrentTool
                
                'Redraw the viewport
                Viewport.Stage3_CompositeCanvas PDImages.GetActiveImage(), srcCanvas
                
            End If
        
        End If
    
    End If

End Sub

Public Sub NotifySelectionMouseLeave(ByRef srcCanvas As pdCanvas)
    
    'Ensure input behavior is normalized
    m_IgnoreUserInput = False
    
    'When the polygon selection tool is being used, redraw the canvas when the mouse leaves
    If (g_CurrentTool = SELECT_POLYGON) Then Viewport.Stage3_CompositeCanvas PDImages.GetActiveImage(), srcCanvas

End Sub

Public Sub NotifySelectionMouseMove(ByRef srcCanvas As pdCanvas, ByVal lmbState As Boolean, ByVal Shift As ShiftConstants, ByVal imgX As Single, ByVal imgY As Single, ByVal numOfCanvasMoveEvents As Long)
        
    If m_IgnoreUserInput Then Exit Sub
    m_HasMouseMoved = True
    
    'Handling varies based on the current mouse state, obviously.
    If m_MouseDown Then
        
        'Basic selection tools
        Select Case g_CurrentTool
            
            Case SELECT_RECT, SELECT_CIRC, SELECT_POLYGON
                
                'First, check to see if a selection is both active and transformable.
                If PDImages.GetActiveImage.IsSelectionActive And (PDImages.GetActiveImage.MainSelection.GetSelectionShape <> ss_Raster) Then
                    
                    'If the SHIFT key is down, notify the selection engine that a square shape is requested
                    PDImages.GetActiveImage.MainSelection.RequestSquare m_ShiftForConstrain
                    
                    'Pass new points to the active selection
                    PDImages.GetActiveImage.MainSelection.SetAdditionalCoordinates imgX, imgY
                    SelectionUI.SyncTextToCurrentSelection PDImages.GetActiveImageID()
                    
                End If
                
                'Force a redraw of the viewport
                If (numOfCanvasMoveEvents > 1) Then Viewport.Stage3_CompositeCanvas PDImages.GetActiveImage(), srcCanvas
            
            'Lasso selections are handled specially, because mouse move events control the drawing of the lasso
            Case SELECT_LASSO
            
                'First, check to see if a selection is active
                If PDImages.GetActiveImage.IsSelectionActive Then
                    
                    'Pass new points to the active selection
                    PDImages.GetActiveImage.MainSelection.SetAdditionalCoordinates imgX, imgY
                                        
                End If
                
                'To spare the debug logger from receiving too many events, forcibly prevent logging of this message
                ' while in debug mode.
                If UserPrefs.GenerateDebugLogs Then
                    Message "Release the mouse button to complete the lasso selection", "DONOTLOG"
                Else
                    Message "Release the mouse button to complete the lasso selection"
                End If
                
                'Force a redraw of the viewport
                If (numOfCanvasMoveEvents > 1) Then Viewport.Stage3_CompositeCanvas PDImages.GetActiveImage(), srcCanvas
            
            'Wand selections are easier than other selection types, because they don't support any special transforms
            Case SELECT_WAND
                If PDImages.GetActiveImage.IsSelectionActive Then
                    PDImages.GetActiveImage.MainSelection.SetAdditionalCoordinates imgX, imgY
                    Viewport.Stage3_CompositeCanvas PDImages.GetActiveImage(), srcCanvas
                End If
        
        End Select
    
    'The left mouse button is *not* down
    Else
        
        'Notify the selection of the currently hovered point of interest, if any
        Dim selPOI As PD_PointOfInterest
        selPOI = SelectionUI.IsCoordSelectionPOI(imgX, imgY, PDImages.GetActiveImage())
        
        If (selPOI <> PDImages.GetActiveImage.MainSelection.GetActiveSelectionPOI(False)) Then
            PDImages.GetActiveImage.MainSelection.SetActiveSelectionPOI selPOI
            Viewport.Stage4_FlipBufferAndDrawUI PDImages.GetActiveImage(), srcCanvas
        Else
            If (g_CurrentTool = SELECT_POLYGON) Then Viewport.Stage4_FlipBufferAndDrawUI PDImages.GetActiveImage(), srcCanvas
        End If
        
    End If
        
End Sub

Public Sub NotifySelectionMouseUp(ByRef srcCanvas As pdCanvas, ByVal Shift As ShiftConstants, ByVal imgX As Single, ByVal imgY As Single, ByVal clickEventAlsoFiring As Boolean, ByVal wasSelectionActiveBeforeMouseEvents As Boolean)
        
    m_MouseDown = False
    m_HasMouseMoved = False
    
    'If a double-click just occurred, reset the flag and exit - do NOT process this click further
    If m_DblClickOccurred Then
        m_DblClickOccurred = False
        Exit Sub
    End If
    
    'Failsafe for bad mice notifications - if we receive an unexpected trigger while ignoring input,
    ' reset all flags but disallow the interrupted action.
    If m_IgnoreUserInput Then
        m_IgnoreUserInput = False
        Exit Sub
    End If
    
    'Ensure other actions don't trigger while this one is still processing (only affects this class!)
    m_IgnoreUserInput = True
    
    'Composite selections have some interesting possible outcomes vs other selection types.
    ' In particular, there are many ways to produce composite selections with no selected pixels.
    ' (e.g. Use "subtract" mode to remove the previous selection completely.)
    '
    'To prevent this from creating a "nothing selected" state, we auto-detect this state on _MouseUp
    ' and initiate a "Remove Selection" action.
    If PDImages.GetActiveImage.MainSelection.IsCompositeSelection() Then
        
        'Some shapes do not auto-generate a composite mask while drawing (for perf reasons).
        ' Ensure a valid composite mask exists before proceeding.
        With PDImages.GetActiveImage.MainSelection
            If ((.GetSelectionShape = ss_Polygon) And .GetPolygonClosedState) Or (.GetSelectionShape = ss_Lasso) Then
                If (.GetSelectionShape = ss_Lasso) Then PDImages.GetActiveImage.MainSelection.SetLassoClosedState True
                PDImages.GetActiveImage.MainSelection.RequestNewMask
            End If
        End With
        
        'Look for at least one selected pixel.
        If PDImages.GetActiveImage.MainSelection.AreAllCoordinatesInvalid(True) Then
            
            'No pixels are selected. Remove the existing selection, then exit.
            Process "Remove selection", , , IIf(wasSelectionActiveBeforeMouseEvents, UNDO_Selection, UNDO_Nothing), g_CurrentTool
            Viewport.Stage3_CompositeCanvas PDImages.GetActiveImage(), srcCanvas
            GoTo FinishedMouseUp
            
        '/Else do nothing; normal handling, below, covers all other bases!
        End If
    
    End If
    
    'In default REPLACE mode, a single in-place click will erase the current selection.
    ' (In other combine modes, this behavior must be ignored or overridden.)
    Dim eraseThisSelection As Boolean: eraseThisSelection = False
    
    Select Case g_CurrentTool
    
        'Most selection tools finalize the current selection on a _MouseUp event
        Case SELECT_RECT, SELECT_CIRC, SELECT_LASSO
        
            'If a selection was being drawn, lock it into place
            If PDImages.GetActiveImage.IsSelectionActive Then
                
                'Check to see if this mouse location is the same as the initial mouse press. If it is, and that particular
                ' point falls outside the selection, clear the selection from the image.
                Dim selBounds As RectF
                selBounds = PDImages.GetActiveImage.MainSelection.GetCornersLockedRect
                
                'We only enable selection erasing on a click in REPLACE mode.  Other combine modes
                ' (add, subtract, etc) do not erase on a click.
                eraseThisSelection = (clickEventAlsoFiring And (IsCoordSelectionPOI(imgX, imgY, PDImages.GetActiveImage()) = poi_Undefined))
                If (Not eraseThisSelection) Then eraseThisSelection = ((selBounds.Width <= 0) And (selBounds.Height <= 0))
                
                If eraseThisSelection Then
                    
                    'In "replace" mode, just remove the active selection (if any)
                    If (toolpanel_Selections.btsCombine.ListIndex = pdsm_Replace) Then
                        Process "Remove selection", , , IIf(wasSelectionActiveBeforeMouseEvents, UNDO_Selection, UNDO_Nothing), g_CurrentTool
                    
                    'In other modes, squash any active selections together into a single selection object.
                    Else
                        PDImages.GetActiveImage.MainSelection.SquashCompositeToRaster
                    End If
                    
                'The mouse is being released after a significant move event, or on a point of interest to the current selection.
                Else
                
                    'If the selection is not raster-type, pass these final mouse coordinates to it
                    If (PDImages.GetActiveImage.MainSelection.GetSelectionShape <> ss_Raster) Then
                        PDImages.GetActiveImage.MainSelection.RequestSquare m_ShiftForConstrain
                        PDImages.GetActiveImage.MainSelection.SetAdditionalCoordinates imgX, imgY
                        SyncTextToCurrentSelection PDImages.GetActiveImageID()
                    End If
                
                    'Check to see if all selection coordinates are invalid (e.g. off-image).
                    ' If they are, forget about this selection.
                    If PDImages.GetActiveImage.MainSelection.AreAllCoordinatesInvalid Then
                        If (PDImages.GetActiveImage.MainSelection.GetSelectionCombineMode() <> pdsm_Replace) Then
                            PDImages.GetActiveImage.MainSelection.SquashCompositeToRaster
                        Else
                            Process "Remove selection", , , IIf(wasSelectionActiveBeforeMouseEvents, UNDO_Selection, UNDO_Nothing), g_CurrentTool
                        End If
                    Else
                        
                        'Depending on the type of transformation that may or may not have been applied, call the appropriate processor function.
                        ' This is required to add the current selection event to the Undo/Redo chain.
                        If (g_CurrentTool = SELECT_LASSO) Then
                        
                            'Creating a new selection
                            If (PDImages.GetActiveImage.MainSelection.GetActiveSelectionPOI = poi_Undefined) Then
                                
                                'Ensure the lasso is closed
                                PDImages.GetActiveImage.MainSelection.SetLassoClosedState True
                                
                                '*Now* we can create the selection
                                Process "Create selection", , PDImages.GetActiveImage.MainSelection.GetSelectionAsXML, UNDO_Selection, g_CurrentTool
                            
                            'Moving an existing selection
                            Else
                                Process "Move selection", , PDImages.GetActiveImage.MainSelection.GetSelectionAsXML, UNDO_Selection, g_CurrentTool
                            End If
                        
                        'All other selection types use identical transform identifiers
                        Else
                        
                            Dim transformType As PD_PointOfInterest
                            transformType = PDImages.GetActiveImage.MainSelection.GetActiveSelectionPOI
                            
                            'Creating a new selection
                            If (transformType = poi_Undefined) Then
                                Process "Create selection", , PDImages.GetActiveImage.MainSelection.GetSelectionAsXML, UNDO_Selection, g_CurrentTool
                            
                            'Moving an existing selection
                            ElseIf (transformType = poi_Interior) Then
                                Process "Move selection", , PDImages.GetActiveImage.MainSelection.GetSelectionAsXML, UNDO_Selection, g_CurrentTool
                                
                            'Anything else is assumed to be resizing an existing selection
                            Else
                                Process "Resize selection", , PDImages.GetActiveImage.MainSelection.GetSelectionAsXML, UNDO_Selection, g_CurrentTool
                                        
                            End If
                        
                        End If
                        
                    End If
                    
                End If
                
                'Creating a brand new selection always necessitates a redraw of the current canvas
                Viewport.Stage3_CompositeCanvas PDImages.GetActiveImage(), srcCanvas
                
            'If the selection is not active, make sure it stays that way
            Else
                PDImages.GetActiveImage.MainSelection.LockRelease
            End If
            
            'Synchronize the selection text box values with the final selection
            SelectionUI.SyncTextToCurrentSelection PDImages.GetActiveImageID()
        
        'As usual, polygon selections require special considerations.
        Case SELECT_POLYGON
            
            'If a selection was being drawn, lock it into place
            If PDImages.GetActiveImage.IsSelectionActive Then
                
                'Check to see if the selection is already locked in.  If it is, we need to check for an "erase selection" click.
                eraseThisSelection = PDImages.GetActiveImage.MainSelection.GetPolygonClosedState And clickEventAlsoFiring
                eraseThisSelection = eraseThisSelection And (IsCoordSelectionPOI(imgX, imgY, PDImages.GetActiveImage()) = -1)
                
                If eraseThisSelection Then
                    Process "Remove selection", , , IIf(wasSelectionActiveBeforeMouseEvents, UNDO_Selection, UNDO_Nothing), g_CurrentTool
                Else
                    
                    'If the polygon is already closed, we want to lock in the newly modified polygon
                    If PDImages.GetActiveImage.MainSelection.GetPolygonClosedState Then
                        
                        'Polygons use a different transform numbering convention than other selection tools, because the number
                        ' of points involved aren't fixed.
                        Dim polyPoint As Long
                        polyPoint = SelectionUI.IsCoordSelectionPOI(imgX, imgY, PDImages.GetActiveImage())
                        
                        'Move selection
                        If (polyPoint = PDImages.GetActiveImage.MainSelection.GetNumOfPolygonPoints) Then
                            Process "Move selection", , PDImages.GetActiveImage.MainSelection.GetSelectionAsXML, UNDO_Selection, g_CurrentTool
                        
                        'Create OR resize, depending on whether the initial point is being clicked for the first time, or whether
                        ' it's being click-moved
                        ElseIf (polyPoint = 0) Then
                            If clickEventAlsoFiring Then
                                Process "Create selection", , PDImages.GetActiveImage.MainSelection.GetSelectionAsXML, UNDO_Selection, g_CurrentTool
                            Else
                                Process "Resize selection", , PDImages.GetActiveImage.MainSelection.GetSelectionAsXML, UNDO_Selection, g_CurrentTool
                            End If
                                
                        'No point of interest means this click lies off-image; this could be a "clear selection" event
                        ' (if a Click event is also firing), or a "move polygon point" event (if the user dragged a
                        ' point off-image).
                        ElseIf (polyPoint = -1) Then
                            
                            'If the user has clicked a blank spot unrelated to the selection, we want to remove the active selection
                            If clickEventAlsoFiring Then
                                Process "Remove selection", , , IIf(wasSelectionActiveBeforeMouseEvents, UNDO_Selection, UNDO_Nothing), g_CurrentTool
                                
                            'If they haven't clicked, this could simply indicate that they dragged a polygon point off the polygon
                            ' and into some new region of the image.
                            Else
                                PDImages.GetActiveImage.MainSelection.SetAdditionalCoordinates imgX, imgY
                                Process "Resize selection", , PDImages.GetActiveImage.MainSelection.GetSelectionAsXML, UNDO_Selection, g_CurrentTool
                            End If
                            
                        'Anything else is a resize
                        Else
                            Process "Resize selection", , PDImages.GetActiveImage.MainSelection.GetSelectionAsXML, UNDO_Selection, g_CurrentTool
                        End If
                        
                        'After all that work, we want to perform one final check to see if all selection coordinates are invalid
                        ' (e.g. if they all lie off-image, which can happen if the user drags all polygon points off-image).
                        ' If they are, we're going to erase this selection, as it's invalid.
                        eraseThisSelection = PDImages.GetActiveImage.MainSelection.IsLockedIn And PDImages.GetActiveImage.MainSelection.AreAllCoordinatesInvalid
                        If eraseThisSelection Then Process "Remove selection", , , IIf(wasSelectionActiveBeforeMouseEvents, UNDO_Selection, UNDO_Nothing), g_CurrentTool
                        
                    'If the polygon is *not* closed, we want to add this as a new polygon point
                    Else
                    
                        'Pass these final mouse coordinates to the selection engine
                        PDImages.GetActiveImage.MainSelection.SetAdditionalCoordinates imgX, imgY
                        
                        'To spare the debug logger from receiving too many events, forcibly prevent logging of this message
                        ' while in debug mode.
                        If (Not wasSelectionActiveBeforeMouseEvents) Then
                            If UserPrefs.GenerateDebugLogs Then
                                Message "Click on the first point to complete the polygon selection", "DONOTLOG"
                            Else
                                Message "Click on the first point to complete the polygon selection"
                            End If
                        End If
                        
                    End If
                
                'End erase vs create check
                End If
                
                'After all selection settings have been applied, forcibly redraw the source canvas
                Viewport.Stage3_CompositeCanvas PDImages.GetActiveImage(), srcCanvas
            
            '(Failsafe check) - if a selection is not active, make sure it stays that way
            Else
                PDImages.GetActiveImage.MainSelection.LockRelease
            End If
            
        'Magic wand selections are actually the easiest to handle, as they don't really support post-creation transforms
        Case SELECT_WAND
            
            'Failsafe check for active selections
            If PDImages.GetActiveImage.IsSelectionActive Then
                
                'Supply the final coordinates to the selection engine (as the user may be dragging around the active point)
                PDImages.GetActiveImage.MainSelection.SetAdditionalCoordinates imgX, imgY
                
                'Check to see if all selection coordinates are invalid (e.g. off-image).
                ' - If they are, forget about this selection.
                ' - If they are not, commit this selection permanently
                eraseThisSelection = PDImages.GetActiveImage.MainSelection.AreAllCoordinatesInvalid(True)
                If eraseThisSelection Then
                    Process "Remove selection", , , IIf(wasSelectionActiveBeforeMouseEvents, UNDO_Selection, UNDO_Nothing), g_CurrentTool
                Else
                    Process "Create selection", , PDImages.GetActiveImage.MainSelection.GetSelectionAsXML, UNDO_Selection, g_CurrentTool
                End If
                
                'Force a redraw of the screen
                Viewport.Stage3_CompositeCanvas PDImages.GetActiveImage(), srcCanvas
                
            'Failsafe check for inactive selections
            Else
                PDImages.GetActiveImage.MainSelection.LockRelease
            End If
            
    End Select
    
FinishedMouseUp:
    m_IgnoreUserInput = False
    
    'If the user pressed a shift/ctrl/alt key to set a temporary combine mode,
    ' and released the key while the mouse was down, we need to reset their original
    ' combine mode now.
    If m_RestoreCombineMode And (m_CurrentShiftState = 0) Then
        m_RestoreCombineMode = False
        toolpanel_Selections.btsCombine.ListIndex = m_OriginalCombineMode
    End If
    
End Sub

Private Sub SyncCombineModeToHotkeys()

    'Use the current state to determine which combine mode to set
    Dim newCombineMode As PD_SelectionCombine
    If (m_CurrentShiftState = vbShiftMask) Then
        newCombineMode = pdsm_Add
    ElseIf (m_CurrentShiftState = vbAltMask) Then
        newCombineMode = pdsm_Subtract
    ElseIf (m_CurrentShiftState = (vbShiftMask Or vbAltMask)) Then
        newCombineMode = pdsm_Intersect
    End If
    
    'Ensure the UI reflects the new setting
    toolpanel_Selections.btsCombine.ListIndex = newCombineMode
    
End Sub

'Use this to populate the text boxes on the main form with the current selection values.
' (Note that this does not cause a screen refresh, by design.)
Public Sub SyncTextToCurrentSelection(ByVal srcImageID As Long)
    
    Dim i As Long
    
    'Only synchronize the text boxes if a selection is active
    Dim selectionIsActive As Boolean
    selectionIsActive = Selections.SelectionsAllowed(False)
    
    Dim selectionToolActive As Boolean
    If selectionIsActive Then
        If PDImages.IsImageActive(srcImageID) Then selectionToolActive = Tools.IsSelectionToolActive()
    End If
    
    'See if a selection exists
    If selectionIsActive And selectionToolActive Then
        
        PDImages.GetImageByID(srcImageID).MainSelection.SuspendAutoRefresh True
        
        'Selection coordinate toolboxes appear on three different selection subpanels: rect, ellipse, and line.
        ' To access their indicies properly, we must calculate an offset.
        Dim subpanelOffset As Long
        subpanelOffset = SelectionUI.GetSelectionSubPanelFromSelectionShape(PDImages.GetImageByID(srcImageID))
        
        'Additional syncing is done if the selection is transformable.
        ' (If it is not transformable, clear and lock the location text boxes.)
        If PDImages.GetImageByID(srcImageID).MainSelection.IsTransformable Then
            
            Dim tmpRectF As RectF
            
            'Different types of selections will display size and position differently
            Select Case PDImages.GetImageByID(srcImageID).MainSelection.GetSelectionShape()
                
                'Rectangular and elliptical selections display left, top, width, height, and aspect ratio (in the form X:Y)
                Case ss_Rectangle, ss_Circle
                    
                    'Indices for spin controls for rectangle selections are:
                    ' 1) size [0, 1]
                    ' 2) aspect ratio [2, 3]
                    ' 3) position [4, 5]
                    ' (add 6 to each value for ellipse selections)
                    Dim baseSizeIndex As Long
                    If (PDImages.GetImageByID(srcImageID).MainSelection.GetSelectionShape() = ss_Rectangle) Then
                        baseSizeIndex = 0
                    Else
                        baseSizeIndex = 6
                    End If
                    
                    tmpRectF = PDImages.GetImageByID(srcImageID).MainSelection.GetCornersLockedRect()
                    
                    toolpanel_Selections.tudSel(baseSizeIndex + 0).Value = tmpRectF.Width
                    toolpanel_Selections.tudSel(baseSizeIndex + 1).Value = tmpRectF.Height
                    
                    'Failsafe DBZ check before calculating aspect ratio
                    If (tmpRectF.Height > 0) Then
                    
                        Dim fracNumerator As Long, fracDenominator As Long
                        PDMath.ConvertToFraction tmpRectF.Width / tmpRectF.Height, fracNumerator, fracDenominator, 0.005
                        
                        'Aspect ratios are typically given in terms of base 10 if possible, so change values like 8:5 to 16:10
                        If (fracDenominator = 5) Then
                            fracNumerator = fracNumerator * 2
                            fracDenominator = fracDenominator * 2
                        End If
                        
                        toolpanel_Selections.tudSel(baseSizeIndex + 2).Value = fracNumerator
                        toolpanel_Selections.tudSel(baseSizeIndex + 3).Value = fracDenominator
                        
                    End If
                    
                    toolpanel_Selections.tudSel(baseSizeIndex + 4).Value = tmpRectF.Left
                    toolpanel_Selections.tudSel(baseSizeIndex + 5).Value = tmpRectF.Top
                    
                    'Also make sure the "lock" icon, if any, matches the current lock state
                    baseSizeIndex = baseSizeIndex \ 2
                    toolpanel_Selections.cmdLock(baseSizeIndex).Value = PDImages.GetImageByID(srcImageID).MainSelection.GetPropertyLockedState(pdsl_Width)
                    toolpanel_Selections.cmdLock(baseSizeIndex + 1).Value = PDImages.GetImageByID(srcImageID).MainSelection.GetPropertyLockedState(pdsl_Height)
                    toolpanel_Selections.cmdLock(baseSizeIndex + 2).Value = PDImages.GetImageByID(srcImageID).MainSelection.GetPropertyLockedState(pdsl_AspectRatio)
                    
            End Select
            
        Else
        
            For i = 0 To toolpanel_Selections.tudSel.Count - 1
                If (toolpanel_Selections.tudSel(i).Min > 0) Then
                    If (toolpanel_Selections.tudSel(i).Value <> toolpanel_Selections.tudSel(i).Min) Then toolpanel_Selections.tudSel(i).Value = toolpanel_Selections.tudSel(i).Min
                Else
                    If (toolpanel_Selections.tudSel(i).Value <> 0) Then toolpanel_Selections.tudSel(i).Value = 0
                End If
            Next i
            
        End If
        
        'Next, sync all non-coordinate information
        If (PDImages.GetImageByID(srcImageID).MainSelection.GetSelectionShape <> ss_Raster) And (PDImages.GetImageByID(srcImageID).MainSelection.GetSelectionShape <> ss_Wand) Then
            toolpanel_Selections.cboSelArea(SelectionUI.GetSelectionSubPanelFromSelectionShape(PDImages.GetImageByID(srcImageID))).ListIndex = PDImages.GetImageByID(srcImageID).MainSelection.GetSelectionProperty_Long(sp_Area)
            toolpanel_Selections.sltSelectionBorder(SelectionUI.GetSelectionSubPanelFromSelectionShape(PDImages.GetImageByID(srcImageID))).Value = PDImages.GetImageByID(srcImageID).MainSelection.GetSelectionProperty_Long(sp_BorderWidth)
        End If
        
        If (toolpanel_Selections.cboSelSmoothing.ListIndex <> PDImages.GetImageByID(srcImageID).MainSelection.GetSelectionProperty_Long(sp_Smoothing)) Then toolpanel_Selections.cboSelSmoothing.ListIndex = PDImages.GetImageByID(srcImageID).MainSelection.GetSelectionProperty_Long(sp_Smoothing)
        If (toolpanel_Selections.sltSelectionFeathering.Value <> PDImages.GetImageByID(srcImageID).MainSelection.GetSelectionProperty_Long(sp_FeatheringRadius)) Then toolpanel_Selections.sltSelectionFeathering.Value = PDImages.GetImageByID(srcImageID).MainSelection.GetSelectionProperty_Long(sp_FeatheringRadius)
        
        'Finally, sync any shape-specific information
        Select Case PDImages.GetImageByID(srcImageID).MainSelection.GetSelectionShape
        
            Case ss_Rectangle
                If (toolpanel_Selections.sltCornerRounding.Value <> PDImages.GetImageByID(srcImageID).MainSelection.GetSelectionProperty_Float(sp_RoundedCornerRadius)) Then toolpanel_Selections.sltCornerRounding.Value = PDImages.GetImageByID(srcImageID).MainSelection.GetSelectionProperty_Float(sp_RoundedCornerRadius)
            
            Case ss_Circle
            
            Case ss_Lasso
                If toolpanel_Selections.sltSmoothStroke.Value <> PDImages.GetImageByID(srcImageID).MainSelection.GetSelectionProperty_Float(sp_SmoothStroke) Then toolpanel_Selections.sltSmoothStroke.Value = PDImages.GetImageByID(srcImageID).MainSelection.GetSelectionProperty_Float(sp_SmoothStroke)
                
            Case ss_Polygon
                If toolpanel_Selections.sltPolygonCurvature.Value <> PDImages.GetImageByID(srcImageID).MainSelection.GetSelectionProperty_Float(sp_PolygonCurvature) Then toolpanel_Selections.sltPolygonCurvature.Value = PDImages.GetImageByID(srcImageID).MainSelection.GetSelectionProperty_Float(sp_PolygonCurvature)
                
            Case ss_Wand
                If toolpanel_Selections.btsWandArea.ListIndex <> PDImages.GetImageByID(srcImageID).MainSelection.GetSelectionProperty_Long(sp_WandSearchMode) Then toolpanel_Selections.btsWandArea.ListIndex = PDImages.GetImageByID(srcImageID).MainSelection.GetSelectionProperty_Long(sp_WandSearchMode)
                If toolpanel_Selections.btsWandMerge.ListIndex <> PDImages.GetImageByID(srcImageID).MainSelection.GetSelectionProperty_Long(sp_WandSampleMerged) Then toolpanel_Selections.btsWandMerge.ListIndex = PDImages.GetImageByID(srcImageID).MainSelection.GetSelectionProperty_Long(sp_WandSampleMerged)
                If toolpanel_Selections.sltWandTolerance.Value <> PDImages.GetImageByID(srcImageID).MainSelection.GetSelectionProperty_Float(sp_WandTolerance) Then toolpanel_Selections.sltWandTolerance.Value = PDImages.GetImageByID(srcImageID).MainSelection.GetSelectionProperty_Float(sp_WandTolerance)
                If toolpanel_Selections.cboWandCompare.ListIndex <> PDImages.GetImageByID(srcImageID).MainSelection.GetSelectionProperty_Long(sp_WandCompareMethod) Then toolpanel_Selections.cboWandCompare.ListIndex = PDImages.GetImageByID(srcImageID).MainSelection.GetSelectionProperty_Long(sp_WandCompareMethod)
        
        End Select
        
        PDImages.GetImageByID(srcImageID).MainSelection.SuspendAutoRefresh False
    
    'A selection is *not* active; disable various selection-related UI options
    Else
        
        'If a selection exists, we need to leave available menu commands like "remove selection", etc.
        Interface.SetUIGroupState PDUI_Selections, selectionIsActive
        
        'Transformable settings do *not* need to be available
        Interface.SetUIGroupState PDUI_SelectionTransforms, False
        
        'This branch is only followed if a selection is *not* active but a selection tool *is* active, in which case
        ' we need to disable some commands on the selection toolbar.
        If Tools.IsSelectionToolActive Then
            For i = 0 To toolpanel_Selections.tudSel.Count - 1
                If (toolpanel_Selections.tudSel(i).Min > 0) Then
                    If (toolpanel_Selections.tudSel(i).Value <> toolpanel_Selections.tudSel(i).Min) Then toolpanel_Selections.tudSel(i).Value = toolpanel_Selections.tudSel(i).Min
                Else
                    If (toolpanel_Selections.tudSel(i).Value <> 0) Then toolpanel_Selections.tudSel(i).Value = 0
                End If
            Next i
        End If
        
    End If
    
    'Update PD's central status bar as well
    FormMain.MainCanvas(0).SetSelectionState selectionIsActive
    
End Sub
