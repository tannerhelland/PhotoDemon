Attribute VB_Name = "Snap"
'***************************************************************************
'Snap-to-target Handler
'Copyright 2024-2024 by Tanner Helland
'Created: 16/April/24
'Last updated: 16/April/24
'Last update: migrate all snap setting and behavior management to one central place
'
'In 2024, snap-to-target support was added to various PhotoDemon tools.  Thank you to all the users
' who suggested this feature!
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

Public Enum PD_SnapTargets
    pdst_Global = 0
    pdst_CanvasEdge = 1
End Enum

#If False Then
    Private Const pdst_Global = 0, pdst_CanvasEdge = 1
#End If

'When snapping coordinates, we need to compare all possible snap targets and choose the best independent
' x and y snap coordinate (assuming they fall beneath the snap threshold for the current zoom level).
' Two distances are provided: one each for left/right (or top/bottom).
Private Type SnapComparison
    cValue As Double
    cDistance1 As Double
    cDistance2 As Double
End Type

'To improve performance, snap-to settings are cached locally (instead of traveling out to
' the user preference engine on every call).
Private m_SnapGlobal As Boolean, m_SnapToCanvasEdge As Boolean, m_SnapDistance As Long

'Returns TRUE if *any* snap-to-edge behaviors are enabled.  Useful for skipping all snap checks.
Public Function GetSnap_Any() As Boolean
    GetSnap_Any = m_SnapGlobal
    If m_SnapGlobal Then
        GetSnap_Any = m_SnapToCanvasEdge
        'TODO: OR against other snap options when added
    End If
End Function

Public Function GetSnap_CanvasEdge() As Boolean
    GetSnap_CanvasEdge = m_SnapToCanvasEdge
End Function

Public Function GetSnap_Distance() As Long
    
    GetSnap_Distance = m_SnapDistance
    
    'Failsafe only; should never trigger
    If (GetSnap_Distance < 1) Then GetSnap_Distance = 8
    
End Function

'Returns TRUE if the top-level "View > Snap" menu is checked.  Note that the user can enable/disable
' individual snap targets regardless of this setting, but if this setting is FALSE, we must ignore all
' other snap options.  (This is how Photoshop behaves; the top-level Snap setting is mapped to a
' keyboard accelerator so the user can quickly enable/disable snap behavior without losing current
' per-target snap settings.)
Public Function GetSnap_Global() As Boolean
    GetSnap_Global = m_SnapGlobal
End Function

Public Sub SetSnap_CanvasEdge(ByVal newState As Boolean)
    m_SnapToCanvasEdge = newState
End Sub

Public Sub SetSnap_Distance(ByVal newDistance As Long)
    m_SnapDistance = newDistance
    If (m_SnapDistance < 1) Then m_SnapDistance = 1
    If (m_SnapDistance > 255) Then m_SnapDistance = 255     'GIMP uses a 255 max value; that seems reasonable
End Sub

Public Sub SetSnap_Global(ByVal newState As Boolean)
    m_SnapGlobal = newState
End Sub

'Toggle one of the "snap to..." settings in the View menu.
' To forcibly set to a specific state (instead of toggling), set the forceInsteadOfToggle param to TRUE.
Public Sub ToggleSnapOptions(ByVal snapTarget As PD_SnapTargets, Optional ByVal forceInsteadOfToggle As Boolean = False, Optional ByVal newState As Boolean = True)
    
    'While calculating which on-screen menu to update, we also need to relay changes to two places:
    ' 1) the tools_move module (which handles actual snap calculations)
    ' 2) the user preferences file (to ensure everything is synchronized between sessions)
    Select Case snapTarget
        Case pdst_Global
            If (Not forceInsteadOfToggle) Then newState = Not Snap.GetSnap_Global()
            Snap.SetSnap_Global newState
            UserPrefs.SetPref_Boolean "Interface", "snap-global", newState
            Menus.SetMenuChecked "snap_global", newState
            
        Case pdst_CanvasEdge
            If (Not forceInsteadOfToggle) Then newState = Not Snap.GetSnap_CanvasEdge()
            Snap.SetSnap_CanvasEdge newState
            UserPrefs.SetPref_Boolean "Interface", "snap-canvas-edge", newState
            Menus.SetMenuChecked "snap_canvasedge", newState
            
    End Select
    
End Sub

'Snap the passed point to any relevant snap targets (based on the user's current snap settings).
Public Sub SnapPointByMoving(ByRef srcPointF As PointFloat, ByRef dstPointF As PointFloat)
    
    'If no snap targets exist (because the user has disabled snapping), ensure the destination point
    ' mirrors the source point
    dstPointF = srcPointF
    
    'Skip any further processing if the user hasn't enabled snapping
    If (Not Snap.GetSnap_Any()) Then Exit Sub
    
    'Start by constructing a list of potential snap targets, based on current user settings.
    Dim xSnaps() As SnapComparison, ySnaps() As SnapComparison, numXSnaps As Long, numYSnaps As Long
    numXSnaps = GetSnapTargets_X(xSnaps)
    numYSnaps = GetSnapTargets_Y(ySnaps)
    
    'Ensure some snap targets exist
    If (numXSnaps = 0) Or (numYSnaps = 0) Then Exit Sub
    
    'We now have a list of snap comparison targets.  We don't care what these targets represent -
    ' we just want to find the "best" one from each list.
    Dim i As Long, idxSmallestX As Long, minDistX As Double
    
    'Set the minimum distance to an arbitrarily huge number, then find minimum x-distances
    minDistX = DOUBLE_MAX
    For i = 0 To numXSnaps - 1
        With xSnaps(i)
            .cDistance1 = PDMath.DistanceOneDimension(srcPointF.x, .cValue)
            If (.cDistance1 < minDistX) Then
                minDistX = .cDistance1
                idxSmallestX = i
            End If
        End With
    Next i
    
    'Repeat all the above steps for y-coordinates
    Dim idxSmallestY As Long, minDistY As Double
    minDistY = DOUBLE_MAX
    
    For i = 0 To numYSnaps - 1
        With ySnaps(i)
            .cDistance1 = PDMath.DistanceOneDimension(srcPointF.y, .cValue)
            If (.cDistance1 < minDistY) Then
                minDistY = .cDistance1
                idxSmallestY = i
            End If
        End With
    Next i
    
    'Determine the minimum snap distance required for this zoom value.
    Dim snapThreshold As Double
    snapThreshold = GetSnapDistanceScaledForZoom()
    
    'If the minimum value falls beneath the minimum snap distance, snap away!
    If (minDistX < snapThreshold) Then dstPointF.x = xSnaps(idxSmallestX).cValue
    If (minDistY < snapThreshold) Then dstPointF.y = ySnaps(idxSmallestY).cValue
    
End Sub

'Given a list of points, compare each to all snap points and find the *best* match among them.  Based on that,
' return the x/y offset required to move the best-match point onto the snap target.
Public Sub SnapPointListByMoving(ByRef srcPoints() As PointFloat, ByVal numOfPoints As Long, ByRef dstOffsetX As Long, ByRef dstOffsetY As Long)
    
    dstOffsetX = 0
    dstOffsetY = 0
    
    'Failsafe only; caller (in PD) will never set this to <= 0
    If (numOfPoints <= 0) Then Exit Sub
    
    'Skip any further processing if the user hasn't enabled snapping
    If (Not Snap.GetSnap_Any()) Then Exit Sub
    
    'Start by constructing a list of potential snap targets, based on current user settings.
    Dim xSnaps() As SnapComparison, ySnaps() As SnapComparison, numXSnaps As Long, numYSnaps As Long
    numXSnaps = GetSnapTargets_X(xSnaps)
    numYSnaps = GetSnapTargets_Y(ySnaps)
    
    'Ensure some snap targets exist
    If (numXSnaps = 0) Or (numYSnaps = 0) Then Exit Sub
    
    'We now have a list of snap comparison targets.  We don't care what these targets represent -
    ' we just want to find the "best" one from each list.
    Dim idxSmallestX As Long, idxSmallestPointX As Long, minDistX As Double
    
    'Set the minimum distance to an arbitrarily huge number, then find minimum x-distances
    minDistX = DOUBLE_MAX
    
    Dim i As Long, j As Long
    For j = 0 To numOfPoints - 1
        For i = 0 To numXSnaps - 1
            With xSnaps(i)
                .cDistance1 = PDMath.DistanceOneDimension(srcPoints(j).x, .cValue)
                If (.cDistance1 < minDistX) Then
                    minDistX = .cDistance1
                    idxSmallestX = i
                    idxSmallestPointX = j
                End If
            End With
        Next i
    Next j
    
    'Repeat all the above steps for y-coordinates
    Dim idxSmallestY As Long, idxSmallestPointY As Long, minDistY As Double
    minDistY = DOUBLE_MAX
    
    For j = 0 To numOfPoints - 1
        For i = 0 To numYSnaps - 1
            With ySnaps(i)
                .cDistance1 = PDMath.DistanceOneDimension(srcPoints(j).y, .cValue)
                If (.cDistance1 < minDistY) Then
                    minDistY = .cDistance1
                    idxSmallestY = i
                    idxSmallestPointY = j
                End If
            End With
        Next i
    Next j
    
    'Determine the minimum snap distance required for this zoom value.
    Dim snapThreshold As Double
    snapThreshold = GetSnapDistanceScaledForZoom()
    
    'If the minimum value falls beneath the minimum snap distance, snap away!
    If (minDistX < snapThreshold) Then dstOffsetX = (xSnaps(idxSmallestX).cValue - srcPoints(idxSmallestPointX).x)
    If (minDistY < snapThreshold) Then dstOffsetY = (ySnaps(idxSmallestY).cValue - srcPoints(idxSmallestPointY).y)
    
End Sub

'Snap the passed rect to any relevant snap targets (based on the user's current snap settings).
' Because this function snaps only by moving the target rect, it is guaranteed that only the
' top and left values will be changed by the function (width/height will *not*).
Public Sub SnapRectByMoving(ByRef srcRectF As RectF, ByRef dstRectF As RectF)
    
    'By default, return the same rect.  (This is important if the user has disabled snapping.)
    dstRectF = srcRectF
    
    'Skip any further processing if the user hasn't enabled snapping
    If (Not Snap.GetSnap_Any()) Then Exit Sub
    
    'Start by constructing a list of potential snap targets, based on current user settings.
    Dim xSnaps() As SnapComparison, ySnaps() As SnapComparison, numXSnaps As Long, numYSnaps As Long
    numXSnaps = GetSnapTargets_X(xSnaps)
    numYSnaps = GetSnapTargets_Y(ySnaps)
    
    'Ensure some snap targets exist
    If (numXSnaps = 0) Or (numYSnaps = 0) Then Exit Sub
    
    'We now have a list of snap comparison targets.  We don't care what these targets represent -
    ' we just want to find the "best" one from each list.
    
    'Convert the source snap rectangle into a right/bottom rect (instead of a default width/height one)
    Dim compareRectF As RectF_RB
    compareRectF.Left = srcRectF.Left
    compareRectF.Top = srcRectF.Top
    compareRectF.Right = srcRectF.Left + srcRectF.Width - 1
    compareRectF.Bottom = srcRectF.Top + srcRectF.Height - 1
    
    Dim i As Long, idxSmallestX As Long, minDistX As Double
    
    'Set the minimum distance to an arbitrarily huge number, then find the smallest x-distance
    minDistX = DOUBLE_MAX
    For i = 0 To numXSnaps - 1
        With xSnaps(i)
            .cDistance1 = PDMath.DistanceOneDimension(compareRectF.Left, .cValue)
            If (.cDistance1 < minDistX) Then
                minDistX = .cDistance1
                idxSmallestX = i
            End If
            .cDistance2 = PDMath.DistanceOneDimension(compareRectF.Right, .cValue)
            If (.cDistance2 < minDistX) Then
                minDistX = .cDistance2
                idxSmallestX = i
            End If
        End With
    Next i
    
    'Repeat all the above steps for y-coordinates
    Dim idxSmallestY As Long, minDistY As Double
    minDistY = DOUBLE_MAX
    
    For i = 0 To numYSnaps - 1
        With ySnaps(i)
            .cDistance1 = PDMath.DistanceOneDimension(compareRectF.Top, .cValue)
            If (.cDistance1 < minDistY) Then
                minDistY = .cDistance1
                idxSmallestY = i
            End If
            .cDistance2 = PDMath.DistanceOneDimension(compareRectF.Bottom, .cValue)
            If (.cDistance2 < minDistY) Then
                minDistY = .cDistance2
                idxSmallestY = i
            End If
        End With
    Next i
    
    'Determine the minimum snap distance required for this zoom value.
    Dim snapThreshold As Double
    snapThreshold = GetSnapDistanceScaledForZoom()
    
    'If the minimum value falls beneath the minimum snap distance, snap away!
    If (minDistX < snapThreshold) Then
        If (xSnaps(idxSmallestX).cDistance1 < xSnaps(idxSmallestX).cDistance2) Then
            dstRectF.Left = xSnaps(idxSmallestX).cValue
        Else
            dstRectF.Left = xSnaps(idxSmallestX).cValue - dstRectF.Width
        End If
    End If
    
    If (minDistY < snapThreshold) Then
        If (ySnaps(idxSmallestY).cDistance1 < ySnaps(idxSmallestY).cDistance2) Then
            dstRectF.Top = ySnaps(idxSmallestY).cValue
        Else
            dstRectF.Top = ySnaps(idxSmallestY).cValue - dstRectF.Height
        End If
    End If
    
End Sub

Private Function GetSnapDistanceScaledForZoom() As Double
    GetSnapDistanceScaledForZoom = Snap.GetSnap_Distance() * (1# / Zoom.GetZoomRatioFromIndex(PDImages.GetActiveImage.ImgViewport.GetZoomIndex))
End Function

'Get a list of current x-snap targets (determined by user settings).
' RETURNS: number of entries in the list, or 0 if snapping is disabled by the user.
Private Function GetSnapTargets_X(ByRef dstSnaps() As SnapComparison) As Long
    
    'Start with some arbitrarily sized list (these will be enlarged as necessary)
    ReDim dstSnaps(0 To 15) As SnapComparison
    GetSnapTargets_X = 0
    
    'Canvas edges first
    If Snap.GetSnap_CanvasEdge() Then
        
        'Ensure at space is available in the target array
        If (UBound(dstSnaps) < GetSnapTargets_X + 1) Then ReDim Preserve dstSnaps(0 To GetSnapTargets_X * 2 - 1) As SnapComparison
        
        'Add canvas boundaries to the snap list
        dstSnaps(GetSnapTargets_X).cValue = 0#
        dstSnaps(GetSnapTargets_X + 1).cValue = PDImages.GetActiveImage.Width
        GetSnapTargets_X = GetSnapTargets_X + 2
        
    End If
    
    'TODO: more snap targets in the future...
    
End Function

'Get a list of current y-snap targets (determined by user settings).
' RETURNS: number of entries in the list, or 0 if snapping is disabled by the user.
Private Function GetSnapTargets_Y(ByRef dstSnaps() As SnapComparison) As Long
    
    'Start with some arbitrarily sized list (these will be enlarged as necessary)
    ReDim dstSnaps(0 To 15) As SnapComparison
    GetSnapTargets_Y = 0
    
    'Canvas edges first
    If Snap.GetSnap_CanvasEdge() Then
        
        'Ensure at space is available in the target array
        If (UBound(dstSnaps) < GetSnapTargets_Y + 1) Then ReDim Preserve dstSnaps(0 To GetSnapTargets_Y * 2 - 1) As SnapComparison
        
        'Add canvas boundaries to the snap list
        dstSnaps(GetSnapTargets_Y).cValue = 0#
        dstSnaps(GetSnapTargets_Y + 1).cValue = PDImages.GetActiveImage.Height
        GetSnapTargets_Y = GetSnapTargets_Y + 2
        
    End If
    
    'TODO: more snap targets in the future...
    
End Function
