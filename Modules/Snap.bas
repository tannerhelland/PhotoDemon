Attribute VB_Name = "Snap"
'***************************************************************************
'Snap-to-target Handler
'Copyright 2024-2024 by Tanner Helland
'Created: 16/April/24
'Last updated: 19/April/24
'Last update: add support for snapping to centerlines
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
    pdst_Centerline = 2
    pdst_Layer = 3
End Enum

#If False Then
    Private Const pdst_Global = 0, pdst_CanvasEdge = 1, pdst_Centerline = 2, pdst_Layer = 3
#End If

'When snapping coordinates, we need to compare all possible snap targets and choose the best independent
' x and y snap coordinate (assuming they fall beneath the snap threshold for the current zoom level).
' Two distances are provided: one each for left/right (or top/bottom).
Private Type SnapComparison
    cValue As Double
    cDistance1 As Double    'Left/Top distance
    cDistance2 As Double    'Right/Bottom distance
    cDistanceCX As Double    'X-Center distance (only if enabled)
    cDistanceCY As Double    'Y-Center distance (only if enabled)
    cCenterComparison As Boolean    'Set to TRUE if center distance is smallest distance; this is relevant for rects
                                    ' and point lists, because we need to snap the *center*, not the boundaries
End Type

'To improve performance, snap-to settings are cached locally (instead of traveling out to
' the user preference engine on every call).
Private m_SnapGlobal As Boolean, m_SnapToCanvasEdge As Boolean, m_SnapToCenterline As Boolean, m_SnapToLayer As Boolean
Private m_SnapDistance As Long

'Returns TRUE if *any* snap-to-edge behaviors are enabled.  Useful for skipping all snap checks.
Public Function GetSnap_Any() As Boolean
    GetSnap_Any = m_SnapGlobal
    If m_SnapGlobal Then
        GetSnap_Any = m_SnapToCanvasEdge Or m_SnapToCenterline Or m_SnapToLayer
    End If
End Function

Public Function GetSnap_CanvasEdge() As Boolean
    GetSnap_CanvasEdge = m_SnapToCanvasEdge
End Function

Public Function GetSnap_Centerline() As Boolean
    GetSnap_Centerline = m_SnapToCenterline
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

Public Function GetSnap_Layer() As Boolean
    GetSnap_Layer = m_SnapToLayer
End Function

Public Sub SetSnap_CanvasEdge(ByVal newState As Boolean)
    m_SnapToCanvasEdge = newState
End Sub

Public Sub SetSnap_Centerline(ByVal newState As Boolean)
    m_SnapToCenterline = newState
End Sub

Public Sub SetSnap_Distance(ByVal newDistance As Long)
    m_SnapDistance = newDistance
    If (m_SnapDistance < 1) Then m_SnapDistance = 1
    If (m_SnapDistance > 255) Then m_SnapDistance = 255     'GIMP uses a 255 max value; that seems reasonable
End Sub

Public Sub SetSnap_Global(ByVal newState As Boolean)
    m_SnapGlobal = newState
End Sub

Public Sub SetSnap_Layer(ByVal newState As Boolean)
    m_SnapToLayer = newState
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
            
        Case pdst_Centerline
            If (Not forceInsteadOfToggle) Then newState = Not Snap.GetSnap_Centerline()
            Snap.SetSnap_Centerline newState
            UserPrefs.SetPref_Boolean "Interface", "snap-centerline", newState
            Menus.SetMenuChecked "snap_centerline", newState
        
        Case pdst_Layer
            If (Not forceInsteadOfToggle) Then newState = Not Snap.GetSnap_Layer()
            Snap.SetSnap_Layer newState
            UserPrefs.SetPref_Boolean "Interface", "snap-layer", newState
            Menus.SetMenuChecked "snap_layer", newState
            
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
                    .cCenterComparison = False
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
                    .cCenterComparison = False
                End If
            End With
        Next i
    Next j
    
    'If centerline snapping is enabled, repeat the above steps, but for the center point of the list only
    Dim pathTest As pd2DPath, pathRect As RectF
    Set pathTest = New pd2DPath
    pathTest.AddLines numOfPoints, VarPtr(srcPoints(0))
    pathRect = pathTest.GetPathBoundariesF()
    
    Dim cX As Double, cY As Double
    cX = pathRect.Left + pathRect.Width * 0.5
    cY = pathRect.Top + pathRect.Height * 0.5
    
    For i = 0 To numXSnaps - 1
        With xSnaps(i)
            .cDistanceCX = PDMath.DistanceOneDimension(cX, .cValue)
            If (.cDistanceCX < minDistX) Then
                minDistX = .cDistanceCX
                idxSmallestX = i
                .cCenterComparison = True
            End If
        End With
    Next i
    
    For i = 0 To numYSnaps - 1
        With ySnaps(i)
            .cDistanceCY = PDMath.DistanceOneDimension(cY, .cValue)
            If (.cDistanceCY < minDistY) Then
                minDistY = .cDistanceCY
                idxSmallestY = i
                .cCenterComparison = True
            End If
        End With
    Next i
    
    'Determine the minimum snap distance required for this zoom value.
    Dim snapThreshold As Double
    snapThreshold = GetSnapDistanceScaledForZoom()
    
    'If the minimum value falls beneath the minimum snap distance, snap away!
    If (minDistX < snapThreshold) Then
        
        'Center comparisons require us to align the center point of the rect
        If xSnaps(idxSmallestX).cCenterComparison Then
            dstOffsetX = xSnaps(idxSmallestX).cValue - (pathRect.Left + pathRect.Width * 0.5)
        Else
            dstOffsetX = (xSnaps(idxSmallestX).cValue - srcPoints(idxSmallestPointX).x)
        End If
        
    End If
    
    If (minDistY < snapThreshold) Then
        
        'Center comparisons require us to align the center point of the rect
        If ySnaps(idxSmallestY).cCenterComparison Then
            dstOffsetY = ySnaps(idxSmallestY).cValue - (pathRect.Top + pathRect.Height * 0.5)
        Else
            dstOffsetY = (ySnaps(idxSmallestY).cValue - srcPoints(idxSmallestPointY).y)
        End If
    End If
    
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
                .cCenterComparison = False
            End If
            .cDistance2 = PDMath.DistanceOneDimension(compareRectF.Right, .cValue)
            If (.cDistance2 < minDistX) Then
                minDistX = .cDistance2
                idxSmallestX = i
                .cCenterComparison = False
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
                .cCenterComparison = False
            End If
            .cDistance2 = PDMath.DistanceOneDimension(compareRectF.Bottom, .cValue)
            If (.cDistance2 < minDistY) Then
                minDistY = .cDistance2
                idxSmallestY = i
                .cCenterComparison = False
            End If
        End With
    Next i
    
    'If centerline snapping is enabled, repeat the above steps, but for the center point of the rect only
    Dim cX As Double, cY As Double
    cX = compareRectF.Left + (compareRectF.Right - compareRectF.Left) * 0.5
    cY = compareRectF.Top + (compareRectF.Bottom - compareRectF.Top) * 0.5
    
    For i = 0 To numXSnaps - 1
        With xSnaps(i)
            .cDistanceCX = PDMath.DistanceOneDimension(cX, .cValue)
            If (.cDistanceCX < minDistX) Then
                minDistX = .cDistanceCX
                idxSmallestX = i
                .cCenterComparison = True
            End If
        End With
    Next i
    
    For i = 0 To numYSnaps - 1
        With ySnaps(i)
            .cDistanceCY = PDMath.DistanceOneDimension(cY, .cValue)
            If (.cDistanceCY < minDistY) Then
                minDistY = .cDistanceCY
                idxSmallestY = i
                .cCenterComparison = True
            End If
        End With
    Next i
    
    'Determine the minimum snap distance required for this zoom value.
    Dim snapThreshold As Double
    snapThreshold = GetSnapDistanceScaledForZoom()
    
    'If the minimum value falls beneath the minimum snap distance, snap away!
    If (minDistX < snapThreshold) Then
        
        'Center comparisons require us to align the center point of the rect
        If xSnaps(idxSmallestX).cCenterComparison Then
            dstRectF.Left = xSnaps(idxSmallestX).cValue - (compareRectF.Right - compareRectF.Left) * 0.5
        
        'Otherwise, align the left or right boundary of the rect, as relevant
        Else
            If (xSnaps(idxSmallestX).cDistance1 < xSnaps(idxSmallestX).cDistance2) Then
                dstRectF.Left = xSnaps(idxSmallestX).cValue
            Else
                dstRectF.Left = xSnaps(idxSmallestX).cValue - dstRectF.Width
            End If
        End If
        
    End If
    
    If (minDistY < snapThreshold) Then
        If ySnaps(idxSmallestY).cCenterComparison Then
            dstRectF.Top = ySnaps(idxSmallestY).cValue - (compareRectF.Bottom - compareRectF.Top) * 0.5
        Else
            If (ySnaps(idxSmallestY).cDistance1 < ySnaps(idxSmallestY).cDistance2) Then
                dstRectF.Top = ySnaps(idxSmallestY).cValue
            Else
                dstRectF.Top = ySnaps(idxSmallestY).cValue - dstRectF.Height
            End If
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
    
    'Centerline (of canvas only; layers is handled below)
    If Snap.GetSnap_Centerline() Then
        If (UBound(dstSnaps) < GetSnapTargets_X) Then ReDim Preserve dstSnaps(0 To GetSnapTargets_X * 2 - 1) As SnapComparison
        dstSnaps(GetSnapTargets_X).cValue = Int(PDImages.GetActiveImage.Width / 2)
        GetSnapTargets_X = GetSnapTargets_X + 1
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
    
    'Centerline (of canvas only; layers is handled below)
    If Snap.GetSnap_Centerline() Then
        If (UBound(dstSnaps) < GetSnapTargets_Y) Then ReDim Preserve dstSnaps(0 To GetSnapTargets_Y * 2 - 1) As SnapComparison
        dstSnaps(GetSnapTargets_Y).cValue = Int(PDImages.GetActiveImage.Height / 2)
        GetSnapTargets_Y = GetSnapTargets_Y + 1
    End If
    
    'TODO: more snap targets in the future...
    
End Function
