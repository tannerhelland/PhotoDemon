Attribute VB_Name = "Snap"
'***************************************************************************
'Snap-to-target Handler
'Copyright 2024-2026 by Tanner Helland
'Created: 16/April/24
'Last updated: 16/January/25
'Last update: implement angle snapping for the move/size tool
'
'In 2024, snap-to-target support was added to various PhotoDemon tools.  Thank you to all the users
' who suggested this feature!
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Prints snap info to the debug log; do NOT activate in production builds
Private Const DEBUG_SNAP_REPORTING As Boolean = False

Public Enum PD_SnapTargets
    pdst_Global = 0
    pdst_CanvasEdge = 1
    pdst_Centerline = 2
    pdst_Layer = 4
    pdst_Angle90 = 8
    pdst_Angle45 = 16
    pdst_Angle30 = 32
End Enum

#If False Then
    Private Const pdst_Global = 0, pdst_CanvasEdge = 1, pdst_Centerline = 2, pdst_Layer = 4
    Private Const pdst_Angle90 = 8, pdst_Angle45 = 16, pdst_Angle30 = 32
#End If

'When snapping coordinates, we need to compare all possible snap targets and choose the best independent
' x and y snap coordinate (assuming they fall beneath the snap threshold for the current zoom level).
' Two distances are provided: one each for left/right (or top/bottom).
Private Type SnapComparison
    cValue As Double
    cDistance1 As Double    'Left/Top distance
    cDistance2 As Double    'Right/Bottom distance
    cDistanceCX As Double   'X-Center distance (only if enabled)
    cDistanceCY As Double   'Y-Center distance (only if enabled)
    cCenterComparison As Boolean    'Set to TRUE if center distance is smallest distance; this is relevant for rects
                                    ' and point lists, because we need to snap the *center*, not the boundaries
    cSnapSource As Integer  'OR'd flags that define the source of this point (e.g. LAYER or CENTERLINE for a layer's center)
    cSnapLinePt1 As PointFloat  'Relevant line defining what's being snapped (e.g. the left line of a layer or similar).
    cSnapLinePt2 As PointFloat  ' This line is used to generate an on-screen smart guide if enabled by the user.
    cSnapName As String         'Debug only; stores the name of the snap target (useful for figuring out wtf is being snapped)
End Type

'To improve performance, snap-to settings are cached locally (instead of traveling out to
' the user preference engine on every call).
Private m_SnapGlobal As Boolean, m_SnapToCanvasEdge As Boolean, m_SnapToCenterline As Boolean, m_SnapToLayer As Boolean
Private m_SnapAngle90 As Boolean, m_SnapAngle45 As Boolean, m_SnapAngle30 As Boolean
Private m_SnapDistance As Long, m_SnapDegrees As Single

'When a snap request was successful, these flags are set to TRUE and points defining the snapped line are
' also generated (so the renderer can display a smart guide, if enabled).
Private m_SnappedX As Boolean, m_SnappedXPt1 As PointFloat, m_SnappedXPt2 As PointFloat, m_SnappedXName As String
Private m_SnappedY As Boolean, m_SnappedYPt1 As PointFloat, m_SnappedYPt2 As PointFloat, m_SnappedYName As String

'Returns TRUE if *any* snap-to-edge behaviors are enabled.  Useful for skipping all snap checks.
Public Function GetSnap_Any() As Boolean
    GetSnap_Any = m_SnapGlobal
    If m_SnapGlobal Then
        GetSnap_Any = m_SnapToCanvasEdge Or m_SnapToCenterline Or m_SnapToLayer
        GetSnap_Any = GetSnap_Any Or m_SnapAngle90 Or m_SnapAngle45 Or m_SnapAngle30
    End If
End Function

Public Function GetSnap_Angle90() As Boolean
    GetSnap_Angle90 = m_SnapAngle90
End Function

Public Function GetSnap_Angle45() As Boolean
    GetSnap_Angle45 = m_SnapAngle45
End Function

Public Function GetSnap_Angle30() As Boolean
    GetSnap_Angle30 = m_SnapAngle30
End Function

Public Function GetSnap_CanvasEdge() As Boolean
    GetSnap_CanvasEdge = m_SnapToCanvasEdge
End Function

Public Function GetSnap_Centerline() As Boolean
    GetSnap_Centerline = m_SnapToCenterline
End Function

Public Function GetSnap_Degrees() As Single

    GetSnap_Degrees = m_SnapDegrees
    
    'Failsafe only; should never trigger
    If (GetSnap_Degrees < 1!) Then GetSnap_Degrees = 5!
    If (GetSnap_Degrees > 15!) Then GetSnap_Degrees = 15!
    
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

Public Sub GetSnappedX_SmartGuide(ByRef dstPt1 As PointFloat, ByRef dstPt2 As PointFloat)
    dstPt1 = m_SnappedXPt1
    dstPt2 = m_SnappedXPt2
End Sub

Public Sub GetSnappedY_SmartGuide(ByRef dstPt1 As PointFloat, ByRef dstPt2 As PointFloat)
    dstPt1 = m_SnappedYPt1
    dstPt2 = m_SnappedYPt2
End Sub

Public Function IsSnapped_X() As Boolean
    IsSnapped_X = m_SnappedX
End Function

Public Function IsSnapped_Y() As Boolean
    IsSnapped_Y = m_SnappedY
End Function

'When interacting with a POI on the canvas that doesn't support snapping (like rotation),
' call this function to prevent rendering of smart guides from past interactions.
Public Sub NotifyNoSnapping()
    m_SnappedX = False
    m_SnappedY = False
End Sub

'When locking aspect ratio and resizing a layer, we can't snap in both directions.  The caller
' chooses which direction to prioritize, and forcibly disables the other.
Public Sub NotifyNoSnapping_X()
    m_SnappedX = False
End Sub

Public Sub NotifyNoSnapping_Y()
    m_SnappedY = False
End Sub

Public Sub SetSnap_Angle90(ByVal newState As Boolean)
    m_SnapAngle90 = newState
End Sub

Public Sub SetSnap_Angle45(ByVal newState As Boolean)
    m_SnapAngle45 = newState
End Sub

Public Sub SetSnap_Angle30(ByVal newState As Boolean)
    m_SnapAngle30 = newState
End Sub

Public Sub SetSnap_CanvasEdge(ByVal newState As Boolean)
    m_SnapToCanvasEdge = newState
End Sub

Public Sub SetSnap_Centerline(ByVal newState As Boolean)
    m_SnapToCenterline = newState
End Sub

Public Sub SetSnap_Degrees(ByVal newDegrees As Single)
    m_SnapDegrees = newDegrees
    If (m_SnapDegrees < 1!) Then m_SnapDegrees = 1!
    If (m_SnapDegrees > 15!) Then m_SnapDegrees = 15!
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
            
        Case pdst_Angle90
            If (Not forceInsteadOfToggle) Then newState = Not Snap.GetSnap_Angle90()
            Snap.SetSnap_Angle90 newState
            UserPrefs.SetPref_Boolean "Interface", "snap-angle-90", newState
            Menus.SetMenuChecked "snap_angle_90", newState
            
        Case pdst_Angle45
            If (Not forceInsteadOfToggle) Then newState = Not Snap.GetSnap_Angle45()
            Snap.SetSnap_Angle45 newState
            UserPrefs.SetPref_Boolean "Interface", "snap-angle-45", newState
            Menus.SetMenuChecked "snap_angle_45", newState
            
        Case pdst_Angle30
            If (Not forceInsteadOfToggle) Then newState = Not Snap.GetSnap_Angle30()
            Snap.SetSnap_Angle30 newState
            UserPrefs.SetPref_Boolean "Interface", "snap-angle-30", newState
            Menus.SetMenuChecked "snap_angle_30", newState
            
            
    End Select
    
End Sub

'Given a source angle, snap it using current snap settings.
' RETURNS: snapped angle, as relevant.  Original angle if angle snapping is turned OFF.
Public Function SnapAngle(ByVal srcAngle As Single) As Single
    
    SnapAngle = srcAngle
    If (Not Snap.GetSnap_Any()) Then Exit Function
    
    If (Snap.GetSnap_Angle30()) Then
        If (Abs(PDMath.Modulo(srcAngle, 30#)) < m_SnapDegrees) Then
            SnapAngle = Int(srcAngle / 30!) * 30!
        ElseIf (Abs(PDMath.Modulo(srcAngle, 30#)) > (30 - m_SnapDegrees)) Then
            SnapAngle = Int((srcAngle + m_SnapDegrees) / 30!) * 30!
        End If
    End If
    
    If (Snap.GetSnap_Angle45()) Then
        If (Abs(PDMath.Modulo(srcAngle, 45#)) < m_SnapDegrees) Then
            SnapAngle = Int(srcAngle / 45!) * 45!
        ElseIf (Abs(PDMath.Modulo(srcAngle, 45#)) > (45! - m_SnapDegrees)) Then
            SnapAngle = Int((srcAngle + m_SnapDegrees) / 45!) * 45!
        End If
    End If
    
    If (Snap.GetSnap_Angle90()) Then
        If (Abs(PDMath.Modulo(srcAngle, 90#)) < m_SnapDegrees) Then
            SnapAngle = Int(srcAngle / 90!) * 90!
        ElseIf (Abs(PDMath.Modulo(srcAngle, 90#)) > (90! - m_SnapDegrees)) Then
            SnapAngle = Int((srcAngle + m_SnapDegrees) / 90!) * 90!
        End If
    End If
    
End Function

'Given a source angle, snap it using some arbitrary snap threshold.  This is used by the Move/Size tool
' when rotating and the SHIFT key is held down.  (The code for this behavior seemed to fit here as much
' as anywhere, despite it not being tied to the View > Snap To menu...)
'
'Because this function ignores broader Snap preferences, it is up to the user to ensure its behavior
' is what they want *BEFORE* calling this function.
'
' RETURNS: source angle, snapped to the nearest increment of snapTarget.
Public Function SnapAngle_Arbitrary(ByVal srcAngle As Single, ByVal snapTarget As Single) As Single
    
    Dim halfAngle As Single
    halfAngle = snapTarget * 0.5!
    
    If (Abs(PDMath.Modulo(srcAngle, snapTarget)) < halfAngle) Then
        SnapAngle_Arbitrary = Int(srcAngle / snapTarget) * snapTarget
    Else
        SnapAngle_Arbitrary = Int((srcAngle + halfAngle) / snapTarget) * snapTarget
    End If
    
End Function

'Snap the passed point to any relevant snap targets (based on the user's current snap settings).
Public Sub SnapPointByMoving(ByRef srcPointF As PointFloat, ByRef dstPointF As PointFloat)
    
    m_SnappedX = False
    m_SnappedY = False
    
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
    If (minDistX < snapThreshold) Then
        m_SnappedX = True
        m_SnappedXName = xSnaps(idxSmallestX).cSnapName
        dstPointF.x = xSnaps(idxSmallestX).cValue
    End If
    
    If (minDistY < snapThreshold) Then
        m_SnappedY = True
        m_SnappedYName = ySnaps(idxSmallestY).cSnapName
        dstPointF.y = ySnaps(idxSmallestY).cValue
    End If
    
    'With all snaps calculated, we now have enough to generate smart guide coordinates
    If m_SnappedX Then BuildSmartGuideLine_X xSnaps, numXSnaps, idxSmallestX, dstPointF.x, dstPointF.y
    If m_SnappedY Then BuildSmartGuideLine_Y ySnaps, numYSnaps, idxSmallestY, dstPointF.x, dstPointF.y
    
    'Debug only
    If DEBUG_SNAP_REPORTING Then
        If m_SnappedX Then PDDebug.LogAction "Successfully snapped x to " & m_SnappedXName
        If m_SnappedY Then PDDebug.LogAction "Successfully snapped y to " & m_SnappedYName
    End If
    
End Sub

'Given a list of points, compare each to all snap points and find the *best* match among them.  Based on that,
' return the x/y offset required to move the best-match point onto the snap target.
Public Sub SnapPointListByMoving(ByRef srcPoints() As PointFloat, ByVal numOfPoints As Long, ByRef dstOffsetX As Long, ByRef dstOffsetY As Long)
    
    dstOffsetX = 0
    dstOffsetY = 0
    
    m_SnappedX = False
    m_SnappedY = False
    
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
    Dim cx As Double, cy As Double
    If Snap.GetSnap_Centerline() Then
        
        Dim pathTest As pd2DPath, pathRect As RectF
        Set pathTest = New pd2DPath
        pathTest.AddLines numOfPoints, VarPtr(srcPoints(0))
        pathRect = pathTest.GetPathBoundariesF()
        
        cx = pathRect.Left + pathRect.Width * 0.5
        cy = pathRect.Top + pathRect.Height * 0.5
        
        For i = 0 To numXSnaps - 1
            With xSnaps(i)
                .cDistanceCX = PDMath.DistanceOneDimension(cx, .cValue)
                If (.cDistanceCX < minDistX) Then
                    minDistX = .cDistanceCX
                    idxSmallestX = i
                    .cCenterComparison = True
                End If
            End With
        Next i
        
        For i = 0 To numYSnaps - 1
            With ySnaps(i)
                .cDistanceCY = PDMath.DistanceOneDimension(cy, .cValue)
                If (.cDistanceCY < minDistY) Then
                    minDistY = .cDistanceCY
                    idxSmallestY = i
                    .cCenterComparison = True
                End If
            End With
        Next i
        
    End If
        
    'Determine the minimum snap distance required for this zoom value.
    Dim snapThreshold As Double
    snapThreshold = GetSnapDistanceScaledForZoom()
    
    'If the minimum value falls beneath the minimum snap distance, snap away!
    If (minDistX < snapThreshold) Then
        
        m_SnappedX = True
        m_SnappedXName = xSnaps(idxSmallestX).cSnapName
        
        'Center comparisons require us to align the center point of the rect
        If xSnaps(idxSmallestX).cCenterComparison Then
            dstOffsetX = xSnaps(idxSmallestX).cValue - (pathRect.Left + pathRect.Width * 0.5)
        Else
            dstOffsetX = (xSnaps(idxSmallestX).cValue - srcPoints(idxSmallestPointX).x)
        End If
        
    End If
    
    If (minDistY < snapThreshold) Then
        
        m_SnappedY = True
        m_SnappedYName = ySnaps(idxSmallestY).cSnapName
        
        'Center comparisons require us to align the center point of the rect
        If ySnaps(idxSmallestY).cCenterComparison Then
            dstOffsetY = ySnaps(idxSmallestY).cValue - (pathRect.Top + pathRect.Height * 0.5)
        Else
            dstOffsetY = (ySnaps(idxSmallestY).cValue - srcPoints(idxSmallestPointY).y)
        End If
        
    End If
    
    'After both rects have been snapped, we can generate some smart guides for the viewport renderer
    If m_SnappedX Then
    
        'Construct a final smart guideline for the viewport renderer
        If xSnaps(idxSmallestX).cCenterComparison Then
            BuildSmartGuideLine_X xSnaps, numXSnaps, idxSmallestX, xSnaps(idxSmallestX).cValue, cy + dstOffsetY
        Else
        
            BuildSmartGuideLine_X xSnaps, numXSnaps, idxSmallestX, xSnaps(idxSmallestX).cValue, srcPoints(idxSmallestPointX).y + dstOffsetY
        
            'Append any other points in the source list to the line, if they also fall beneath the snap threshold
            For i = 0 To numOfPoints - 1
                If (PDMath.DistanceOneDimension(srcPoints(i).x, xSnaps(idxSmallestX).cValue) < snapThreshold) Then AppendPointToSmartGuideLine_X xSnaps(idxSmallestX).cValue, srcPoints(i).y + dstOffsetY
            Next i
            
        End If
        
    End If
    
    If m_SnappedY Then
    
        'Construct a final smart guideline for the viewport renderer
        If ySnaps(idxSmallestY).cCenterComparison Then
            BuildSmartGuideLine_Y ySnaps, numYSnaps, idxSmallestY, cx + dstOffsetX, ySnaps(idxSmallestY).cValue
        Else
            BuildSmartGuideLine_Y ySnaps, numYSnaps, idxSmallestY, srcPoints(idxSmallestPointY).x + dstOffsetX, ySnaps(idxSmallestY).cValue
            
            'Append any other points in the source list to the line, if they also fall beneath the snap threshold
            For i = 0 To numOfPoints - 1
                If (PDMath.DistanceOneDimension(srcPoints(i).y, ySnaps(idxSmallestY).cValue) < snapThreshold) Then AppendPointToSmartGuideLine_Y srcPoints(i).x + dstOffsetX, ySnaps(idxSmallestY).cValue
            Next i
            
        End If
        
    End If
    
    'Debug only
    If DEBUG_SNAP_REPORTING Then
        If m_SnappedX Then PDDebug.LogAction "Successfully snapped x to " & m_SnappedXName
        If m_SnappedY Then PDDebug.LogAction "Successfully snapped y to " & m_SnappedYName
    End If
    
End Sub

'Snap the passed rect to any relevant snap targets (based on the user's current snap settings).
' Because this function snaps only by moving the target rect, it is guaranteed that only the
' top and left values will be changed by the function (width/height will *not*).
Public Sub SnapRectByMoving(ByRef srcRectF As RectF, ByRef dstRectF As RectF)
    
    m_SnappedX = False
    m_SnappedY = False
    
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
    If Snap.GetSnap_Centerline() Then
        
        Dim cx As Double, cy As Double
        cx = compareRectF.Left + (compareRectF.Right - compareRectF.Left) * 0.5
        cy = compareRectF.Top + (compareRectF.Bottom - compareRectF.Top) * 0.5
        
        For i = 0 To numXSnaps - 1
            With xSnaps(i)
                .cDistanceCX = PDMath.DistanceOneDimension(cx, .cValue)
                If (.cDistanceCX < minDistX) Then
                    minDistX = .cDistanceCX
                    idxSmallestX = i
                    .cCenterComparison = True
                End If
            End With
        Next i
        
        For i = 0 To numYSnaps - 1
            With ySnaps(i)
                .cDistanceCY = PDMath.DistanceOneDimension(cy, .cValue)
                If (.cDistanceCY < minDistY) Then
                    minDistY = .cDistanceCY
                    idxSmallestY = i
                    .cCenterComparison = True
                End If
            End With
        Next i
        
    End If
        
    'Determine the minimum snap distance required for this zoom value.
    Dim snapThreshold As Double
    snapThreshold = GetSnapDistanceScaledForZoom()
    
    'If the minimum value falls beneath the minimum snap distance, snap away!
    If (minDistX < snapThreshold) Then
        
        m_SnappedX = True
        m_SnappedXName = xSnaps(idxSmallestX).cSnapName
        
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
        
        'Construct a final smart guideline for the viewport renderer (and append all four rect points,
        ' in case they also lie on this line
        BuildSmartGuideLine_X xSnaps, numXSnaps, idxSmallestX, dstRectF.Left, dstRectF.Top
        AppendPointToSmartGuideLine_X dstRectF.Left + dstRectF.Width, dstRectF.Top
        AppendPointToSmartGuideLine_X dstRectF.Left, dstRectF.Top + dstRectF.Height
        AppendPointToSmartGuideLine_X dstRectF.Left + dstRectF.Width, dstRectF.Top + dstRectF.Height
        
    End If
    
    If (minDistY < snapThreshold) Then
        
        m_SnappedY = True
        m_SnappedYName = ySnaps(idxSmallestY).cSnapName
        
        If ySnaps(idxSmallestY).cCenterComparison Then
            dstRectF.Top = ySnaps(idxSmallestY).cValue - (compareRectF.Bottom - compareRectF.Top) * 0.5
        Else
            If (ySnaps(idxSmallestY).cDistance1 < ySnaps(idxSmallestY).cDistance2) Then
                dstRectF.Top = ySnaps(idxSmallestY).cValue
            Else
                dstRectF.Top = ySnaps(idxSmallestY).cValue - dstRectF.Height
            End If
        End If
        
        'Construct a final smart guideline for the viewport renderer (and append the bottom of the rect too)
        BuildSmartGuideLine_Y ySnaps, numYSnaps, idxSmallestY, dstRectF.Left, dstRectF.Top
        AppendPointToSmartGuideLine_Y dstRectF.Left + dstRectF.Width, dstRectF.Top
        AppendPointToSmartGuideLine_Y dstRectF.Left, dstRectF.Top + dstRectF.Height
        AppendPointToSmartGuideLine_Y dstRectF.Left + dstRectF.Width, dstRectF.Top + dstRectF.Height
        
    End If
    
    'Debug only
    If DEBUG_SNAP_REPORTING Then
        If m_SnappedX Then PDDebug.LogAction "Successfully snapped x to " & m_SnappedXName
        If m_SnappedY Then PDDebug.LogAction "Successfully snapped y to " & m_SnappedYName
    End If
    
End Sub

'From the snapped point, build a list of all snap targets that lie on the same line.
Private Sub BuildSmartGuideLine_X(ByRef srcSnaps() As SnapComparison, ByVal numSnaps As Long, ByVal idxSnapped As Long, ByVal srcX As Single, ByVal srcY As Single)
    
    'Set the initial line endpoints, and ensure that the top-most line is point 1
    If ((srcSnaps(idxSnapped).cSnapSource And pdst_Centerline) <> 0) Then
        m_SnappedXPt1 = srcSnaps(idxSnapped).cSnapLinePt1
        m_SnappedXPt2 = srcSnaps(idxSnapped).cSnapLinePt1
    Else
        If (srcSnaps(idxSnapped).cSnapLinePt1.y < srcSnaps(idxSnapped).cSnapLinePt2.y) Then
            m_SnappedXPt1 = srcSnaps(idxSnapped).cSnapLinePt1
            m_SnappedXPt2 = srcSnaps(idxSnapped).cSnapLinePt2
        Else
            m_SnappedXPt1 = srcSnaps(idxSnapped).cSnapLinePt2
            m_SnappedXPt2 = srcSnaps(idxSnapped).cSnapLinePt1
        End If
    End If
    
    'Extend the line to include any other snap targets that also lie on this line
    Dim i As Long
    For i = 0 To numSnaps - 1
        If (i <> idxSnapped) Then
            If (srcSnaps(i).cValue = srcSnaps(idxSnapped).cValue) Then
                AppendPointToSmartGuideLine_X srcSnaps(i).cSnapLinePt1.x, srcSnaps(i).cSnapLinePt1.y
                AppendPointToSmartGuideLine_X srcSnaps(i).cSnapLinePt2.x, srcSnaps(i).cSnapLinePt2.y
            End If
        End If
    Next i
    
    'Also append the passed point
    AppendPointToSmartGuideLine_X srcX, srcY
    
    'The smart guide line is now inclusive of all snap targets lying on the same line.  (This way, if multiple layers
    ' or objects share the same boundary line, we include *all* of them in the smart guide, instead of the one arbitrary
    ' one that was chosen as the "closest" during point comparisons.)
    
End Sub

'Append snapped points, if valid, to our constructed "smart guide line"
Private Sub AppendPointToSmartGuideLine_X(ByVal srcX As Single, ByVal srcY As Single)
    
    'Ensure same x-values
    If (srcX = m_SnappedXPt1.x) Then
        
        'See if this point lies beyond the existing line
        If (srcY < m_SnappedXPt1.y) Then m_SnappedXPt1.y = srcY
        If (srcY > m_SnappedXPt2.y) Then m_SnappedXPt2.y = srcY
        
    End If
    
End Sub

'From the snapped point, build a list of all snap targets that lie on the same line.
Private Sub BuildSmartGuideLine_Y(ByRef srcSnaps() As SnapComparison, ByVal numSnaps As Long, ByVal idxSnapped As Long, ByVal srcX As Single, ByVal srcY As Single)

    'Set the initial line endpoints, and ensure that the top-most line is point 1
    If ((srcSnaps(idxSnapped).cSnapSource And pdst_Centerline) <> 0) Then
        m_SnappedYPt1 = srcSnaps(idxSnapped).cSnapLinePt1
        m_SnappedYPt2 = srcSnaps(idxSnapped).cSnapLinePt1
    Else
        If (srcSnaps(idxSnapped).cSnapLinePt1.x < srcSnaps(idxSnapped).cSnapLinePt2.x) Then
            m_SnappedYPt1 = srcSnaps(idxSnapped).cSnapLinePt1
            m_SnappedYPt2 = srcSnaps(idxSnapped).cSnapLinePt2
        Else
            m_SnappedYPt1 = srcSnaps(idxSnapped).cSnapLinePt2
            m_SnappedYPt2 = srcSnaps(idxSnapped).cSnapLinePt1
        End If
    End If
        
    'Extend the line to include any other snap targets that also lie on this line
    Dim i As Long
    For i = 0 To numSnaps - 1
        If (i <> idxSnapped) Then
            If (srcSnaps(i).cValue = srcSnaps(idxSnapped).cValue) Then
                AppendPointToSmartGuideLine_Y srcSnaps(i).cSnapLinePt1.x, srcSnaps(i).cSnapLinePt1.y
                AppendPointToSmartGuideLine_Y srcSnaps(i).cSnapLinePt2.x, srcSnaps(i).cSnapLinePt2.y
            End If
        End If
    Next i
    
    'Also append the passed point
    AppendPointToSmartGuideLine_Y srcX, srcY
    
    'The smart guide line is now inclusive of all snap targets lying on the same line.  (This way, if multiple layers
    ' or objects share the same boundary line, we include *all* of them in the smart guide, instead of the one arbitrary
    ' one that was chosen as the "closest" during point comparisons.)
    
End Sub

'Append snapped points, if valid, to our constructed "smart guide line"
Private Sub AppendPointToSmartGuideLine_Y(ByVal srcX As Single, ByVal srcY As Single)
    
    'Ensure same y-values
    If (srcY = m_SnappedYPt1.y) Then
        
        'See if this point lies beyond the existing line
        If (srcX < m_SnappedYPt1.x) Then m_SnappedYPt1.x = srcX
        If (srcX > m_SnappedYPt2.x) Then m_SnappedYPt2.x = srcX
        
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
        
        'Ensure space is available in the target array
        If (UBound(dstSnaps) < GetSnapTargets_X + 1) Then ReDim Preserve dstSnaps(0 To GetSnapTargets_X * 2 - 1) As SnapComparison
        
        'Add canvas boundaries to the snap list
        With dstSnaps(GetSnapTargets_X)
            .cValue = 0#
            .cSnapSource = pdst_CanvasEdge
            .cSnapLinePt1.x = 0!
            .cSnapLinePt1.y = 0!
            .cSnapLinePt2.x = 0!
            .cSnapLinePt2.y = PDImages.GetActiveImage.Height
            .cSnapName = "canvas left"
        End With
        
        With dstSnaps(GetSnapTargets_X + 1)
            .cValue = PDImages.GetActiveImage.Width
            .cSnapSource = pdst_CanvasEdge
            .cSnapLinePt1.x = PDImages.GetActiveImage.Width
            .cSnapLinePt1.y = 0!
            .cSnapLinePt2.x = PDImages.GetActiveImage.Width
            .cSnapLinePt2.y = PDImages.GetActiveImage.Height
            .cSnapName = "canvas right"
        End With
        
        GetSnapTargets_X = GetSnapTargets_X + 2
        
        'Centerline (of canvas only; layer centers are handled below)
        If Snap.GetSnap_Centerline() Then
            
            If (UBound(dstSnaps) < GetSnapTargets_X) Then ReDim Preserve dstSnaps(0 To GetSnapTargets_X * 2 - 1) As SnapComparison
            With dstSnaps(GetSnapTargets_X)
                .cValue = Int(PDImages.GetActiveImage.Width / 2)
                .cSnapSource = pdst_CanvasEdge Or pdst_Centerline
                .cSnapLinePt1.x = .cValue
                .cSnapLinePt1.y = Int(PDImages.GetActiveImage.Height / 2)
                .cSnapName = "canvas center"
            End With
            
            GetSnapTargets_X = GetSnapTargets_X + 1
            
        End If
            
    End If
    
    'Layer boundaries next
    If Snap.GetSnap_Layer() Then
        
        Dim layerRectF As RectF
        
        Dim i As Long
        For i = 0 To PDImages.GetActiveImage.GetNumOfLayers - 1
            
            'Do *not* snap the active layer (or it will always snap to itself because that's what's closest, lol).
            ' (This is only relevant for layer-centric tools, like the move/size tool.)
            Dim skipActiveLayer As Boolean
            skipActiveLayer = (g_CurrentTool = NAV_MOVE) And (i = PDImages.GetActiveImage.GetActiveLayerIndex)
            If (Not skipActiveLayer) Then
                
                'Ignore invisible layers
                If PDImages.GetActiveImage.GetActiveLayer.GetLayerVisibility() Then
                
                    'Ensure space is available in the target array
                    If (UBound(dstSnaps) < GetSnapTargets_X + 2) Then ReDim Preserve dstSnaps(0 To GetSnapTargets_X * 2 - 1) As SnapComparison
                    
                    'Add layer boundaries to the snap list
                    PDImages.GetActiveImage.GetLayerByIndex(i).GetLayerBoundaryRect layerRectF
                    With dstSnaps(GetSnapTargets_X)
                        .cValue = layerRectF.Left
                        .cSnapSource = pdst_Layer
                        .cSnapLinePt1.x = .cValue
                        .cSnapLinePt1.y = layerRectF.Top
                        .cSnapLinePt2.x = .cValue
                        .cSnapLinePt2.y = layerRectF.Top + layerRectF.Height
                        .cSnapName = PDImages.GetActiveImage.GetLayerByIndex(i).GetLayerName & " left"
                    End With
                    
                    With dstSnaps(GetSnapTargets_X + 1)
                        .cValue = layerRectF.Left + layerRectF.Width
                        .cSnapSource = pdst_Layer
                        .cSnapLinePt1.x = .cValue
                        .cSnapLinePt1.y = layerRectF.Top
                        .cSnapLinePt2.x = .cValue
                        .cSnapLinePt2.y = layerRectF.Top + layerRectF.Height
                        .cSnapName = PDImages.GetActiveImage.GetLayerByIndex(i).GetLayerName & " right"
                    End With
                    GetSnapTargets_X = GetSnapTargets_X + 2
                    
                    'If centerlines are enabled, add the layer's centerpoint too
                    If Snap.GetSnap_Centerline() Then
                        With dstSnaps(GetSnapTargets_X)
                            .cValue = Int(layerRectF.Left + layerRectF.Width * 0.5)
                            .cSnapSource = pdst_Layer Or pdst_Centerline
                            .cSnapLinePt1.x = .cValue
                            .cSnapLinePt1.y = Int(layerRectF.Top + layerRectF.Height * 0.5)
                            .cSnapName = PDImages.GetActiveImage.GetLayerByIndex(i).GetLayerName & " center"
                        End With
                        GetSnapTargets_X = GetSnapTargets_X + 1
                    End If
                    
                End If
                
            End If
                
        Next i
        
    End If
    
    'TODO: more snap targets in the future...?
    
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
        With dstSnaps(GetSnapTargets_Y)
            .cValue = 0#
            .cSnapSource = pdst_CanvasEdge
            .cSnapLinePt1.x = 0!
            .cSnapLinePt1.y = 0!
            .cSnapLinePt2.x = PDImages.GetActiveImage.Width
            .cSnapLinePt2.y = 0!
            .cSnapName = "canvas top"
        End With
        
        With dstSnaps(GetSnapTargets_Y + 1)
            .cValue = PDImages.GetActiveImage.Height
            .cSnapSource = pdst_CanvasEdge
            .cSnapLinePt1.x = 0!
            .cSnapLinePt1.y = PDImages.GetActiveImage.Height
            .cSnapLinePt2.x = PDImages.GetActiveImage.Width
            .cSnapLinePt2.y = PDImages.GetActiveImage.Height
            .cSnapName = "canvas bottom"
        End With
        
        GetSnapTargets_Y = GetSnapTargets_Y + 2
        
        'Centerline (of canvas only; layers is handled below)
        If Snap.GetSnap_Centerline() Then
            
            If (UBound(dstSnaps) < GetSnapTargets_Y) Then ReDim Preserve dstSnaps(0 To GetSnapTargets_Y * 2 - 1) As SnapComparison
            
            With dstSnaps(GetSnapTargets_Y)
                .cValue = Int(PDImages.GetActiveImage.Height / 2)
                .cSnapSource = pdst_CanvasEdge Or pdst_Centerline
                .cSnapLinePt1.x = Int(PDImages.GetActiveImage.Width / 2)
                .cSnapLinePt1.y = .cValue
                .cSnapName = "canvas center"
            End With
            
            GetSnapTargets_Y = GetSnapTargets_Y + 1
            
        End If
        
    End If
    
    'Layer boundaries next
    If Snap.GetSnap_Layer() Then
        
        Dim layerRectF As RectF
        
        Dim i As Long
        For i = 0 To PDImages.GetActiveImage.GetNumOfLayers - 1
            
            'Do *not* snap the active layer (or it will always snap to itself because that's what's closest, lol)
            If (i <> PDImages.GetActiveImage.GetActiveLayerIndex) Then
                
                'Ignore invisible layers
                If PDImages.GetActiveImage.GetActiveLayer.GetLayerVisibility() Then
                
                    'Ensure space is available in the target array
                    If (UBound(dstSnaps) < GetSnapTargets_Y + 2) Then ReDim Preserve dstSnaps(0 To GetSnapTargets_Y * 2 - 1) As SnapComparison
                    
                    'Add layer boundaries to the snap list
                    PDImages.GetActiveImage.GetLayerByIndex(i).GetLayerBoundaryRect layerRectF
                    
                    With dstSnaps(GetSnapTargets_Y)
                        .cValue = layerRectF.Top
                        .cSnapSource = pdst_Layer
                        .cSnapLinePt1.x = layerRectF.Left
                        .cSnapLinePt1.y = .cValue
                        .cSnapLinePt2.x = layerRectF.Left + layerRectF.Width
                        .cSnapLinePt2.y = .cValue
                        .cSnapName = PDImages.GetActiveImage.GetLayerByIndex(i).GetLayerName & " top"
                    End With
                    
                    With dstSnaps(GetSnapTargets_Y + 1)
                        .cValue = layerRectF.Top + layerRectF.Height
                        .cSnapSource = pdst_Layer
                        .cSnapLinePt1.x = layerRectF.Left
                        .cSnapLinePt1.y = .cValue
                        .cSnapLinePt2.x = layerRectF.Left + layerRectF.Width
                        .cSnapLinePt2.y = .cValue
                        .cSnapName = PDImages.GetActiveImage.GetLayerByIndex(i).GetLayerName & " bottom"
                    End With
                    
                    GetSnapTargets_Y = GetSnapTargets_Y + 2
                    
                    'If centerlines are enabled, add the layer's centerline too
                    If Snap.GetSnap_Centerline() Then
                        With dstSnaps(GetSnapTargets_Y)
                            .cValue = Int(layerRectF.Top + layerRectF.Height * 0.5)
                            .cSnapSource = pdst_Layer Or pdst_Centerline
                            .cSnapLinePt1.x = Int(layerRectF.Left + layerRectF.Width * 0.5)
                            .cSnapLinePt1.y = .cValue
                            .cSnapName = PDImages.GetActiveImage.GetLayerByIndex(i).GetLayerName & " center"
                        End With
                        GetSnapTargets_Y = GetSnapTargets_Y + 1
                    End If
                    
                End If
                
            End If
                
        Next i
        
    End If
    
    'TODO: more snap targets in the future...?
    
End Function
