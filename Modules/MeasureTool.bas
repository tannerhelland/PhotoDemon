Attribute VB_Name = "Tools_Measure"
'***************************************************************************
'Measure tool interface
'Copyright 2018-2026 by Tanner Helland
'Created: 11/July/17
'Last updated: 28/May/14
'Last update: add support for percent as a measurement unit
'
'PD's Measure tool is very straightforward.  Additional details can be found in the associated form
' (toolpanel_Measure).
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'First and second mouse coordinates, and their respective "set" status (e.g. the user has created them)
Private m_Points() As PointFloat, m_PointsSet() As Boolean

'If at least one point has been set, the measurement status will be set to "ready"; this means we can
' perform calculations without crashing anything.
Private m_MeasurementReady As Boolean

'If the user is currently clicking on a mouse point (e.g. click-dragging to move a point), this will
' be set to the relevant index.
Private m_ActivePointIndex As Long, m_HoverPointIndex As Long

'Current mouse position, if any - note that these will be in IMAGE coordinates, by convention
Private m_MouseX As Single, m_MouseY As Single

Public Sub InitializeMeasureTool()
    
    'Initialize the measurement points to some arbitrarily out-of-range value
    ReDim m_Points(0 To 1) As PointFloat
    m_Points(0).x = SINGLE_MAX
    m_Points(0).y = SINGLE_MAX
    m_Points(1).x = SINGLE_MAX
    m_Points(1).y = SINGLE_MAX
    
    'Initialize all other trackers to FALSE; these won't be set until the user interacts with the canvas
    ReDim m_PointsSet(0 To 1) As Boolean
    m_MeasurementReady = False
    m_ActivePointIndex = -1
    m_HoverPointIndex = -1
    
End Sub

Public Sub NotifyMouseDown(ByRef srcCanvas As pdCanvas, ByVal imgX As Single, ByVal imgY As Single)
    
    m_ActivePointIndex = -1
    
    'Failsafe checks for a valid source image
    m_MeasurementReady = PDImages.IsImageActive()
    
    If m_MeasurementReady Then
        
        Dim activeIndex As Long
        activeIndex = -1
        
        'If a previous point has been set, see if the user is re-clicking it
        If (m_PointsSet(0) Or m_PointsSet(1)) Then
            
            activeIndex = IsMouseOverPoint(imgX, imgY)
        
            'If the user is *not* moving an existing point, we want to redraw the entire measurement line;
            ' this is automatically handled by the *next* block
            If (activeIndex >= 0) Then
                m_ActivePointIndex = activeIndex
                m_Points(activeIndex).x = Int(imgX + 0.5!)
                m_Points(activeIndex).y = Int(imgY + 0.5!)
            End If
        
        End If
        
        'If this is the first time the user is using the tool, or if the user is click-drawing on
        ' a new region of the image (separate from any existing points), set both points to match
        ' the initial input point.
        If (activeIndex = -1) Then
            m_Points(0).x = Int(imgX + 0.5!)
            m_Points(0).y = Int(imgY + 0.5!)
            m_Points(1).x = Int(imgX + 0.5!)
            m_Points(1).y = Int(imgY + 0.5!)
            m_PointsSet(0) = True
            m_PointsSet(1) = True
            m_ActivePointIndex = 1
        End If
        
    End If
    
    'Update the display as necessary
    If m_MeasurementReady Then toolpanel_Measure.UpdateUIText
    
End Sub

Public Sub NotifyMouseMove(ByVal lmbDown As Boolean, ByVal Shift As ShiftConstants, ByVal imgX As Single, ByVal imgY As Single)

    If (Not m_MeasurementReady) Then Exit Sub
    
    'Update the current mouse position trackers
    m_MouseX = imgX
    m_MouseY = imgY
    
    'If the left mouse button is down, move the active point to the specified location
    If lmbDown Then
    
        m_HoverPointIndex = m_ActivePointIndex
        m_Points(m_ActivePointIndex).x = Int(imgX + 0.5!)
        m_Points(m_ActivePointIndex).y = Int(imgY + 0.5!)
    
    'If the left mouse button is *not* down, simply update the current "hover point" index
    Else
        m_HoverPointIndex = IsMouseOverPoint(imgX, imgY)
    End If
    
    'Notify the UI of the new measurements, so it can update its measurement values accordingly
    toolpanel_Measure.UpdateUIText

End Sub

Public Sub NotifyMouseUp(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal imgX As Single, ByVal imgY As Single, ByVal numOfMouseMovements As Long, ByVal clickEventAlsoFiring As Boolean)
    
    If (Not m_MeasurementReady) Then Exit Sub
    
    'If the user has just clicked (e.g. this was not a click-drag action), remove the active measurement
    If clickEventAlsoFiring Then
        InitializeMeasureTool
        toolpanel_Measure.UpdateUIText
        Viewport.Stage4_FlipBufferAndDrawUI PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    Else
    
        'Update the final position, if any
        If (m_ActivePointIndex >= 0) Then
            m_Points(m_ActivePointIndex).x = Int(imgX + 0.5!)
            m_Points(m_ActivePointIndex).y = Int(imgY + 0.5!)
        End If
        
        'Note that no points are currently "active"
        m_ActivePointIndex = -1
        
        'Update the hover index
        m_HoverPointIndex = IsMouseOverPoint(imgX, imgY)
        
    End If
    
End Sub

Public Function ArePointsReady() As Boolean
    If (Not m_MeasurementReady) Then Exit Function
    If (Not m_PointsSet(0)) Or (Not m_PointsSet(1)) Then Exit Function
    ArePointsReady = True
End Function

Public Function GetFirstPoint(ByRef dstPoint As PointFloat) As Boolean
    If (Not m_MeasurementReady) Then Exit Function
    If (Not m_PointsSet(0)) Then Exit Function
    dstPoint = m_Points(0)
    GetFirstPoint = True
End Function

Public Function GetSecondPoint(ByRef dstPoint As PointFloat) As Boolean
    If (Not m_MeasurementReady) Then Exit Function
    If (Not m_PointsSet(1)) Then Exit Function
    dstPoint = m_Points(1)
    GetSecondPoint = True
End Function

Public Sub RequestRedraw()

    If (Not m_MeasurementReady) Then Exit Sub
    If (Not m_PointsSet(0)) Then Exit Sub
    
    Viewport.Stage4_FlipBufferAndDrawUI PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    toolpanel_Measure.UpdateUIText
    
End Sub

Public Sub SetPointsManually(ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single)
    m_PointsSet(0) = True
    m_PointsSet(1) = True
    m_Points(0).x = x1
    m_Points(0).y = y1
    m_Points(1).x = x2
    m_Points(1).y = y2
End Sub

Public Function SpecialCursorWanted() As Boolean
    SpecialCursorWanted = (m_ActivePointIndex >= 0) Or (m_HoverPointIndex >= 0)
End Function

'Swap the anchor and secondary points
Public Sub SwapPoints()
    
    If (Not m_MeasurementReady) Then Exit Sub
    If (Not m_PointsSet(0)) Then Exit Sub
    
    Dim tmpPoint As PointFloat
    tmpPoint = m_Points(0)
    m_Points(0) = m_Points(1)
    m_Points(1) = tmpPoint

    Viewport.Stage4_FlipBufferAndDrawUI PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    toolpanel_Measure.UpdateUIText
    
End Sub

'Returns: the index of which point the mouse is over, if any; -1 if the mouse is *not* over a point
Private Function IsMouseOverPoint(ByVal chkX As Single, ByVal chkY As Single) As Long
    
    'MouseAccuracy in PD is a global value, but because we are working in image coordinates, we must compensate for the
    ' current zoom value.  (Otherwise, when zoomed out the user would be forced to work with tighter accuracy!)
    ' (TODO: come up with a better solution for this.  Accuracy should *really* be handled in the canvas coordinate space,
    '        so perhaps the caller should specify an image x/y and a radius...?)
    Dim mouseAccuracy As Double
    mouseAccuracy = Drawing.ConvertCanvasSizeToImageSize(Interface.GetStandardInteractionDistance(), PDImages.GetActiveImage)
    
    IsMouseOverPoint = -1
    
    Dim curDistance As Single, minDistance As Single, minIndex As Long
    minDistance = SINGLE_MAX
    
    Dim i As Long
    For i = 0 To 1
        curDistance = PDMath.DistanceTwoPoints(chkX, chkY, m_Points(i).x, m_Points(i).y)
        If (curDistance < minDistance) Then
            minDistance = curDistance
            minIndex = i
        End If
    Next i
    
    If (minDistance < mouseAccuracy) Then IsMouseOverPoint = minIndex

End Function

'Returns: TRUE if the angle exists and is valid; FALSE otherwise
Public Function GetAngleInDegrees(ByRef dstAngle As Double) As Boolean
    If (Not m_MeasurementReady) Then Exit Function
    If (Not m_PointsSet(0)) Or (Not m_PointsSet(1)) Then Exit Function
    dstAngle = PDMath.Atan2(m_Points(1).y - m_Points(0).y, m_Points(1).x - m_Points(0).x)
    dstAngle = PDMath.RadiansToDegrees(dstAngle)
    GetAngleInDegrees = True
End Function

'Returns: TRUE if both points exist and are valid; FALSE otherwise
Public Function GetDistanceInPx(ByRef dstDistance As Double) As Boolean
    If (Not m_MeasurementReady) Then Exit Function
    If (Not m_PointsSet(0)) Or (Not m_PointsSet(1)) Then Exit Function
    dstDistance = PDMath.DistanceTwoPoints(m_Points(0).x, m_Points(0).y, m_Points(1).x, m_Points(1).y)
    GetDistanceInPx = True
End Function

'When something external changes the current measurement unit, call this sub; it will take care of
' refreshing all necessary UI bits.
Public Sub NotifyUnitChange()
    toolpanel_Measure.UpdateUIText
End Sub

'Render the current measurement UI onto the specified canvas
Public Sub RenderMeasureUI(ByRef targetCanvas As pdCanvas)
    
    'If measurements aren't ready, bail
    If (Not m_MeasurementReady) Then Exit Sub
    
    'Update our internal angle calculation, if any
    Dim curAngle As Double
    If (Not GetAngleInDegrees(curAngle)) Then Exit Sub
    
    'Start by converting the measurement positions to canvas coordinates
    Dim canvasCoordsX(0 To 1) As Double, canvasCoordsY(0 To 1) As Double
    
    Dim i As Long
    For i = 0 To 1
        Drawing.ConvertImageCoordsToCanvasCoords targetCanvas, PDImages.GetActiveImage(), m_Points(i).x, m_Points(i).y, canvasCoordsX(i), canvasCoordsY(i)
    Next i
    
    'Clone a pair of UI pens from the main rendering module.  (Note that we clone them because we need
    ' to modify some rendering properties, and we don't want to fuck up the central cache.)
    Dim basePenInactive As pd2DPen, topPenInactive As pd2DPen
    Dim basePenActive As pd2DPen, topPenActive As pd2DPen
    Drawing.CloneCachedUIPens basePenInactive, topPenInactive, False
    Drawing.CloneCachedUIPens basePenActive, topPenActive, True
    
    'Specify rounded line edges for our pens; this looks better for this particular tool
    basePenInactive.SetPenLineCap P2_LC_Round
    topPenInactive.SetPenLineCap P2_LC_Round
    basePenActive.SetPenLineCap P2_LC_Round
    topPenActive.SetPenLineCap P2_LC_Round
    
    basePenInactive.SetPenLineJoin P2_LJ_Round
    topPenInactive.SetPenLineJoin P2_LJ_Round
    basePenActive.SetPenLineJoin P2_LJ_Round
    topPenActive.SetPenLineJoin P2_LJ_Round
    
    'Create other required pd2D drawing tools (a surface)
    Dim cSurface As pd2DSurface
    Drawing2D.QuickCreateSurfaceFromDC cSurface, targetCanvas.hDC, True
    cSurface.SetSurfacePixelOffset P2_PO_Normal
    
    'We now want to add all lines and arcs to a path, which we'll render all at once
    Dim measurePath As pd2DPath
    Set measurePath = New pd2DPath
    
    'Add the line between the two points.  (Note the we explicitly draw it *backward*, so that it
    ' will junction nicely with the horizontal line.)
    measurePath.AddLine canvasCoordsX(1), canvasCoordsY(1), canvasCoordsX(0), canvasCoordsY(0)
    
    'Add a 0-degree "baseline" oriented horizontally in the direction most relevant to the current angle
    Dim baselineLength As Single
    baselineLength = Interface.FixDPIFloat(35)
    If (Abs(curAngle) <= 90#) Then
        measurePath.AddLine canvasCoordsX(0), canvasCoordsY(0), canvasCoordsX(0) + baselineLength, canvasCoordsY(0)
    Else
        measurePath.AddLine canvasCoordsX(0), canvasCoordsY(0), canvasCoordsX(0) - baselineLength, canvasCoordsY(0)
    End If
    
    'We now want to draw an arc demonstrating the current angle.  However, if the current measurement
    ' is extremely small, we will suspend the arc until the distance grows.
    Dim arcRadius As Single
    arcRadius = baselineLength * 0.8
    
    'Calculate the current line size *on-screen*
    Dim lineDistance As Double
    lineDistance = PDMath.DistanceTwoPoints(canvasCoordsX(0), canvasCoordsY(0), canvasCoordsX(1), canvasCoordsY(1))
    
    If (arcRadius < lineDistance) Then
    
        measurePath.StartNewFigure
        
        'Arc faces right
        If (Abs(curAngle) <= 90#) Then
            measurePath.AddArcCircular canvasCoordsX(0), canvasCoordsY(0), arcRadius, 0#, curAngle
            
        'Arc faces left
        Else
            measurePath.AddArcCircular canvasCoordsX(0), canvasCoordsY(0), arcRadius, 180#, Sgn(curAngle) * (Abs(curAngle) - 180#)
        End If
        
    End If
    
    'Stroke the path
    PD2D.DrawPath cSurface, basePenInactive, measurePath
    PD2D.DrawPath cSurface, topPenInactive, measurePath
    
    'The two measuring points need an on-screen size at which they will be drawn
    Dim circRadius As Single
    circRadius = 7!
    
    'Create a path for the "crosshair" over the secondary point
    Dim crosshairPath As pd2DPath
    Set crosshairPath = New pd2DPath
    crosshairPath.AddLine canvasCoordsX(1), canvasCoordsY(1) + circRadius, canvasCoordsX(1), canvasCoordsY(1) - circRadius
    crosshairPath.StartNewFigure
    crosshairPath.AddLine canvasCoordsX(1) + circRadius, canvasCoordsY(1), canvasCoordsX(1) - circRadius, canvasCoordsY(1)
    
    'If either point is active (or hovered), render them with the active color; otherwise, default to
    ' white-on-black.
    For i = 0 To 1
        If (m_ActivePointIndex = i) Or (m_HoverPointIndex = i) Then
            
            'The anchor point gets a circle; the secondary point gets a crosshair
            If (i = 0) Then
                PD2D.DrawCircleF cSurface, basePenActive, canvasCoordsX(i), canvasCoordsY(i), circRadius
                PD2D.DrawCircleF cSurface, topPenActive, canvasCoordsX(i), canvasCoordsY(i), circRadius
            Else
                PD2D.DrawPath cSurface, basePenActive, crosshairPath
                PD2D.DrawPath cSurface, topPenActive, crosshairPath
            End If
        Else
            
            If (i = 0) Then
                PD2D.DrawCircleF cSurface, basePenInactive, canvasCoordsX(i), canvasCoordsY(i), circRadius
                PD2D.DrawCircleF cSurface, topPenInactive, canvasCoordsX(i), canvasCoordsY(i), circRadius
            Else
                PD2D.DrawPath cSurface, basePenInactive, crosshairPath
                PD2D.DrawPath cSurface, topPenInactive, crosshairPath
            End If
            
        End If
    Next i
    
    Set cSurface = Nothing
    Set basePenInactive = Nothing: Set topPenInactive = Nothing
    Set basePenActive = Nothing: Set topPenActive = Nothing
    
End Sub

Public Sub ResetPoints(Optional ByVal alsoRedrawViewport As Boolean = True)
    InitializeMeasureTool
    toolpanel_Measure.UpdateUIText
    If alsoRedrawViewport Then Viewport.Stage4_FlipBufferAndDrawUI PDImages.GetActiveImage(), FormMain.MainCanvas(0)
End Sub

Public Sub Rotate2ndPoint90Degrees()

    If (Not m_MeasurementReady) Then Exit Sub
    If (Not m_PointsSet(0)) Then Exit Sub
    
    Dim tmpPoint As PointFloat
    PDMath.RotatePointAroundPoint m_Points(1).x, m_Points(1).y, m_Points(0).x, m_Points(0).y, PDMath.DegreesToRadians(90), tmpPoint.x, tmpPoint.y
    m_Points(1) = tmpPoint
    
    Viewport.Stage4_FlipBufferAndDrawUI PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    toolpanel_Measure.UpdateUIText
    
End Sub

'NOTE: this code is a duplicate of the FormStraighten command string generator.
' Any changes here also need to be mirrored there.
Public Sub StraightenImageToMatch()

    Dim curAngle As Double
    If Tools_Measure.GetAngleInDegrees(curAngle) Then
        
        'The straighten tool only works on angles on the range [-45, 45].  Beyond this point,
        ' the image has to be resized to an absurd degree.  Convert the image to the range
        ' [-90, 90] to start
        If (Abs(curAngle) > 90#) Then
            If (curAngle > 0#) Then curAngle = curAngle - 180# Else curAngle = curAngle + 180#
        End If
        
        'If the angle is > 45 from the horizon, assume it's measuring a vertical line (not a horizontal one).
        If (Abs(curAngle) > 45#) Then
            If (curAngle > 0#) Then curAngle = curAngle - 90# Else curAngle = curAngle + 90#
        End If
        
        Dim cParams As pdSerialize
        Set cParams = New pdSerialize
        
        With cParams
            .AddParam "angle", -1# * curAngle
            .AddParam "target", pdat_Image
        End With
        
        Process "Straighten image", , cParams.GetParamString(), UNDO_Image
        
    End If

End Sub

Public Sub StraightenLayerToMatch()

    Dim curAngle As Double
    If Tools_Measure.GetAngleInDegrees(curAngle) Then
    
        'The straighten tool only works on angles on the range [-45, 45].  Beyond this point,
        ' the image has to be resized to an absurd degree.  Convert the image to the range
        ' [-90, 90] to start
        If (Abs(curAngle) > 90#) Then
            If (curAngle > 0#) Then curAngle = curAngle - 180# Else curAngle = curAngle + 180#
        End If
        
        'If the angle is > 45 from the horizon, assume it's measuring a vertical line (not a horizontal one).
        If (Abs(curAngle) > 45#) Then
            If (curAngle > 0#) Then curAngle = curAngle - 90# Else curAngle = curAngle + 90#
        End If
        
        Dim cParams As pdSerialize
        Set cParams = New pdSerialize
        
        With cParams
            .AddParam "angle", -1# * curAngle
            .AddParam "target", pdat_SingleLayer
        End With
        
        Process "Straighten layer", , cParams.GetParamString(), UNDO_Layer
        
    End If

End Sub
