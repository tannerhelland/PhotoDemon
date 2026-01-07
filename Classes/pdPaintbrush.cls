VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdPaintbrush"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'Paintbrush tool interface
'Copyright 2016-2026 by Tanner Helland
'Created: 1/November/16
'Last updated: 28/January/25
'Last update: again provide a floating-point line interpolator (and give users a toggle for strict vs loose pixel alignment)
'
'This module handles the messy business of translating mouse events (or really, a list of input
' coordinates) into a stream of paint dabs.
'
'All of PD's paint tools rely on this class in some way.
'
'The class uses a simple design:
' 1) As the caller receives input events, it should simply forward them to this class
' 2) This class uses current paint tool settings to calculate a corresponding list of "dabs"
' 3) The final list of "dabs" can be queried by the caller
' 4) The caller applies the dabs however they want!
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Current mouse/pen input values.  These are blindly relayed to us by the canvas, and it's up to us
' to perform any special tracking calculations.
Private m_MouseDown As Boolean
Private Const MOUSE_OOB As Single = -9.99999E+14!

'Current dab dynamics, if any.  (These are not currently calculated, but will be someday!)
'Private m_Dynamics As BrushDynamics

'If the current stroke is the first or last stroke, these flags will be set accordingly
Private m_IsFirstDab As Boolean, m_IsLastDab As Boolean

'To improve responsiveness, we measure the time delta between viewport refreshes.  If painting occurs
' fast enough, we coalesce screen updates together, as they are (by far) the most time-consuming segment
' of paint rendering; similarly, if painting is too slow, we temporarily reduce viewport update frequency
' until the background painting task "catches up."  This is IMO a preferable UI approach to other software
' the attempts to update the screen in asynchronous "chunks", which causes persistent, unavoidable
' checkerboarding across the image.
Private m_TimeSinceLastRender As Currency, m_NetTimeToRender As Currency, m_NumRenders As Long, m_FramesDropped As Long

'Current stack of points that the caller needs to draw.  We relay these to them on-demand (as they call
' the RetrievePoint() function), and when all points are exhausted, we reset the pointer to position 0.
Private Const INIT_POINT_STACK_SIZE As Long = 8
Private m_PointStack() As PointFloat
Private m_NumPoints As Long, m_PointIndex As Long

'The last *processed* point, e.g. the last point the caller retrieved (and presumably handled).
' This is used to implement shift+click functionality, where the paint engine auto-draws connecting
' lines between the current point and the last-clicked point.
Private m_LastPoint As PointFloat

'Our parent should call this function before updating the screen, and it should pass a high-def
' timer value that was retrieved *before* it started processing the current stroke.
'
'Why? This class automatically calculates frame rate (based on event notifications) and if screen
' updates start happening too slowly, it will tell you to suspend updates until background painting
' tasks "catch up".  Because this process is critical to *all* on-canvas tools, we process it here
' (rather than forcing each tool to implement something similar individually).
Friend Function IsItTimeForScreenUpdate(ByVal timeAtStrokeStart As Currency) As Boolean

    'If this is the first paint stroke, we always want to update the viewport to reflect that.
    Dim updateViewportNow As Boolean
    updateViewportNow = Me.IsFirstDab
    
    'In the background, paint tool rendering is uncapped.  (60+ fps is achievable on most modern PCs, thankfully.)
    ' However, relaying those paint tool updates to the screen is a time-consuming process, as we have to composite
    ' the full image, apply color management, calculate zoom, and a whole bunch of other crap.  Because of this,
    ' it improves the user experience to run background paint calculations and on-screen viewport updates at
    ' different framerates, with an emphasis on making sure the *background* paint tool rendering gets top priority.
    If (Not updateViewportNow) Then
        
        'If this is the first frame we're rendering (which should have already been caught by the "isFirstStroke"
        ' check above), force a render
        If (m_NumRenders > 0) Then
        
            'Perform some quick heuristics to determine if brush performance is lagging; if it is, we can
            ' artificially delay viewport updates to compensate.  (On large images and/or at severe zoom-out values,
            ' viewport rendering consumes a disproportionate portion of the brush rendering process.)
            'Debug.Print "Average render time: " & Format$((m_NetTimeToRender / m_NumRenders) * 1000, "0000") & " ms"
            
            'Calculate an average per-frame render time for the current stroke, in ms.
            Dim avgFrameTime As Currency
            avgFrameTime = (m_NetTimeToRender / m_NumRenders) * 1000
            
            'If our average rendering time is "good" (above 15 fps), allow viewport updates to occur "in realtime",
            ' e.g. as fast as the background brush rendering.
            If (avgFrameTime < 66) Then
                updateViewportNow = True
            
            'If our average frame rendering time drops below 15 fps, start dropping viewport rendering frames, but only
            ' until we hit the (barely workable) threshold of 2 fps - at that point, we have to provide visual feedback,
            ' whatever the cost.
            Else
                
                'Never skip so many frames that viewport updates drop below 2 fps.  (This is absolutely a
                ' "worst-case" scenario, and it should never be relevant except on the lowliest of PCs.)
                updateViewportNow = (VBHacks.GetTimerDifferenceNow(m_TimeSinceLastRender) * 1000 > 500#)
                
                'If we're somewhere between 2 and 15 fps, keep an eye on how many frames we're dropping.  If we drop
                ' *too* many, the performance gain is outweighed by the obnoxiousness of stuttering screen renders.
                If (Not updateViewportNow) Then
                    
                    'This frame is a candidate for dropping.
                    Dim frameCutoff As Long
                    
                    'Next, determine how many frames we're allowed to drop.  As our average frame time increases,
                    ' we get more aggressive about dropping frames to compensate.  (This sliding scale tops out at
                    ' dropping 5 consecutive frames, which is pretty damn severe - but note that framerate drops
                    ' are also limited by the 2 fps check before this If/Then block.)
                    If (avgFrameTime < 100) Then
                        frameCutoff = 1
                    ElseIf (avgFrameTime < 133) Then
                        frameCutoff = 2
                    ElseIf (avgFrameTime < 167) Then
                        frameCutoff = 3
                    ElseIf (avgFrameTime < 200) Then
                        frameCutoff = 4
                    Else
                        frameCutoff = 5
                    End If
                    
                    'Keep track of how many frames we've dropped in a row
                    m_FramesDropped = m_FramesDropped + 1
                    
                    'If we've dropped too many frames proportionate to the current framerate, cancel this drop and
                    ' update the viewport.
                    If (m_FramesDropped > frameCutoff) Then updateViewportNow = True
                    
                End If
                
            End If
        
        End If
        
    End If
    
    'If a viewport update is recommended, reset the frame drop counter and the
    ' "time since last viewport render" tracker
    If updateViewportNow Then
        m_FramesDropped = 0
        VBHacks.GetHighResTime m_TimeSinceLastRender
    End If
    
    'Report our recommendation to the caller
    IsItTimeForScreenUpdate = updateViewportNow
    
End Function

'Partner function to IsItTimeForScreenUpdate(), above.  Call this *after* redrawing the viewport
' and pass it the same high-def time passed to IsItTimeForScreenUpdate; the elapsed time since
' the stroke began will be factored into our screen update recommendation for the next stroke.
Friend Sub NotifyScreenUpdated(ByVal timeAtStrokeStart As Currency)

    'Update our running "time to render" tracker
    m_NetTimeToRender = m_NetTimeToRender + VBHacks.GetTimerDifferenceNow(timeAtStrokeStart)
    m_NumRenders = m_NumRenders + 1
    
End Sub

Friend Function IsFirstDab() As Boolean
    IsFirstDab = m_IsFirstDab
End Function

Friend Function IsLastDab() As Boolean
    IsLastDab = m_IsLastDab
End Function

Friend Function IsMouseDown() As Boolean
    IsMouseDown = m_MouseDown
End Function

'Notify the brush engine of a new mouse event.  As is standard with all PD brush events, coordinates should
' always be in *image* coordinate space, *not* screen space.
'
'Passing the current shift modifier is also important, as this class will auto-calculate straight lines of
' dabs for shift-click patterns.
Friend Sub NotifyBrushXY(ByVal mouseButtonDown As Boolean, ByVal Shift As ShiftConstants, ByVal srcX As Single, ByVal srcY As Single, ByVal mouseTimeStamp As Long, Optional ByVal useIntegerAlignment As Boolean = False)
    
    'Start by checking for first/last dabs
    m_IsFirstDab = (Not m_MouseDown) And mouseButtonDown
    m_IsLastDab = m_MouseDown And (Not mouseButtonDown)
    
    'Track mouse state
    m_MouseDown = mouseButtonDown
    
    'If too many mouse movements occur close together (important when the viewport is heavily zoomed-in),
    ' we'll skip paint events until a "significant" number of events have occurred.
    Dim storeAsLastPoint As Boolean
    storeAsLastPoint = True
    
    'If this is the first dab of this stroke, reset some internal parameters
    If m_IsFirstDab Then
    
        'Reset all time trackers; we *always* want to update the screen on the first dab,
        ' so the user has visual feedback that the operation is working
        m_NetTimeToRender = 0
        m_NumRenders = 0
        m_FramesDropped = 0
        m_PointIndex = 0
        
        'Make sure the point queue is cleared
        m_PointIndex = 0
        m_NumPoints = 0
        
        'Next, determine if the shift key is being pressed.  If it is, and if the user has already committed a
        ' brush stroke to this image (on a previous paint tool event), we want to draw a smooth line between the
        ' last paint point and the current one.  Note that this special condition is stored at module level,
        ' as we render a custom UI on mouse move events if the mouse button is *not* pressed, to help communicate
        ' what the shift key does.
        Dim useShiftStyle As Boolean
        useShiftStyle = ((Shift And vbShiftMask) <> 0) And (m_LastPoint.x <> MOUSE_OOB) And (m_LastPoint.y <> MOUSE_OOB)
        If useShiftStyle Then useShiftStyle = (m_LastPoint.x <> srcX) Or (m_LastPoint.y <> srcY)
        
        'Add this point to the collection as-is, *UNLESS* the shift button is down.
        ' If it is, we actually want to interpolate between the last-clicked point and this one.
        If useShiftStyle Then
            m_IsFirstDab = False
            If useIntegerAlignment Then
                AddPoints_BresenhamStyle m_LastPoint.x, m_LastPoint.y, srcX, srcY, False
            Else
                AddPoints_NaiveFloat m_LastPoint.x, m_LastPoint.y, srcX, srcY, False
            End If
        Else
            AddSinglePoint srcX, srcY
        End If
        
    'If this is *not* the first dab, we need to add all points between the previous point and this one.
    Else
        
        'Failsafe check for thisPoint = lastPoint; if this happens, skip line traversal
        ' and just add this point "as-is"
        If m_MouseDown Then
            
            'If ((srcX <> m_LastPoint.x) Or (srcY <> m_LastPoint.y)) And ((m_LastPoint.x <> MOUSE_OOB) And (m_LastPoint.y <> MOUSE_OOB)) Then
            
            If ((m_LastPoint.x <> MOUSE_OOB) And (m_LastPoint.y <> MOUSE_OOB)) Then
                
                'Before adding this point, enforce a "minimum distance" (currently 1/4 of a pixel)
                ' that the mouse must move before applying this stroke.  If we don't do this,
                ' small mouse movements on a zoomed-in canvas will generate an explosion of dabs,
                ' which greatly interferes with soft edge brush behavior (because the dabs will all
                ' overlap each other, effectively "erasing" the soft edges).
                If (Abs(m_LastPoint.x - srcX) > 0.25) Or (Abs(m_LastPoint.y - srcY) > 0.25) Then
                    If useIntegerAlignment Then
                        AddPoints_BresenhamStyle m_LastPoint.x, m_LastPoint.y, srcX, srcY, True
                    Else
                        AddPoints_NaiveFloat m_LastPoint.x, m_LastPoint.y, srcX, srcY, True
                    End If
                Else
                    storeAsLastPoint = False
                End If
                
            Else
                AddSinglePoint srcX, srcY
            End If
            
        End If
        
    End If
    
    'If the mouse is down (or this is the "last dab"), update our running last-passed-point tracker
    If ((m_MouseDown Or m_IsLastDab) And storeAsLastPoint) Then
        m_LastPoint.x = srcX
        m_LastPoint.y = srcY
    End If
    
End Sub

Friend Function GetNextPoint(ByRef dstPoint As PointFloat) As Boolean
    
    'Failsafe check; make sure we actually have points to return!
    GetNextPoint = (m_PointIndex < m_NumPoints)
    
    'If points remain, return the current one and increment the index into the point collection
    If GetNextPoint Then
        dstPoint = m_PointStack(m_PointIndex)
        m_PointIndex = m_PointIndex + 1
        
    'If all points have been returned, reset our current stack (but leave the stack allocation
    ' at its current size - it consumes a minimal amount of memory, and we don't want to churn
    ' by constantly resizing it.)
    Else
        m_PointIndex = 0
        m_NumPoints = 0
    End If
    
End Function

Friend Function GetLastAddedPoint(ByRef dstPoint As PointFloat) As Boolean
    dstPoint = m_LastPoint
    GetLastAddedPoint = (m_LastPoint.x <> MOUSE_OOB) And (m_LastPoint.y <> MOUSE_OOB)
End Function

Friend Sub Reset()
    
    'Reset all mouse parameters
    m_MouseDown = False
    m_IsFirstDab = False
    m_IsLastDab = False
    m_LastPoint.x = MOUSE_OOB
    m_LastPoint.y = MOUSE_OOB
    
    'Allocate an initial buffer for the point stack
    ReDim m_PointStack(0 To INIT_POINT_STACK_SIZE - 1) As PointFloat
    m_NumPoints = 0
    
    'Reset all time parameters
    m_NetTimeToRender = 0
    m_NumRenders = 0
    m_FramesDropped = 0
        
End Sub

Private Sub Class_Initialize()
    Me.Reset
End Sub

'Add all relevant points between [1] and [2] to the queue, using a traditional Bresenham line rasterizer
Private Sub AddPoints_BresenhamStyle(ByVal xStart As Single, ByVal yStart As Single, ByVal xEnd As Single, ByVal yEnd As Single, Optional ByVal addFirstPoint As Boolean = True)
    
    'This is a barebones Bresenham implementation.  It will be difficult to improve speed much beyond
    ' this code, short of specialized per-brush implementations, so this is a nice baseline for "fast but
    ' sketchy pixel coverage."  (Note that performance of this function itself is largely irrelevant --
    ' the cost of stroke rendering lies primarily in *rendering* the brush, not calculating its path.)
    
    'In classic Bresenham fashion, calculations are performed against integers.  It is up to the caller
    ' to pre-treat coordinates with any relevant rounding (and indeed this is how we do it in PD; the caller
    ' plots points against pixel centers, so the same modification is performed on coordinates returned
    ' by this rasterizer).
    Dim x0 As Long, x1 As Long, y0 As Long, y1 As Long
    x0 = Int(xStart)
    y0 = Int(yStart)
    x1 = Int(xEnd)
    y1 = Int(yEnd)
    
    'Add this point to the queue (unless it's the first point and the caller has requested otherwise)
    If addFirstPoint Then AddSinglePoint x0, y0
    
    'Calculate deltas
    Dim dx As Long, dy As Long
    dx = Abs(x1 - x0)
    dy = Abs(y1 - y0)
    
    'Calculate step directionality.
    Dim sX As Long, sY As Long
    If (x0 < x1) Then
        sX = 1
    ElseIf (x0 = x1) Then
        sX = 0
    Else
        sX = -1
    End If
    
    If (y0 < y1) Then
        sY = 1
    ElseIf (y0 = y1) Then
        sY = 0
    Else
        sY = -1
    End If
    
    'Running "errors" are used to nudge running pixel calculations in x or y directions
    Dim runningErr As Long, e2 As Long
    runningErr = dx - dy
    
    Do
        
        'Once we hit the final pixel, exit immediately.
        If ((x0 = x1) And (y0 = y1)) Then
            Exit Do
        End If
        
        'Calculate a new error, and determine if we need to advance in the X or Y direction
        e2 = 2 * runningErr
        If (e2 > -dy) Then
            runningErr = runningErr - dy
            x0 = x0 + sX
        End If
        
        If (e2 < dx) Then
            runningErr = runningErr + dx
            y0 = y0 + sY
        End If
        
        'Dab the target pixel
        AddSinglePoint x0, y0
        
    Loop
    
End Sub

'Add all relevant points between [1] and [2] to the queue, ignoring integer requirements and simply
' breaking the line down into ~1px increments.
Private Sub AddPoints_NaiveFloat(ByVal xStart As Single, ByVal yStart As Single, ByVal xEnd As Single, ByVal yEnd As Single, Optional ByVal addFirstPoint As Boolean = True)
    
    If addFirstPoint Then AddSinglePoint xStart, yStart
    
    'Ignore further plotting if start and end points match
    If (xStart <> xEnd) Or (yStart <> yEnd) Then
        
        'Break the line up into roughly 1px offsets.
        Dim lineLength As Single
        lineLength = PDMath.DistanceTwoPoints(xStart, yStart, xEnd, yEnd)
        
        'If the distance is already 1px or less, just plot both points and exit
        If (lineLength <= 1!) Then
            AddSinglePoint xEnd, yEnd
            Exit Sub
        End If
        
        'Line length is > 1.  Divide it into roughly 1px increments, and plot all points accordingly.
        Dim numLineSegments As Long
        numLineSegments = Int(lineLength + 0.5!) + 1
        
        Dim xIncrement As Single, yIncrement As Single
        xIncrement = (xEnd - xStart) / numLineSegments
        yIncrement = (yEnd - yStart) / numLineSegments
        
        Dim i As Long
        For i = 1 To numLineSegments
            AddSinglePoint xStart + (i * xIncrement), yStart + (i * yIncrement)
        Next i
        
    End If
    
End Sub

Private Sub AddSinglePoint(ByVal x As Single, ByVal y As Single)
    If (m_NumPoints > UBound(m_PointStack)) Then ReDim Preserve m_PointStack(0 To m_NumPoints * 2 - 1) As PointFloat
    m_PointStack(m_NumPoints).x = x
    m_PointStack(m_NumPoints).y = y
    m_NumPoints = m_NumPoints + 1
End Sub
