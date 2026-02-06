VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdTimerAnimation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Animation Timer
'Copyright 2016-2026 by Tanner Helland
'Created: 07/February/16
'Last updated: 12/November/20
'Last update: minor tweaks to prepare for the new Animation > Effects menu
'
'This class is an animation-centric wrapper to a normal pdTimer instance.  It handles frame calculation
' and timer interval setting in order to achieve a steady FPS target.
'
'Importantly, it is designed to work with data from animated raster files (GIF and PNG, specifically),
' which allow you to specify custom per-frame timing.
'
'Note that this class performs very few validations on things like frame indices and frame times.
' If you pass bad values, you will get bad results - this is done because timing is incredibly
' perf-sensitive, and checking every value on every access just complicates timing accuracy.
' If bad values are a concern, be proactive and do initial validations on your end.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

Public Event DrawFrame(ByVal idxFrame As Long)
Public Event EndOfAnimation()

'A base timer instance handles callbacks; we simply juggle interval settings for this timer
' to arrive at a steady rate of output events
Private WithEvents m_Timer As pdTimer
Attribute m_Timer.VB_VarHelpID = -1

'Animated GIF and PNG files are incredibly obnoxious, because they allow you to specify
' per-frame delays.  This means that each frame may have a different delay, ranging from as low
' as 0 ms to 65k ms.  Many use a 30-fps-ish target of ~30 ms delays.  Because delays can be large,
' we don't want to waste resources with a 60-fps timer - instead, we constantly modify our timer
' interval(s) to try and achieve a "perfect" match for requested frame times.  (This also allows
' us to better respect system timer patterns on e.g. power-constrained laptops.)
Private m_FrameCount As Long
Private m_FrameTimesMS() As Long
Private m_CurrentFrame As Long
Private m_RepeatAnimation As Boolean

'Frame rendering time is (obviously) tracked
Private m_TimeAtLastFrame As Currency, m_ExpectedTimeToDisplay As Currency

'These values are only used for profiling; they can be commented-out in production code
Private m_FramesDisplayed As Long, m_FrameTimes As Double

Friend Function GetCurrentFrame() As Long
    GetCurrentFrame = m_CurrentFrame
End Function

Friend Function GetRepeat() As Boolean
    GetRepeat = m_RepeatAnimation
End Function

Friend Function IsActive() As Boolean
    IsActive = m_Timer.IsActive()
End Function

'Changing the frame count will invalidate all previous frame times (obviously).
' You *must* specify new frame times after resetting frame count.
Friend Sub NotifyFrameCount(ByVal newFrameCount As Long)
    
    m_FrameCount = newFrameCount
    If (m_FrameCount > 0) Then ReDim m_FrameTimesMS(0 To m_FrameCount - 1) As Long
    
    'Reset frame time trackers (debug only)
    m_FramesDisplayed = 0
    m_FrameTimes = 0#
    
End Sub

'Notify the class of a given frame time for a specific frame.  If a frame is not specified,
' *ALL* frames will be set to the requested time.
Friend Sub NotifyFrameTime(ByVal frameTime As Long, Optional ByVal idxFrame As Long = -1)
    If (m_FrameCount <= 0) Then Exit Sub
    If (idxFrame >= 0) Then
        m_FrameTimesMS(idxFrame) = frameTime
    Else
        Dim i As Long
        For i = 0 To m_FrameCount - 1
            m_FrameTimesMS(i) = frameTime
        Next i
    End If
End Sub

Friend Sub SetCurrentFrame(ByVal idxFrame As Long)
    If (m_CurrentFrame <> idxFrame) Then
        If (idxFrame < 0) Then idxFrame = 0
        If (idxFrame >= m_FrameCount) Then idxFrame = m_FrameCount - 1
        m_CurrentFrame = idxFrame
        RaiseEvent DrawFrame(idxFrame)
    End If
End Sub

Friend Sub SetRepeat(ByVal newState As Boolean)
    m_RepeatAnimation = newState
End Sub

Friend Sub StartTimer()

    'Failsafe: validate the current frame index
    If (m_CurrentFrame < 0) Then m_CurrentFrame = 0
    If (m_CurrentFrame >= m_FrameCount) Then m_CurrentFrame = 0
    
    'Next, we're gonna prep a timer object based on the required delay for this animation frame.
    ' (Note that Windows does not support timer accuracy below 10 ms, so we lock our timer
    ' requests to never be less than 10ms.)
    Dim targetDelayMS As Long
    targetDelayMS = m_FrameTimesMS(m_CurrentFrame)
    If (targetDelayMS < 10) Then targetDelayMS = 10
    
    m_Timer.Interval = targetDelayMS
    m_Timer.StartTimer
    
    'Render the current frame, then exit
    VBHacks.GetHighResTimeInMS m_TimeAtLastFrame
    RaiseEvent DrawFrame(m_CurrentFrame)
    m_CurrentFrame = m_CurrentFrame + 1
    
End Sub

Friend Sub StopTimer()
    StopAnimation
End Sub

Private Sub Class_Initialize()
    Set m_Timer = New pdTimer
End Sub

Private Sub Class_Terminate()
    Set m_Timer = Nothing
End Sub

Private Sub m_Timer_Timer()

    'Failsafe check for "still animating".  (Remember that WM_TIMER messages are low-priority;
    ' they may stack up as other messages are processed.)
    If (Not m_Timer.IsActive) Then Exit Sub
    
    'Failsafe check for frame count
    If OutOfFrames() Then Exit Sub
    
    'Delays are calculated according to the *previous* frame's delay
    Dim relevantFrame As Long
    relevantFrame = m_CurrentFrame - 1
    
    If (relevantFrame < 0) And m_RepeatAnimation Then relevantFrame = m_FrameCount - 1
    
    'If this frame went over-budget, we want to subtract the difference from the next frame's
    ' requested delay; as long as delays are small, this is enough to keep rendering reasonably
    ' well synchronized.
    Dim frameDeficit As Currency, timeElapsedMS As Currency
    frameDeficit = 0
    
    'Perform drop-frame testing (but never on the first frame!)
    If (relevantFrame >= 0) And (m_ExpectedTimeToDisplay <> 0) Then
        
        'If more time has elapsed than the frame delay we originally requested, we may need to skip
        ' the current frame - and possibly even more frames after that.  (Note that timer events are
        ' not especially precise, especially on Win 8+ because we use coalescing timers to improve
        ' battery life - so the likelihood of a "perfect" timer interval is very low.)
        timeElapsedMS = (VBHacks.GetHighResTimeInMSEx() - m_ExpectedTimeToDisplay)
        
        If (timeElapsedMS > 0@) Then
            
            'This frame arrived late.
            
            'See if we're also over-budget for the next frame in line (by measuring the delay of
            ' the *current* frame - remember, delays in animated files specify the delay *after*
            ' the current frame).
            If (timeElapsedMS > m_FrameTimesMS(m_CurrentFrame)) Then
                
                'Damn - we're too late to render this frame in time.  Start searching through the
                ' frame list until we arrive at the frame nearest our current delay.
                Dim netDelay As Long
                netDelay = m_FrameTimesMS(m_CurrentFrame)
                relevantFrame = GetNextFrame(m_CurrentFrame)
                
                'We'll also add a failsafe check for long delays, in case something crazy happens
                ' like suspending the PC mid-animation, then returning later
                Const MAX_FRAMES_SKIPPED As Long = 15
                Dim numFramesSkipped As Long
                numFramesSkipped = 0
                
                Do While (timeElapsedMS > netDelay) And (relevantFrame < m_FrameCount) And (numFramesSkipped < MAX_FRAMES_SKIPPED)
                
                    'Increment the net delay
                    netDelay = netDelay + m_FrameTimesMS(relevantFrame)
                    relevantFrame = GetNextFrame(relevantFrame)
                    numFramesSkipped = numFramesSkipped + 1
                
                Loop
                
                'The net delay now exceeds the delay that has already occurred.  Calculate a time deficit,
                ' then display the frame *before* the currently calculated one.
                relevantFrame = relevantFrame - 1
                If (relevantFrame < 0) Then relevantFrame = m_FrameCount - 1
                netDelay = netDelay - m_FrameTimesMS(relevantFrame)
                
                frameDeficit = -1 * (timeElapsedMS - netDelay)
                
                'Note that we don't need to check "repeat animation" status here, as a single-play animation
                ' will still want to display the final frame before exiting
                m_CurrentFrame = relevantFrame
                
            'This frame arrived late, but there's still plenty of time to display it.  Subtract the
            ' already-acquired delay amount from our next timer request, which will hopefully bring
            ' timings back in line.
            Else
                frameDeficit = -1 * Int(timeElapsedMS + 0.5)
            End If
            
        'Frame is early or exactly on-time.  Calculate a frame deficit, if any, which we'll add to
        ' the next frame's delay.  (This helps correct for millisecond-level variations in timer events.)
        Else
            frameDeficit = Int(timeElapsedMS + 0.5)
        End If
    
    End If
    
    'Want to know average frame-times?  Uncomment these lines.
    'm_FramesDisplayed = m_FramesDisplayed + 1
    'm_FrameTimes = m_FrameTimes + (VBHacks.GetHighResTimeInMSEx() - m_TimeAtLastFrame)
    'Debug.Print Format$(CDbl(m_FrameTimes) / CDbl(m_FramesDisplayed), "0.000")
    
    'Note the current time (so the next frame has a reference point)
    VBHacks.GetHighResTimeInMS m_TimeAtLastFrame
    
    'Notify the caller that it's time to render
    RaiseEvent DrawFrame(m_CurrentFrame)
    
    'Advance the frame counter
    m_CurrentFrame = m_CurrentFrame + 1
    
    'If infinite repeats are active, roll the frame counter around m_FrameCount
    If m_RepeatAnimation And (m_CurrentFrame >= m_FrameCount) Then m_CurrentFrame = 0
    
    'If frames remain, figure out an appropriate timer interval for the next frame
    If (m_CurrentFrame < m_FrameCount) Then
        
        relevantFrame = m_CurrentFrame - 1
        If (relevantFrame < 0) Then relevantFrame = m_FrameCount - 1
        
        Dim timeIntervalToRequest As Long
        timeIntervalToRequest = m_FrameTimesMS(relevantFrame) + frameDeficit
        
        'Cache what time we expect the next frame to display; the next iteration will use this value
        ' to recenter itself accordingly.
        m_ExpectedTimeToDisplay = m_TimeAtLastFrame + timeIntervalToRequest
        
        'Windows timers don't allow timers to trigger faster than 10 ms
        If (timeIntervalToRequest < 10) Then timeIntervalToRequest = 10
        m_Timer.Interval = timeIntervalToRequest
    
    'This is a 1x animation.  Ensure the frame position is valid, then exit
    Else
        m_CurrentFrame = m_FrameCount - 1
        StopAnimation
    End If
    
End Sub

'Given a frame index, return the "next" one.  For loop animations, this automatically wraps frame indices.
' For non-repeating animations, this will return an invalid index (m_FrameCount) by design.  You must
' check for this return and respond accordingly.
Private Function GetNextFrame(ByVal curFrame As Long) As Long
    GetNextFrame = curFrame + 1
    If (GetNextFrame >= m_FrameCount) Then
        If m_RepeatAnimation Then GetNextFrame = 0
    End If
End Function

'Check to see if we've run out of frames to display; this is used for "play once" functionality
Private Function OutOfFrames() As Boolean
    
    OutOfFrames = False
    If (m_FrameCount = 0) Then
        OutOfFrames = True
        StopAnimation
        Exit Function
    End If
    
    'Failsafe check for frame count
    If (m_CurrentFrame >= m_FrameCount) Then
        If m_RepeatAnimation Then
            m_CurrentFrame = 0
        Else
            m_CurrentFrame = m_FrameCount - 1
            StopAnimation
            OutOfFrames = True
        End If
    End If
    
End Function

Friend Sub StopAnimation()
    
    If m_Timer.IsActive Then
        
        m_Timer.StopTimer
        
        'Because the animation timer post-updates the current frame (e.g. it increments it *after* rendering
        ' the current frame), the current frame marker is going to be 1 higher than whatever is displayed on-screen.
        ' Decrement it to ensure that the on-screen image matches our internal tracker.
        If (m_CurrentFrame > 0) And (m_CurrentFrame < m_FrameCount - 1) Then m_CurrentFrame = m_CurrentFrame - 1
        RaiseEvent EndOfAnimation
        
    End If
    
    m_ExpectedTimeToDisplay = 0
    
End Sub
