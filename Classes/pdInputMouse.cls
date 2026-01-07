VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdInputMouse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Mouse Input Handler (mouse, pen, etc) class
'Copyright 2014-2026 by Tanner Helland
'Created: 27/May/14 (though many individual parts existed earlier than this!)
'Last updated: 29/October/19
'Last update: manually raise corresponding _MouseDown() event in _DoubleClick() handler
'
'As part of implementing paint tools in PhotoDemon, a comprehensive input solution was required.  (Prior to that point,
' a combination of intrinsic VB mouse events and a few extra subclassing bits (e.g. mousewheel) covered PD pretty well,
' but drawing tools and more advanced UI features require more detailed input handling, like GetMouseMovePointsEx +
' spline interpolation for buttery smooth mouse input, or improved support for touch and/or pen input.)
'
'A few important notes when using this class:
'
'- This class can optionally handle standard mouse events for a given hWnd (e.g. mouse events VB normally covers, like
'   Click, DoubleClick, MouseMove, etc).  There are a lot of reasons to do this, but among the obvious benefits are things
'   like x/y coordinates for click events, support for 32-bit mouse positions, support for X-keys as a button type, and more.
'   In the future, a custom version of these events could also supply things like a Pressure parameter for MouseDown events.
'
'- Mouse coordinate, button, and shift key modifiers are processed independent of window messages.  This allows the class to
'   supply that data even if a given window message doesn't automatically include them.  It also allows for higher accuracy
'   when tracking mouse move data, as we use 32-bit values instead of 16-bit ones (which seems ridiculous now, but may not
'   in the future).
'
'- As a bonus, this class can also (optionally) handle WM_APP_COMMAND messages.  This is useful for handling some buttons that
'   may appear on a mouse, such as back/forward, instead of using X-button messages (which may not be accurate, if the user has
'   remapped functionality in a non-standard way).
'
'Many thanks to Kroc Camen of camendesign.com, whose bluMouseEvents class served PD well for a long time prior to this
' implementation.  If you need a lightweight mouse-handler that works well as a standalone solution, I recommend using
' that instead of this very PD-specific class.  You can download a copy here (link good as of May '14):
' https://github.com/Kroc/MaSS1VE/tree/master/CODE/Blu
'
'Also, thank you to Steve McMahon for translating the GET_APPCOMMAND_LPARAM, GET_DEVICE_LPARAM, and GET_KEYSTATE_LPARAM macros
' into VB.  You can see Steve's original work on WM_APPCOMMAND messages here:
' http://www.vbaccelerator.com/home/VB/Tips/Responding_to_AppCommands/article.asp
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'This class can potentially raise many events.  Note that events will only be raised if the associated tracking behavior
' is explicitly requested (e.g. you don't *have* to request that the class handles all these events).
Public Event AppCommand(ByVal cmdID As AppCommandConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
Public Event ClickCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
Public Event DoubleClickCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
Public Event MouseDownCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal timeStamp As Long)
Public Event MouseEnter(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
Public Event MouseHover(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
Public Event MouseLeave(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
Public Event MouseMoveCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal timeStamp As Long)
Public Event MouseUpCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal clickEventAlsoFiring As Boolean, ByVal timeStamp As Long)
Public Event MouseWheelHorizontal(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal scrollAmount As Double)
Public Event MouseWheelVertical(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal scrollAmount As Double)
Public Event MouseWheelZoom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal zoomAmount As Double)

'This class tracks all kinds of window messages.  Most objects won't need this level of tracking, but they're there
' if needed.
Private Const WM_MOUSEHWHEEL As Long = &H20E&
Private Const WM_MOUSEWHEEL As Long = &H20A&
Private Const WM_MOUSEHOVER As Long = &H2A1&
Private Const WM_MOUSELEAVE As Long = &H2A3&
Private Const WM_MOUSEMOVE As Long = &H200&

Private Const WM_SETCURSOR As Long = &H20

Private Const WM_LBUTTONDOWN As Long = &H201
Private Const WM_MBUTTONDOWN As Long = &H207
Private Const WM_RBUTTONDOWN As Long = &H204
Private Const WM_XBUTTONDOWN As Long = &H20B

Private Const WM_LBUTTONUP As Long = &H202
Private Const WM_MBUTTONUP As Long = &H208
Private Const WM_RBUTTONUP As Long = &H205
Private Const WM_XBUTTONUP As Long = &H20C

Private Const WM_LBUTTONDBLCLK As Long = &H203
Private Const WM_MBUTTONDBLCLK As Long = &H209
Private Const WM_RBUTTONDBLCLK As Long = &H206
Private Const WM_XBUTTONDBLCLK As Long = &H20D

'X buttons (sometimes called buttons 4 and 5 in MSDN docs) are often used for forward/back maneuvering.
' Users of this class can check their states in Mouse Up/Down events, but PD will preferentially use WM_APPCOMMAND instead,
' as it can be raised by both keyboard and mouse equivalents of forward/back keys, which is typically a better solution.
Private Const WM_APPCOMMAND As Long = &H319

'Mouse-tracking for hover and leave events is not handled automatically by Windows; we must request it.
Private Type TRACKMOUSEEVENT_STRUCT
    cbSize As Long
    dwFlags As Long
    hWndTrack As Long
    dwHoverTime As Long
End Type

Private Const TME_HOVER As Long = &H1&
Private Const TME_LEAVE As Long = &H2&

Private Declare Function TrackMouseEvent Lib "user32" (ByRef lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long

'Virtual key-codes currently supported by pdInput
Private Const VK_LBUTTON As Long = &H1
Private Const VK_RBUTTON As Long = &H2
Private Const VK_MBUTTON As Long = &H4
Private Const VK_XBUTTON1 As Long = &H5
Private Const VK_XBUTTON2 As Long = &H6

'In the future, other virtual key codes can be retrieved here:
' http://msdn.microsoft.com/en-us/library/windows/desktop/dd375731%28v=vs.85%29.aspx

'Mouse buttons can be retrieved from various mouse messages, but for consistency's sake, we pull them straight
' from GetAsyncKeyState.  One thing to note about GetAsyncKeyState is that it returns the physical mouse button
' pressed, *without button swapping* for left-handed mouse users.  We need to check this state and manually
' handle it, using GetSystemMetrics().
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Const SM_SWAPBUTTON As Long = 23

'Mouse capturing provides more predictable behavior, particularly on UCs
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long

'This class also handles cursor management, but because cursors can be requested by external functions, they are declared
' publicly in the Icon and Cursor module.  Look there for cursor-related API constants.

'Retrieve the current cursor position, in screen coordinates
Private Declare Function GetCursorPos Lib "user32" (ByRef lpPoint As PointAPI) As Long

'Set a new cursor for a given class
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorW" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare Sub SetCursorPos Lib "user32" (ByVal newX As Long, ByVal newY As Long)
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long

'Reconstruct a full mouse movement history.  Note that this requires the MOUSEMOVEPOINT type, declared elsewhere.
Private Enum PD_MouseResolution
    GMMP_USE_DISPLAY_POINTS = 1
    GMMP_USE_HIGH_RESOLUTION_POINTS = 2
End Enum

#If False Then
    Const GMMP_USE_DISPLAY_POINTS = 1, GMMP_USE_HIGH_RESOLUTION_POINTS = 2
#End If

Private Declare Function GetMouseMovePointsEx Lib "user32" (ByVal sizeOfMouseMoveStruct As Long, ByRef currentMouseMovePoint As MOUSEMOVEPOINT, ByVal ptrToMouseMovePointArray As Long, ByVal numOfPointsToRetrieve As Long, ByVal resolutionFlag As PD_MouseResolution) As Long
Private Declare Function GetMessageTime Lib "user32" () As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long

'Weird legacy desktop issues require us to transform points passed to/from GetMouseMovePointsEx, relative to the current monitor system
Private Const SM_XVIRTUALSCREEN As Long = 76
Private Const SM_YVIRTUALSCREEN As Long = 77
Private Const SM_CXVIRTUALSCREEN As Long = 78
Private Const SM_CYVIRTUALSCREEN As Long = 79

'API helper functions for converting between screen and client coordinate spaces
Private Declare Function MapWindowPoints Lib "user32" (ByVal hWndFrom As Long, ByVal hWndTo As Long, ByVal ptrToPointList As Long, ByVal numPoints As Long) As Long

'MSDN best practices suggest that we always retrieve the user's setting for scroll wheel sensitivity, as it may change
' while a program is running.  SystemParametersInfo is used for this.  Note that horizontal and vertical settings
' are stored separately.
Private Declare Function SystemParametersInfoW Lib "user32" (ByVal uiAction As Long, ByVal uiParam As Long, ByRef pvParam As Long, ByVal fWinIni As Long) As Long
Private Const SPI_GETWHEELSCROLLLINES As Long = &H68
Private Const SPI_GETWHEELSCROLLCHARS As Long = &H6C

'Central subclasser for all input window messages
Implements ISubclass

'The hWnd currently being tracked.  You can test this value against zero to see if the class is active.
Private m_hWnd As Long

'By default, this class handles all standard mouse events, including enter/leave and wheel events.  You can also request
' other types of special events (like AppCommands); when those are active, the associated tracker variables will also
' be activated, so we know which subclassed events to report back to the user.
Private m_TrackAppCommands As Boolean, m_DisregardWheelEvents As Boolean

'Windows requires you to file formal mouse tracking requests for things like mouse enter/leave messages.  If we have filed a request
' (and said request has not been canceled), this will be TRUE.  Do not file requests for tracking until this returns to FALSE.
Private m_MouseTrackingActive As Boolean

'To prevent multiple hover events from mistakenly being raised, we manually track hover state and disallow events until we detect
' mouse movement since the last hover.
Private m_HoverModeActive As Boolean, m_LastHoverX As Single, m_LastHoverY As Single

'If the user requests that we force a persistent cursor for this object, that cursor handle will be stored here.
Private m_CursorHandle As Long

'This class will generate Click() events using its own algorithm for determining when an action constitutes a "click".
' Basically, a MouseDown/Up combination with two or less MouseMove messages between them constitutes a Click.  This variable
' is used to count MouseMove occurrences.
Private m_MouseMoveCount As Long

'While we can't track mouse events any faster than the WM_MOUSEMOVE messages we receive, we can reconstruct missed mouse events by
' using GetMouseMovePointsEx.  To activate this feature, pass TRUE to the SetHighResolutionTrackingMode sub, which will in turn
' activate this boolean.  If active, the subclasser will reconstruct any missing mouse events history, and raise each missed event
' in turn.
Private m_HighResModeActive As Boolean

'To reconstruct missing mouse movement points, we must track the last returned mouse point.  Otherwise, we risk repeating points.
Private Const MAX_MOUSE_HISTORY As Long = 64
Private m_MouseHistory() As MOUSEMOVEPOINT
Private m_PrevPointStored As Boolean
Private m_LastPointTracked As MOUSEMOVEPOINT

'Our parent control may want to know how many MOUSEMOVEPOINT objects have posted but not been relayed (yet).  This allows them
' to delay handling of other items until mouse events catch up, for example.  We track the current "not yet relayed" MMP count
' here.
Private m_MouseMovePointsRemaining As Long

'Because PD sometimes performs energy-intensive actions within mouse events, it is possible for raised messages to become
' increasingly delayed, especially if high-resolution tracking is active.  In an attempt to mitigate this, this class supports
' auto-dropping of messages delayed beyond a certain threshold.  NOTE: this behavior only affects MouseMove events.  All other
' mouse events will be fired regardless of delay.
Private m_DelayTrackingActive As Boolean
Private m_DelayThreshold As Long
Private Const DEFAULT_DELAY_THRESHOLD_MS = 67

'To improve mouse behavior on UCs, this class will auto-capture the mouse when a button is pressed.  Each press of a button increments
' this internal capture counter.  When a button is released, the capture counter is decremented.  When all buttons have been released,
' the mouse capture is released.
Private m_CaptureCounter As Long

'By default, PD forcibly captures the mouse when a button is pressed.  To disable this behavior, set this value to TRUE.
Private m_DoNotCaptureOnButtonPress As Boolean

'If the mouse is currently captured, this will be set to TRUE
Private m_MouseHasBeenCaptured As Boolean

'Because this class performs subclassing, and a lot of mouse events may be generated, we use module-level variables
' for a number of "inside the subclasser" variables (to prevent flooding the stack with temp variables).
Private m_retShiftConstants As ShiftConstants, m_retMouseButton As PDMouseButtonConstants
Private m_scrollDelta As Long, m_finalScrollAmount As Double
Private m_commandID As Long, m_msgTimestamp As Long
Private m_MouseX As Long, m_MouseY As Long
Private m_MouseXInt As Integer, m_MouseYInt As Integer

'In some rare cases, we still want to allow the default mouse handler to "do its thing" after PD performs
' any specialized handling. If this rare case is active, this will be set to TRUE.
Private m_StillUseDefWindowProc As Boolean

'Initiate input tracking of a given hWnd.  This will replace all of VB's internal mouse events with events generated by this class.
' Mouse Enter/Leave events and MouseWheel events will also be automatically generated.
'
'In addition to the default functionality, you can request a few other types special events:
'
' - Whether to track AppCommand messages for this hWnd.  This is used for "virtual" functions like forward/back,
'    which are typically generated by mouse buttons 4 and 5, but could also be generated by a multimedia keyboard
'    or gestures on a tablet PC.
'
' - Cursor management.  Generally, it's best to let this class handle cursor behavior, as it has more detailed information
'    available than VB does, but you're allowed to handle this manually if you desire.
'
'Note that this function will return TRUE if input tracking was initiated successfully.
Friend Function AddInputTracker(ByVal targetHWnd As Long, Optional ByVal alsoTrackAppCommands As Boolean = False, Optional ByVal disregardWheelEvents As Boolean = False) As Boolean
    
    If (PDMain.IsProgramRunning() And (targetHWnd <> 0)) Then
    
        'Cache relevant parameters at module-level.  These affect the way the subclasser responds to various window messages.
        m_hWnd = targetHWnd
        m_TrackAppCommands = alsoTrackAppCommands
        m_DisregardWheelEvents = disregardWheelEvents
        m_MouseTrackingActive = False
        
        'Attach a subclasser to the target hWnd
        VBHacks.StartSubclassing m_hWnd, Me
                    
        'If enter, leave, and hover events are desired, immediately start mouse tracking
        m_MouseTrackingActive = RequestMouseTrackingForHwnd()
        AddInputTracker = m_MouseTrackingActive
        
    End If
    
End Function

'Use this function enable/disable automatic dropping of severely delayed events
Friend Sub SetAutoDropOfDelayedEvents(ByVal newMode As Boolean)
    m_DelayTrackingActive = newMode
End Sub

'Use this function to set the threshold, in milliseconds, after which a message is considered unacceptably delayed.
' NOTE: system timer resolution limits this value to roughly 16 ms, so setting a threshold below that point may cause
'       nearly all events to be dropped.  Try to limit the requested value to > 100 ms for best results.
Friend Sub SetAutoDropDelayThreshold(ByVal newThreshold As Long)
    m_DelayThreshold = newThreshold
    If (m_DelayThreshold < 16) Then m_DelayThreshold = 16
End Sub

'User controls that implement custom drag/drop behavior may choose to manually disable cursor capture, as it prevents dragging outside
' a control's borders.
Friend Sub SetCaptureOverride(ByVal newCaptureMode As Boolean)
    
    m_DoNotCaptureOnButtonPress = newCaptureMode
    
    'Also, release a current capture, if any
    If (newCaptureMode And m_MouseHasBeenCaptured) Then
        m_CaptureCounter = 0
        ReleaseCapture
        m_MouseHasBeenCaptured = False
    End If
    
End Sub

'Use this function to enable/disable high-resolution tracking (where mouse motions between WM_MOUSEMOVE events are manually reconstructed)
Friend Sub SetHighResolutionTrackingMode(ByVal newMode As Boolean)
    m_HighResModeActive = newMode
    If m_HighResModeActive Then ReDim m_MouseHistory(0 To MAX_MOUSE_HISTORY - 1) As MOUSEMOVEPOINT
End Sub

'Use this function to still call DefWindowProc after a mouse event occurs
Friend Sub SetDefWindowProcFallback(ByVal newMode As Boolean)
    m_StillUseDefWindowProc = newMode
End Sub

Friend Sub ResetTracking()
    m_CaptureCounter = 0
    m_MouseMoveCount = 0
    m_MouseHasBeenCaptured = False
    m_MouseTrackingActive = False
End Sub

'This function is automatically called at Class_Terminate time, but you can also call it manually if you want
' to ensure immediately release of the subclasser.
Friend Sub Shutdown()
    If (m_hWnd <> 0) Then
        VBHacks.StopSubclassing m_hWnd, Me
        m_hWnd = 0
    End If
End Sub

'Assign a system cursor to the underlying hWnd.  Call this function without a parameter to reset the cursor to the default arrow.
Friend Sub SetCursor_System(Optional ByVal systemCursorType As SystemCursorConstant = IDC_DEFAULT, Optional ByVal applyImmediately As Boolean = True)
    
    'Load the relevant cursor handle
    If (systemCursorType = IDC_DEFAULT) Then
        m_CursorHandle = LoadCursor(0, IDC_ARROW)
    Else
        m_CursorHandle = LoadCursor(0, systemCursorType)
    End If
    
    'Check for duplicate cursor requests, and ignore them as necessary
    If applyImmediately Then ApplyCursor
        
End Sub

'Assign a PNG cursor to the subclassed hWnd.  This function leans on the Icon and Cursor module to handle the actual
' retrieval, decompression, and assembly of the cursor.  We handle it there so that requested icons can be cached;
' there is a high probability of PNG icons being requested by more than one class in a given session, so rather than
' repeat all the steps for each request, we only do it once, then cache the results.  PD itself will handle unloading
' custom cursors at exit time, so this class need not concern itself with unloading requested cursors.
Friend Sub SetCursor_Resource(ByVal pngResourceName As String, Optional ByVal cursorHotspotX As Long = 0, Optional ByVal cursorHotspotY As Long = 0)
    
    Dim tmpCursorHandle As Long
    tmpCursorHandle = IconsAndCursors.RequestCustomCursor(pngResourceName, cursorHotspotX, cursorHotspotY)
    
    If (m_CursorHandle <> tmpCursorHandle) Then
        m_CursorHandle = tmpCursorHandle
        ApplyCursor
    End If
    
End Sub

'Private helper function used by SetCursor_System and SetCursor_Resource, above.
Private Sub ApplyCursor()
    SetCursor m_CursorHandle
End Sub

'Populate a generic TrackMouseEvent struct and pass it to the TrackMouseEvent API.  (This is done frequently,
' as Windows automatically deactivates mouse tracking after raising a mouse-related event.)
Private Function RequestMouseTrackingForHwnd() As Boolean
    
    Dim eventTracker As TRACKMOUSEEVENT_STRUCT
    
    With eventTracker
        .cbSize = LenB(eventTracker)
        .dwFlags = TME_LEAVE Or TME_HOVER
        .hWndTrack = m_hWnd
        .dwHoverTime = &HFFFFFFFF        'Use the system default hover time (400 ms, I believe)
    End With
    
    'TrackMouseEvent returns non-zero values for success
    RequestMouseTrackingForHwnd = (TrackMouseEvent(eventTracker) <> 0)
    
End Function

'Note that the vKey constant below is a virtual key mapping, not necessarily a standard VB key constant
Private Function IsMouseButtonDown(ByVal vKey As Long) As Boolean
    
    'Check for left/right mouse button switching, which is a system-wide user setting
    If (GetSystemMetrics(SM_SWAPBUTTON) <> 0) Then
    
        'Reassign left/right mouse buttons as necessary
        If vKey = VK_LBUTTON Then
            vKey = VK_RBUTTON
        Else
            If vKey = VK_RBUTTON Then vKey = VK_LBUTTON
        End If
    
    End If
    
    IsMouseButtonDown = IsVirtualKeyDown(vKey)
    
End Function

'Parse out the actual app command from the lParam of a WM_APPCOMMAND message.
' Thank you to Steve McMahon for translating the GET_APPCOMMAND_LPARAM, GET_DEVICE_LPARAM, and GET_KEYSTATE_LPARAM macros
' into VB; you can see his original work here: http://www.vbaccelerator.com/home/VB/Tips/Responding_to_AppCommands/article.asp
Private Function ParseAppCommand(ByVal lParam As Long, ByRef GET_APPCOMMAND_LPARAM As Long, Optional ByRef GET_DEVICE_LPARAM As Long, Optional ByRef GET_KEYSTATE_LPARAM As Long) As Long

    'The command itself is stored as the hiword of the message, with the highest 4 bits excluded:
    GET_APPCOMMAND_LPARAM = (lParam And &HFFF0000) / &H10000
    
    'Device (mouse, keyboard, other) is derived from the highest 4 bits:
    GET_DEVICE_LPARAM = (lParam And &H70000000) / &H10000
    If (lParam And &H80000000) = &H80000000 Then
        GET_DEVICE_LPARAM = GET_DEVICE_LPARAM Or &H8000&
    End If
    
    'Key details are in the loword:
    GET_KEYSTATE_LPARAM = lParam And &HFFFF&

End Function

'Use GetCursorPos to retrieve the current mouse pointer coordinates.  Note that GetCursorPos always uses screen coordinates,
' so we need to manually translate the coords into the space of our current hWnd.
Private Function GetCurrentCursorPosition(ByRef controlX As Long, ByRef controlY As Long) As Boolean

    Dim tmpPoint As PointAPI
    If (GetCursorPos(tmpPoint) <> 0) Then
        
        'Convert the screen coordinates into the coordinate space of the supplied hWnd
        MapWindowPoints 0, m_hWnd, VarPtr(tmpPoint), 1
        controlX = tmpPoint.x
        controlY = tmpPoint.y
        GetCurrentCursorPosition = True
        
    Else
        GetCurrentCursorPosition = False
    End If

End Function

'The primary hWnd can use this function to move the mouse to some new position.  (This is not currently used.)
Friend Sub MoveCursorToNewPosition(ByVal xRelativeToHwnd As Double, ByVal yRelativeToHwnd As Double)

    'Retrieve the requested position, in screen coordinates
    Dim tmpPoint As PointAPI
    tmpPoint.x = xRelativeToHwnd
    tmpPoint.y = yRelativeToHwnd
    MapWindowPoints m_hWnd, 0, VarPtr(tmpPoint), 1
    
    'Cursor trails may cause ghosting, where we move the cursor, but a cursor is still shown at the old mouse position.
    ' To account for that, we must hide the cursor, move it, then re-display it.
    ShowCursor 0
    SetCursorPos tmpPoint.x, tmpPoint.y
    ShowCursor 1
    
End Sub

'Use SystemParametersInfo to retrieve the current user setting for mouse wheel sensitivity
Private Function GetUserScrollSetting(ByVal directionIsVertical As Long) As Double

    Dim tmpUserScrollReturn As Long
    If directionIsVertical Then
        SystemParametersInfoW SPI_GETWHEELSCROLLLINES, 0, tmpUserScrollReturn, 0
    Else
        SystemParametersInfoW SPI_GETWHEELSCROLLCHARS, 0, tmpUserScrollReturn, 0
    End If
    
    'Because we're going to use this function to calculate scroll amounts, don't allow it to be zero.
    ' (I don't know if Windows itself allows a 0 value, but it doesn't hurt to check, especially because
    ' the SystemParametersInfo call above could technically fail to return a value.)
    If (tmpUserScrollReturn = 0) Then tmpUserScrollReturn = 3
    
    GetUserScrollSetting = CDbl(tmpUserScrollReturn) / 120#

End Function

'Hi and Lo word values can be retrieved using these helper functions.  They are required because VB doesn't have an unsigned int type,
' so we have to handle the signed bit specially (argh).
Private Function GetHiWord(ByVal lParam As Long) As Integer
    If (lParam And &H80000000) Then
        GetHiWord = (lParam \ 65535) - 1
    Else
        GetHiWord = lParam \ 65535
    End If
End Function

Private Function GetLoWord(ByVal lParam As Long) As Integer
    If (lParam And &H8000&) Then
        GetLoWord = &H8000 Or (lParam And &H7FFF&)
    Else
        GetLoWord = lParam And &HFFFF&
    End If
End Function

'If high-resolution tracking is active, the subclasser will call this function when it receives a WM_MOUSEMOVE message.  This function will
' then use GetMouseMovePointsEx to reconstruct any missing mouse movement history.
Private Sub ReconstructMouseMoveHistory(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal msgTimestamp As Long)

    'First, note that the passed points are in *window* coordinates.  We need to convert them to *screen* coordinates.
    Dim clientX As Long, clientY As Long
    clientX = x
    clientY = y
    
    Dim tmpPoint As PointAPI
    tmpPoint.x = clientX
    tmpPoint.y = clientY
    MapWindowPoints m_hWnd, 0, VarPtr(tmpPoint), 1
    
    Dim screenX As Long, screenY As Long
    screenX = tmpPoint.x
    screenY = tmpPoint.y
    
    'As a failsafe, see if we have stored the value of a previously tracked point.  (The only way this fails is for the first MouseMove
    ' event a control receives; when this happens, there is nothing to reconstruct, so we should exit immediately.)
    If (Not m_PrevPointStored) Then
    
        m_PrevPointStored = True
        
        With m_LastPointTracked
            .x = screenX And &HFFFF&
            .y = screenY And &HFFFF&
            .ptTime = msgTimestamp
        End With
        
        m_MouseMovePointsRemaining = -1
        RaiseEvent MouseMoveCustom(Button, Shift, clientX, clientY, msgTimestamp)
        Exit Sub
        
    End If
    
    'If we made it all the way here, a previous point has been stored.  We will now reconstruct a full mouse history between that point
    ' and the current one.
    Dim eventAllowed As Boolean

    'Current tracking mode.  We can use either high-resolution pen input coordinates, or regular mouse coordinates.
    ' For now, PD only uses regular mouse coordinates, but perhaps we could make use of the pen code in the future,
    ' so I've included it here.
    Dim curTrackingResolution As PD_MouseResolution
    curTrackingResolution = GMMP_USE_DISPLAY_POINTS
    
    'First, we need to prepare a reference MOUSEMOVEPOINTS struct.  This struct is used as the reference for "current point",
    ' and GetMouseMovePointsEx will automatically find all mouse movement events up to and including this one.
    
    'Because GetMouseMovePointsEx does not support the concept of negative coordinates, we must transform all points passed to/from
    ' it to account for multimonitor setups.  Start by retrieving the virtual desktop size.
    Dim nVirtualWidth As Long, nVirtualHeight As Long, nVirtualLeft As Long, nVirtualTop As Long
    nVirtualWidth = GetSystemMetrics(SM_CXVIRTUALSCREEN)
    nVirtualHeight = GetSystemMetrics(SM_CYVIRTUALSCREEN)
    nVirtualLeft = GetSystemMetrics(SM_XVIRTUALSCREEN)
    nVirtualTop = GetSystemMetrics(SM_YVIRTUALSCREEN)
    
    'Populate our reference point
    Dim curMouseMovePoint As MOUSEMOVEPOINT
    With curMouseMovePoint
        .x = screenX And &HFFFF&
        .y = screenY And &HFFFF&
        .ptTime = msgTimestamp
    End With
    
    'Prepare a buffer to receive the mouse movement history.  The max number of retrievable events is 64.
    FillMemory VarPtr(m_MouseHistory(0)), LenB(curMouseMovePoint) * MAX_MOUSE_HISTORY, 0
    
    'Retrieve the mouse movement history.  The return value of the function is the number of points filled.
    Dim numPoints As Long
    numPoints = GetMouseMovePointsEx(Len(curMouseMovePoint), curMouseMovePoint, VarPtr(m_MouseHistory(0)), MAX_MOUSE_HISTORY, curTrackingResolution)
        
    'If one or more valid points were returned, carry on
    If (numPoints > 0) Then
    
        'We now want to search the mouse movement history for the previous point we returned.  All points after that will need to be manually raised.
        Dim i As Long
        For i = 0 To numPoints - 1
        
            'If we reach a point that occurred prior to our last tracked time, exit
            If (m_MouseHistory(i).ptTime < m_LastPointTracked.ptTime) Then Exit For
            
            'If we reach our exact last point time, also exit
            If (m_MouseHistory(i).ptTime = m_LastPointTracked.ptTime) And (m_MouseHistory(i).x = m_LastPointTracked.x) And (m_MouseHistory(i).y = m_LastPointTracked.y) Then Exit For
        
        Next i
        
        '"i" now contains the index of the previous point (or one that occurred just prior to it, if the previous point wasn't found).
        ' Iterate through all points following that one, and convert them to valid screen coordinates.
        numPoints = i - 1
        If (numPoints < 0) Then numPoints = 0
        
        For i = numPoints To 0 Step -1
        
            'Convert the unsigned 16-bit return values to signed 32-bit values
            Select Case curTrackingResolution
            
                Case GMMP_USE_DISPLAY_POINTS
                    If (m_MouseHistory(i).x > 32767) Then m_MouseHistory(i).x = m_MouseHistory(i).x - 65536
                    If (m_MouseHistory(i).y > 32767) Then m_MouseHistory(i).y = m_MouseHistory(i).y - 65536
                
                Case GMMP_USE_HIGH_RESOLUTION_POINTS
                    m_MouseHistory(i).x = ((m_MouseHistory(i).x * (nVirtualWidth - 1)) - (nVirtualLeft * 65536)) / nVirtualWidth
                    m_MouseHistory(i).y = ((m_MouseHistory(i).y * (nVirtualHeight - 1)) - (nVirtualTop * 65536)) / nVirtualHeight
                
            End Select
            
            'With this point successfully transformed, the last thing we need to do is convert it from screen to window coordinates.
            tmpPoint.x = m_MouseHistory(i).x
            tmpPoint.y = m_MouseHistory(i).y
            MapWindowPoints 0, m_hWnd, VarPtr(tmpPoint), 1
            m_MouseHistory(i).x = tmpPoint.x
            m_MouseHistory(i).y = tmpPoint.y
            
        Next i
        
        'Last of all, report the points.  Note that we handle this step separately, as our parent may attempt to manually
        ' retrieve mouse points (without waiting for the events to fire on their own).  This allows us to skip through the
        ' list based on how many points the caller has manually grabbed.
        m_MouseMovePointsRemaining = numPoints
        
        'Regardless of time delays, raise the initial mouse event.
        RaiseEvent MouseMoveCustom(Button, Shift, m_MouseHistory(m_MouseMovePointsRemaining).x, m_MouseHistory(m_MouseMovePointsRemaining).y, m_MouseHistory(m_MouseMovePointsRemaining).ptTime)
        m_MouseMovePointsRemaining = m_MouseMovePointsRemaining - 1
        
        'Note that our parent may have manually pulled one or more points from the queue, so m_MouseMovePointsRemaining may have
        ' changed since our last access.  Because of this, the following loop may not need to be called.
        
        'Now, proceed to fire each successive event in turn, while respecting the user's preference for delayed message handling
        Do While (m_MouseMovePointsRemaining >= 0)
            
            'If automatic message dropping is active, make sure this message falls within the acceptable threshold
            ' before firing it.
            eventAllowed = True
            If m_DelayTrackingActive Then eventAllowed = (GetMessageDelay(m_MouseHistory(m_MouseMovePointsRemaining).ptTime) <= m_DelayThreshold)
            If eventAllowed Then RaiseEvent MouseMoveCustom(Button, Shift, m_MouseHistory(m_MouseMovePointsRemaining).x, m_MouseHistory(m_MouseMovePointsRemaining).y, m_MouseHistory(m_MouseMovePointsRemaining).ptTime)
            
            m_MouseMovePointsRemaining = m_MouseMovePointsRemaining - 1
            
        Loop
        
    'If the current point is also the previous point, GetMouseMovePointsEx may not have any prior points to return.  This is fine -
    ' simply raise the point we received, then exit.
    Else
        
        m_MouseMovePointsRemaining = -1
        eventAllowed = True
        If m_DelayTrackingActive Then eventAllowed = (GetMessageDelay(msgTimestamp) <= m_DelayThreshold)
        If eventAllowed Then RaiseEvent MouseMoveCustom(Button, Shift, clientX, clientY, msgTimestamp)
        
    End If
    
    'Update the "previous point" tracker with the current point, then exit
    With m_LastPointTracked
        .x = screenX And &HFFFF&
        .y = screenY And &HFFFF&
        .ptTime = msgTimestamp
    End With
    
End Sub

'Given a message time (as retrieved by GetMessageTime(), typically), report the delay in milliseconds between the posted time and the current time.
Private Function GetMessageDelay(ByVal srcMessageTime As Long) As Long
    GetMessageDelay = GetTickCount() - srcMessageTime
End Function

'See if there are one (or more) points remaining in the GetMouseMovePointsEx queue.  This will always be zero if the
' HighResInput property is disabled.
Friend Function GetNumMouseMovePointsRemaining() As Long
    GetNumMouseMovePointsRemaining = m_MouseMovePointsRemaining
End Function

'Retrieve the next point in the GetMouseMovePointsEx queue.  This only succeeds if the HighResInput property is enabled.
' Returns: TRUE if another point exists in the queue; FALSE otherwise.
Friend Function GetNextMouseMovePoint(ByRef dstMMP As MOUSEMOVEPOINT) As Boolean

    If (m_MouseMovePointsRemaining > 0) Then
        m_MouseMovePointsRemaining = m_MouseMovePointsRemaining - 1
        dstMMP = m_MouseHistory(m_MouseMovePointsRemaining)
        GetNextMouseMovePoint = True
    End If
    
End Function

Private Sub Class_Initialize()
    
    m_MouseTrackingActive = False
    m_CursorHandle = 0
    m_CaptureCounter = 0
    
    'Note that auto-dropping of severely delayed events is activated by default.  Some UI objects (like the primary canvas)
    ' may disable this behavior under certain conditions.
    m_DelayTrackingActive = True
    m_DelayThreshold = DEFAULT_DELAY_THRESHOLD_MS
    
    'In debug builds, high-resolution mode is disabled by default.  In PD, we typically only use this for paintbrushes;
    ' it will be manually activated by paint tools, as relevant.
    m_HighResModeActive = False
    m_MouseMovePointsRemaining = -1
    
End Sub

Private Sub Class_Terminate()
    Me.Shutdown
End Sub

'Most (all?) of the mouse events this class raises include button modifier information, as a convenience to the caller.  This function
' updates the module-level mouse tracker values on-demand.
Private Sub UpdateMouseModifiers(Optional ByVal lParam As Long = 0)

    'As VB already does with its own mouse events, it is helpful to supply key modifiers directly in the event
    ' params.  Unfortunately, the window messages that report key states will only report SHIFT and CTRL masks,
    ' not ALT.  (This is by design, as most windows use the Alt key to forcibly switch focus to the menu, which
    ' in turn makes it irrelevant for mouse modifications.  However, this behavior can be overridden if the
    ' mouse has been captured by a window, which we may end up doing for PD's canvas, and which is why I'm
    ' adding ALT handling here.)
    
    'Anyway, because the Alt key requires special handling, and some messages don't report button state, I just
    ' ignore window message reports entirely and pull key states manually using GetAsyncKeyState.
    m_retShiftConstants = 0
    If IsVirtualKeyDown(VK_SHIFT) Then m_retShiftConstants = m_retShiftConstants Or vbShiftMask
    If IsVirtualKeyDown(VK_CONTROL) Then m_retShiftConstants = m_retShiftConstants Or vbCtrlMask
        
    'NOTE!  MSDN provides odd instructions for tracking the ALT key (see Remarks here:
    ' http://msdn.microsoft.com/en-us/library/ms646242%28v=vs.85%29.aspx).  They state explicitly to use
    ' use GetKeyState and not GetAsyncKeyState, but they don't provide any rationale for this.  I have no
    ' trouble with GetAsyncKeyState properly reporting Alt status on Win XP through 10, so I'm leaving this strategy
    ' for now, but I remain intrigued by MSDN's oddly specific instructions on the point...
    If IsVirtualKeyDown(VK_ALT) Then m_retShiftConstants = m_retShiftConstants Or vbAltMask
    
    'Similarly, always grab mouse buttons independent of the window message
    m_retMouseButton = 0
    If IsMouseButtonDown(VK_LBUTTON) Then m_retMouseButton = m_retMouseButton Or pdLeftButton
    If IsMouseButtonDown(VK_MBUTTON) Then m_retMouseButton = m_retMouseButton Or pdMiddleButton
    If IsMouseButtonDown(VK_RBUTTON) Then m_retMouseButton = m_retMouseButton Or pdRightButton
    If IsMouseButtonDown(VK_XBUTTON1) Then m_retMouseButton = m_retMouseButton Or pdXButtonOne
    If IsMouseButtonDown(VK_XBUTTON2) Then m_retMouseButton = m_retMouseButton Or pdXButtonTwo
        
    'Some functions also supply mouse button and coordinate values; we'll retrieve these as well.
    ' To improve accuracy of the retrieved coordinates (and make GetMouseMovePointsEx easier to implement),
    ' we manually retrieve coordinates using GetCursorPos, then translate these into the coordinate space
    ' of the target hWnd assigned to this class.
    GetCurrentCursorPosition m_MouseX, m_MouseY
    
    'As a failsafe, crucial mouse events (button down and up, particularly) can never afford to be delayed.
    ' We must retrieve their values at *exactly* the moment the message arrives.  To that end, we always
    ' store copies of these integer-only values, using the values passed in the message parameters.
    m_MouseXInt = GetLoWord(lParam)
    m_MouseYInt = GetHiWord(lParam)
    
End Sub

'Wrapped up with MouseMove events are mouse enter/leave events, as well.
Private Function HandleMouseMove(ByVal uiMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByRef msgEaten As Boolean) As Long
    
    'Grab current mouse and button state.  (The results are stored in module-level variables.)
    UpdateMouseModifiers lParam
    
    'Also retrieve the current message's timestamp.  We'll use this to reconstruct mouse move history, as necessary.
    m_msgTimestamp = GetMessageTime()
    
    'Regardless of button or tracking state, increment the MouseMove counter
    m_MouseMoveCount = m_MouseMoveCount + 1
        
    'Move events always cause us to reset hover mode tracking
    m_HoverModeActive = False

    'If the mouse is moving, that means it has entered the control.  If we aren't already tracking it,
    ' do so now, and raise a MouseEnter event to match.
    If (Not m_MouseTrackingActive) Then
        m_MouseTrackingActive = RequestMouseTrackingForHwnd()
        RaiseEvent MouseEnter(m_retMouseButton, m_retShiftConstants, m_MouseX, m_MouseY)
    End If
        
    'There are two ways to raise MouseMove events for a window.  If high-resolution tracking is active,
    ' we need to reconstruct any missing mouse events and raise them all in turn.  Otherwise, we can
    ' simply return this event.
    If m_HighResModeActive Then
        ReconstructMouseMoveHistory m_retMouseButton, m_retShiftConstants, m_MouseX, m_MouseY, m_msgTimestamp
    Else
        
        'If automatic message dropping is active, make sure this message falls within the acceptable threshold
        ' before firing it.
        If m_DelayTrackingActive Then
            If (GetMessageDelay(m_msgTimestamp) <= m_DelayThreshold) Then
                RaiseEvent MouseMoveCustom(m_retMouseButton, m_retShiftConstants, m_MouseX, m_MouseY, m_msgTimestamp)
            End If
        Else
            RaiseEvent MouseMoveCustom(m_retMouseButton, m_retShiftConstants, m_MouseX, m_MouseY, m_msgTimestamp)
        End If
        
    End If
    
    'Random fact!  If bHandled is set to TRUE, tooltips will not appear for the window in question.
    ' I haven't yet discovered a reason for this, but I can reliably reproduce the issue.  My assumption,
    ' based on the MSDN isntructions of "if an application processes this message, it should return zero,"
    ' is that DefWindowProc needs to pass the message onto the tooltip object so it can process its own
    ' internal timer for showing the tooltip; by setting bHandled to True, we prevent that hand-off, so
    ' the tooltip never gets notification.
    '
    'NOTE: now that PD implements its own custom tooltips, this no longer matters, but I've left the original
    ' handling code as it is "most correct".
    HandleMouseMove = 0
    msgEaten = (Not m_StillUseDefWindowProc)
    
End Function

Private Function HandleMouseDown(ByVal uiMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByRef msgEaten As Boolean) As Long
    
    'Grab current mouse and button state.  (The results are stored in module-level variables.)
    UpdateMouseModifiers lParam
    
    'Also retrieve the current message's timestamp.  We'll use this to reconstruct mouse move history, as necessary.
    m_msgTimestamp = GetMessageTime()
    
    'By default, this function returns 0, but X-buttons have special requirements
    HandleMouseDown = 0
    msgEaten = (Not m_StillUseDefWindowProc)
    
    'Increment the capture counter
    m_CaptureCounter = m_CaptureCounter + 1
    
    'If a tooltip window is active on the target window, kill the tooltip immediately *without a timeout animation*;
    ' this ensures that subsequent mouse events (like _MouseUp) get sent to the correct window.  (Prior to adding
    ' this check, it was possible for the tooltip window to "eat" subsequent mouse actions by accident, because
    ' Windows does weird things with mouse messages over windows that are actively animating to a new state.)
    If UserControls.IsTooltipActive(m_hWnd) Then UserControls.HideUCTooltip True, False
    
    'If we have not captured the mouse, capture it now
    If ((m_CaptureCounter >= 1) And (Not m_DoNotCaptureOnButtonPress) And (Not m_MouseHasBeenCaptured)) Then
        SetCapture m_hWnd
        m_MouseHasBeenCaptured = True
    End If
    
    'Reset the mouse move counter.  (We use this to distinguish "Clicks" from generic MouseDown events.)
    m_MouseMoveCount = 0
    
    'Release any points remaining in the high-resolution input queue
    m_PrevPointStored = True
    With m_LastPointTracked
        .x = m_MouseX And &HFFFF&
        .y = m_MouseY And &HFFFF&
        .ptTime = m_msgTimestamp
    End With
    m_MouseMovePointsRemaining = -1
    
    'Because mouse up/down events tend to focus on the use of a single button, it makes more sense for us to
    ' manually return a single-button type here (rather than a composite flag).  The calling function can
    ' always query this class if it wants more detailed button up/down state data.
    m_retMouseButton = 0
    
    If (uiMsg = WM_LBUTTONDOWN) Then
        m_retMouseButton = vbLeftButton
    ElseIf (uiMsg = WM_RBUTTONDOWN) Then
        m_retMouseButton = vbRightButton
    ElseIf (uiMsg = WM_MBUTTONDOWN) Then
        m_retMouseButton = vbMiddleButton
    ElseIf (uiMsg = WM_XBUTTONDOWN) Then
    
        'In this rare instance, we have to pull data directly from the message, because which XButton triggered
        ' the message can only be found by checking flags in the high-order word of the wParam.
        If ((wParam \ &H10000) = 1) Then
            m_retMouseButton = pdXButtonOne
        Else
            m_retMouseButton = pdXButtonTwo
        End If
        
        'Also strange, X-button events require us to return TRUE!  This only for matters for pre-W2K systems,
        ' but sticking to the rule shows how carefully we read MSDN docs... ;)
        HandleMouseDown = 1
        
    End If
    
    RaiseEvent MouseDownCustom(m_retMouseButton, m_retShiftConstants, m_MouseXInt, m_MouseYInt, m_msgTimestamp)
                
End Function

Private Function HandleMouseUp(ByVal uiMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByRef msgEaten As Boolean) As Long
    
    'Grab current mouse and button state.  (The results are stored in module-level variables.)
    UpdateMouseModifiers lParam
    
    'Also retrieve the current message's timestamp.  We'll use this to reconstruct mouse move history, as necessary.
    m_msgTimestamp = GetMessageTime()
    
    'By default, this function returns 0, but X-buttons have special requirements
    HandleMouseUp = 0
    msgEaten = (Not m_StillUseDefWindowProc)
    
    'Decrement the capture counter; if all buttons have been released, release our capture entirely
    m_CaptureCounter = m_CaptureCounter - 1
    If ((m_CaptureCounter <= 0) And (Not m_DoNotCaptureOnButtonPress) And m_MouseHasBeenCaptured) Then
        m_CaptureCounter = 0
        ReleaseCapture
        m_MouseHasBeenCaptured = False
    End If
                
    'Because mouse up/down events tend to focus on the use of a single button, it makes more sense for us to
    ' manually return a single-button type here (rather than a composite flag).  The calling function can
    ' always query this class if it wants more detailed button up/down state data.
    m_retMouseButton = 0
    
    If (uiMsg = WM_LBUTTONUP) Then
        m_retMouseButton = vbLeftButton
    ElseIf (uiMsg = WM_RBUTTONUP) Then
        m_retMouseButton = vbRightButton
    ElseIf (uiMsg = WM_MBUTTONUP) Then
        m_retMouseButton = vbMiddleButton
    ElseIf (uiMsg = WM_XBUTTONUP) Then
    
        'In this rare instance, we have to pull data directly from the message, because which XButton triggered
        ' the message can only be found by checking flags in the high-order word of the wParam.
        If ((wParam \ &H10000) = 1) Then
            m_retMouseButton = pdXButtonOne
            
            'Because we no longer call DefWindowProc after X-button events, we need to manually raise a
            ' WM_APPCOMMAND message with the relevant back/forward command.
            RaiseEvent AppCommand(AC_BROWSER_BACKWARD, m_retShiftConstants, m_MouseXInt, m_MouseYInt)
        Else
            m_retMouseButton = pdXButtonTwo
            RaiseEvent AppCommand(AC_BROWSER_FORWARD, m_retShiftConstants, m_MouseXInt, m_MouseYInt)
        End If
        
        'Also strange, X-button events require us to return TRUE!  This only for matters for pre-W2K systems,
        ' but sticking to the rule shows how carefully we read MSDN docs... ;)
        HandleMouseUp = 1
        
    End If
    
    'Release any remaining points tracked in the high-resolution input queue
    m_PrevPointStored = False
    m_MouseMovePointsRemaining = -1
                
    'Fun fact!  If PD is busy doing a bunch of processing during MouseMove events (such as drag-resizing a
    ' large layer), the MouseUp event may become artificially delayed.  When it finally triggers, the mouse
    ' may have moved beyond its original point prior to release, causing the MouseUp event to report
    ' inappropriate mouse values.  The best solution to this would be to grab the X/Y coordinates from the
    ' window message, then use GetMouseMovePointsEx to retrieve a high-resolution copy of the coordinates.
    ' Until we implement GetMouseMovePointsEx as a pdInput-wide solution, however, I'm going to simply
    ' take the x/y values from the message and report them for now.
    '
    '(The point of that very long message is to explain why we use mouseXInt and mouseYInt, below.)
    
    'If the user has not moved the mouse an appreciable amount since the last MouseDown event, we'll also
    ' call this a "click" and raise the corresponding Click() event.  (Note that a MouseUp event is raised
    ' either way, and if a Click event is also being raised, we warn the user in the MouseUp event.
    ' This way, they can choose to abandon MouseUp processing if the _Click event is implemented.)
    If (m_MouseMoveCount < 3) Then
        RaiseEvent MouseUpCustom(m_retMouseButton, m_retShiftConstants, m_MouseXInt, m_MouseYInt, True, m_msgTimestamp)
        RaiseEvent ClickCustom(m_retMouseButton, m_retShiftConstants, m_MouseXInt, m_MouseYInt)
    
    'The mouse has moved too much, so a Click event will not be raised.  Only raise a MouseUp event.
    Else
        RaiseEvent MouseUpCustom(m_retMouseButton, m_retShiftConstants, m_MouseXInt, m_MouseYInt, False, m_msgTimestamp)
    End If

End Function

Private Function HandleDblClick(ByVal uiMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByRef msgEaten As Boolean) As Long
    
    'Grab current mouse and button state.  (The results are stored in module-level variables.)
    UpdateMouseModifiers lParam
    
    'Also retrieve the current message's timestamp.  We'll use this to reconstruct mouse move history, as necessary.
    m_msgTimestamp = GetMessageTime()
    
    'By default, this function returns 0, but X-buttons have special requirements
    HandleDblClick = 0
    msgEaten = (Not m_StillUseDefWindowProc)
    
    'Because double-click events tend to focus on the use of a single button, it makes more sense for us to
    ' manually return a single-button type here (rather than a composite flag).
    m_retMouseButton = 0
    
    If (uiMsg = WM_LBUTTONDBLCLK) Then
        m_retMouseButton = vbLeftButton
    ElseIf (uiMsg = WM_RBUTTONDBLCLK) Then
        m_retMouseButton = vbRightButton
    ElseIf (uiMsg = WM_MBUTTONDBLCLK) Then
        m_retMouseButton = vbMiddleButton
    ElseIf (uiMsg = WM_XBUTTONDBLCLK) Then
    
        'In this rare instance, we have to pull data directly from the message, because which XButton triggered
        ' the message can only be found by checking flags in the high-order word of the wParam.
        If ((wParam \ &H10000) = 1) Then
            m_retMouseButton = pdXButtonOne
        Else
            m_retMouseButton = pdXButtonTwo
        End If
        
        'Also strange, X-button events require us to return TRUE!  This only for matters for pre-W2K systems,
        ' but sticking to the rule shows how carefully we read MSDN docs... ;)
        HandleDblClick = 1
        
    End If
    
    'When handled by Windows, double-click messages do not also raise a corresponding WM_*BUTTONDOWN msg.
    ' (See https://docs.microsoft.com/en-us/windows/win32/inputdev/wm-lbuttondblclk)
    
    'To ensure that fast repeat clicks are handled correctly, we manually raise a corresponding
    ' _MouseDown event.  (Note that Windows *will* raise a corresponding WM_*BUTTONUP msg, however.)
    
    'Thank you to wqweto and https://github.com/tannerhelland/PhotoDemon/pull/283 for first catching this.
    RaiseEvent MouseDownCustom(m_retMouseButton, m_retShiftConstants, m_MouseX, m_MouseY, m_msgTimestamp)
    RaiseEvent DoubleClickCustom(m_retMouseButton, m_retShiftConstants, m_MouseX, m_MouseY)
    
End Function

Private Function HandleMouseLeave(ByVal uiMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByRef msgEaten As Boolean) As Long
    
    'Grab current mouse and button state.  (The results are stored in module-level variables.)
    UpdateMouseModifiers lParam
    
    'Also retrieve the current message's timestamp.  We'll use this to reconstruct mouse move history, as necessary.
    m_msgTimestamp = GetMessageTime()
    
    'Note that the default wndproc is *not* called after us
    HandleMouseLeave = 0
    msgEaten = True
    
    'If we've been tracking mouse movements for this hWnd, raise a MouseLeave event now
    If m_MouseTrackingActive Then
        
        m_PrevPointStored = False
        m_MouseTrackingActive = False
        m_LastHoverX = -1
        m_LastHoverY = -1
        
        RaiseEvent MouseLeave(m_retMouseButton, m_retShiftConstants, m_MouseX, m_MouseY)
        
    End If
    
End Function

Private Function HandleMouseHover(ByVal uiMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByRef msgEaten As Boolean) As Long
    
    'Grab current mouse and button state.  (The results are stored in module-level variables.)
    UpdateMouseModifiers lParam
    
    'Also retrieve the current message's timestamp.  We'll use this to reconstruct mouse move history, as necessary.
    m_msgTimestamp = GetMessageTime()
    
    'Note that the default wndproc is *not* called after us
    HandleMouseHover = 0
    msgEaten = True
    
    'Only process hover events if we care about tracking mouse movements for this hWnd
    If m_MouseTrackingActive And (Not m_HoverModeActive) Then
        
        'Note that we are currently in hover mode; this won't be reset until some kind of MouseMove occurs
        m_HoverModeActive = True
        
        'Note the current mouse position.  Windows may mistakenly issue multiple hover events, but we only
        ' want to raise one hover event at a time.
        
        'Note that we don't change the tracking state, because we don't want to generate subsequent
        ' MouseEnter events after this one!  Instead, silently re-request mouse tracking for this hWnd.
        RequestMouseTrackingForHwnd
        
        'Unlike Enter/Leave events, we want to raise hover events regardless of the hWnd; the assumption
        ' here is that the target window will always want these events, and if it doesn't, it can do its
        ' own bounds-checking to determine if a hover event is relevant or not.
        If ((m_LastHoverX <> m_MouseX) Or (m_LastHoverY <> m_MouseY)) Then
            m_LastHoverX = m_MouseX
            m_LastHoverY = m_MouseY
            RaiseEvent MouseHover(m_retMouseButton, m_retShiftConstants, m_MouseX, m_MouseY)
        End If
        
    End If
    
End Function

Private Function HandleMouseWheel(ByVal uiMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByRef msgEaten As Boolean) As Long
    
    'Grab current mouse and button state.  (The results are stored in module-level variables.)
    UpdateMouseModifiers lParam
    
    'Also retrieve the current message's timestamp.  We'll use this to reconstruct mouse move history, as necessary.
    m_msgTimestamp = GetMessageTime()
    
    'Note that the default wndproc is *not* called after us
    HandleMouseWheel = 1
    msgEaten = True
    
    'First things first: retrieve the high-word, which contains the change (delta) in mousewheel position
    m_scrollDelta = wParam \ &H10000
    
    'Next, use the user's scroll wheel setting (set via the Control Panel) to calculate a final scroll amount,
    ' in lines or chars depending on whether vertical or horizontal scrolling is active.
    
    'Vertical scroll only, meaning Vertical Wheel + !Shift and !Ctrl
    If (uiMsg = WM_MOUSEWHEEL) And ((m_retShiftConstants And vbShiftMask) = 0) And ((m_retShiftConstants And vbCtrlMask) = 0) Then
        m_finalScrollAmount = CDbl(m_scrollDelta) * GetUserScrollSetting(True)
        RaiseEvent MouseWheelVertical(m_retMouseButton, m_retShiftConstants, m_MouseX, m_MouseY, m_finalScrollAmount)
    
    'Horizontal scroll, but generated via Vertical Wheel + Shift and !Ctrl
    ElseIf (uiMsg = WM_MOUSEWHEEL) And ((m_retShiftConstants And vbShiftMask) <> 0) And ((m_retShiftConstants And vbCtrlMask) = 0) Then
    
        m_finalScrollAmount = CDbl(m_scrollDelta) * GetUserScrollSetting(True)
        
        'Because the user is using the Shift+VerticalWheel combination, reverse the delta; this makes it so
        ' that shift+up_wheel = left, and shift+down_wheel = right - the idea is that up_wheel and shift+up_wheel
        ' both target the top-left corner of the image.
        m_finalScrollAmount = -1 * m_finalScrollAmount
        RaiseEvent MouseWheelHorizontal(m_retMouseButton, m_retShiftConstants, m_MouseX, m_MouseY, m_finalScrollAmount)
        
    'Zoom scroll, generated via Vertical Wheel + Ctrl and !Shift
    ElseIf (uiMsg = WM_MOUSEWHEEL) And ((m_retShiftConstants And vbCtrlMask) <> 0) And ((m_retShiftConstants And vbShiftMask) = 0) Then
        m_finalScrollAmount = CDbl(m_scrollDelta) * GetUserScrollSetting(True)
        RaiseEvent MouseWheelZoom(m_retMouseButton, m_retShiftConstants, m_MouseX, m_MouseY, m_finalScrollAmount)
    
    'Horizontal scroll, generated via actual Horizontal Scroll/Tilt message
    ElseIf (uiMsg = WM_MOUSEHWHEEL) Then
        m_finalScrollAmount = CDbl(m_scrollDelta) * GetUserScrollSetting(False)
        RaiseEvent MouseWheelHorizontal(m_retMouseButton, m_retShiftConstants, m_MouseX, m_MouseY, m_finalScrollAmount)
        
    'The only combination left is Vertical scroll, but with some weird combination of Shift modifiers; return
    ' a generic vertical wheel event as-is.
    Else
        m_finalScrollAmount = CDbl(m_scrollDelta) * GetUserScrollSetting(True)
        RaiseEvent MouseWheelVertical(m_retMouseButton, m_retShiftConstants, m_MouseX, m_MouseY, m_finalScrollAmount)
    
    End If
    
End Function

'App commands are meta functions like "back" or "increase volume" or "media rewind".  They can be triggered
' by a variety of actions, depending on system configuration: mouse gestures, dedicated keyboard keys,
' or gestures on a touch screen.
Private Function HandleAppCommand(ByVal uiMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByRef msgEaten As Boolean) As Long
    
    'NOTE: PD does not currently care about updating stored mouse state and message time
    ' for app-command-generated events.  I may revisit this in the future if it proves useful
    ' to painting interaction(s).
    'Grab current mouse and button state.  (The results are stored in module-level variables.)
    'UpdateMouseModifiers lParam
    'Also retrieve the current message's timestamp.  We'll use this to reconstruct mouse move history, as necessary.
    'm_msgTimestamp = GetMessageTime()
    
    'Note that the default wndproc is *not* called after us
    HandleAppCommand = 1
    msgEaten = True
    
    'At present, we don't process additional device or key parameters - just the action ID
    ParseAppCommand lParam, m_commandID
    RaiseEvent AppCommand(m_commandID, m_retShiftConstants, m_MouseX, m_MouseY)

End Function

Private Function HandleSetCursor(ByVal uiMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByRef msgEaten As Boolean) As Long
    ApplyCursor
    msgEaten = True
    HandleSetCursor = 1
End Function
            
Private Function ISubclass_WindowMsg(ByVal hWnd As Long, ByVal uiMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal dwRefData As Long) As Long
    
    'Because this class subclasses such an enormous list of messages, it is difficult to know whether default wndprocs still need to
    ' be called after we complete our business.  As such, child subroutines will fill this value to let us know if they've already
    ' called the default wndproc; if they haven't, we'll do it for them.
    Dim msgEaten As Boolean: msgEaten = False
    
    If (uiMsg = WM_MOUSEMOVE) Then
        ISubclass_WindowMsg = HandleMouseMove(uiMsg, wParam, lParam, msgEaten)
        
    ElseIf ((uiMsg = WM_LBUTTONDOWN) Or (uiMsg = WM_MBUTTONDOWN) Or (uiMsg = WM_RBUTTONDOWN) Or (uiMsg = WM_XBUTTONDOWN)) Then
        ISubclass_WindowMsg = HandleMouseDown(uiMsg, wParam, lParam, msgEaten)
        
    ElseIf ((uiMsg = WM_LBUTTONUP) Or (uiMsg = WM_MBUTTONUP) Or (uiMsg = WM_RBUTTONUP) Or (uiMsg = WM_XBUTTONUP)) Then
        ISubclass_WindowMsg = HandleMouseUp(uiMsg, wParam, lParam, msgEaten)
    
    ElseIf ((uiMsg = WM_LBUTTONDBLCLK) Or (uiMsg = WM_MBUTTONDBLCLK) Or (uiMsg = WM_RBUTTONDBLCLK) Or (uiMsg = WM_XBUTTONDBLCLK)) Then
        ISubclass_WindowMsg = HandleDblClick(uiMsg, wParam, lParam, msgEaten)
    
    ElseIf (uiMsg = WM_MOUSELEAVE) Then
        ISubclass_WindowMsg = HandleMouseLeave(uiMsg, wParam, lParam, msgEaten)
    
    ElseIf (uiMsg = WM_MOUSEHOVER) Then
        ISubclass_WindowMsg = HandleMouseHover(uiMsg, wParam, lParam, msgEaten)
    
    ElseIf ((uiMsg = WM_MOUSEWHEEL) Or (uiMsg = WM_MOUSEHWHEEL)) Then
        If (Not m_DisregardWheelEvents) Then ISubclass_WindowMsg = HandleMouseWheel(uiMsg, wParam, lParam, msgEaten)
    
    ElseIf (uiMsg = WM_APPCOMMAND) Then
        If m_TrackAppCommands Then ISubclass_WindowMsg = HandleAppCommand(uiMsg, wParam, lParam, msgEaten)
    
    ElseIf (uiMsg = WM_SETCURSOR) Then
        ISubclass_WindowMsg = HandleSetCursor(uiMsg, wParam, lParam, msgEaten)
    
    'Failsafe check for parent window destruction
    ElseIf (uiMsg = WM_NCDESTROY) Then
        VBHacks.StopSubclassing hWnd, Me
        m_hWnd = 0
        
    End If
    
    If (Not msgEaten) Then ISubclass_WindowMsg = VBHacks.DefaultSubclassProc(hWnd, uiMsg, wParam, lParam)
    
End Function
