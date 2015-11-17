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
'Copyright 2014-2015 by Tanner Helland
'Created: 27/May/14 (though many individual parts existed earlier than this!)
'Last updated: 02/October/14
'Last update: implement high-resolution mouse tracking via GetMouseMovePointsEx and enable by default.  Also, implement
'              auto-dropping of messages that are delayed beyond a certain threshold (100 ms by default).
'
'As I get ever closer to implementing paint tools in PhotoDemon, the need for a more comprehensive input solution has
' become clear.  Until now, a combination of intrinsic VB mouse events and a few extra subclassing bits (e.g. mousewheel)
' has covered PD pretty well, but drawing tools and more advanced UI features require more detailed input handling, like
' GetMouseMovePointsEx + spline interpolation for buttery smooth mouse input, or (ideally) some level of support for
' touch and/or pen input.
'
'Because I would like to keep all input code as unified as possible, it is finally time to start merging PD's many
' disparate pieces of input handling into this single class.  Besides the obvious mouse events that require special
' handling (again, mousewheel), I'm also working on moving PD's (large!) collection of custom mouse cursor code into
' this class, which should reduce code redundancy while improving reliability.
'
'Many thanks to Kroc Camen of camendesign.com, whose bluMouseEvents class served PD well for a long time prior to this
' implementation.  If you need a lightweight mouse-handler that works well as a standalone solution, I recommend using
' that instead of this very PD-specific class.  You can download a copy here (link good as of May '14):
' https://github.com/Kroc/MaSS1VE/tree/master/CODE/Blu
'
'A few important notes when using this class:
'
'- This class can optionally handle standard mouse events for a given hWnd (e.g. mouse events VB normally covers, like
'   Click, DoubleClick, MouseMove, etc).  There are a lot of reasons to do this, but among the obvious benefits are things
'   like x/y coordinates for click events, support for 32-bit mouse positions, support for X-keys as a button type, and more.
'   In the future, the custom version of these events could also do things like supply a Pressure parameter for MouseDown
'   events, or use interpolation to smooth out MouseMove coordinate reports.
'
'- Multiple hWnds can be passed to a single instace of this class, but the first hWnd passed serves as the "master hWnd".
'   This means that all mouse coordinates reported by this class will always be relative to *that hWnd*.  Tracking multiple
'   hWnds at once is helpful if you want a child object to receive all mousewheel events for it and its parent, for example,
'   or if you want to handle all mouse events for a collection of objects with a single MouseMove function (e.g., in a user
'   control).  Similarly, if you want to subclass multiple controls within a single parent, that should work just fine, as
'   long as you don't assign the parent hWnd to multiple instances of this class (e.g. create just ONE pdInput copy for the
'   entire collection, rather than different ones for each child).
'
'- As implied above, mouse coordinate, button, and shift key modifiers are processed independent of window messages.  This
'   allows us to supply such data even if a given window message doesn't automatically include them.  It also allows for
'   higher accuracy when tracking mouse move data, as we use 32-bit values instead of 16-bit ones (which seems ridiculous
'   now, but may not be forever).
'
'- As a bonus, this class can also (optionally) handle WM_APP_COMMAND messages.  This is useful for handling some buttons that
'   may appear on a mouse, such as back/forward, instead of using X-button messages (which may not be accurate, if the user has
'   remapped functionality in a non-standard way).
'
'Also, let me give a big thank you to Steve McMahon for translating the GET_APPCOMMAND_LPARAM, GET_DEVICE_LPARAM, and
' GET_KEYSTATE_LPARAM macros into VB; this class uses these events to respond to actions like "back"/"forward", regardless
' of input source.  You can see Steve's original work on WM_APPCOMMAND messages here:
' http://www.vbaccelerator.com/home/VB/Tips/Responding_to_AppCommands/article.asp
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'This class can raise many events.  The following events are raised for all tracked hWnds:
Public Event MouseLeave(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
Public Event MouseEnter(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
Public Event MouseHover(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
Public Event MouseWheelVertical(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal scrollAmount As Double)
Public Event MouseWheelHorizontal(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal scrollAmount As Double)
Public Event MouseWheelZoom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal zoomAmount As Double)
Public Event ClickCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
Public Event DoubleClickCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
Public Event MouseDownCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
Public Event MouseUpCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal ClickEventAlsoFiring As Boolean)

'The following events are raised only if specifically requested in the addInputTracker function
Public Event MouseMoveCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
Public Event AppCommand(ByVal cmdID As AppCommandConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)

'This class tracks all kinds of window messages.  Most objects won't need this level of tracking, but they're there
' if needed.
Private Const WM_MOUSEHWHEEL As Long = &H20E
Private Const WM_MOUSEWHEEL As Long = &H20A
Private Const WM_MOUSEHOVER As Long = &H2A1
Private Const WM_MOUSELEAVE As Long = &H2A3
Private Const WM_MOUSEMOVE As Long = &H200

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

'X buttons (sometimes called buttons 4 and 5 in MSDN docs) are typically used for forward/back maneuvering.
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

Private Const TME_HOVER As Long = &H1
Private Const TME_LEAVE As Long = &H2

Private Declare Function TrackMouseEvent Lib "user32" (ByRef lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long

'The Alt mask won't be returned by mouse WM notifications, so we need to retrieve it manually
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

'Virtual key-codes currently supported by pdInput
Private Const VK_SHIFT As Long = &H10
Private Const VK_CONTROL As Long = &H11
Private Const VK_ALT As Long = &H12    'Note that VK_ALT is referred to as VK_MENU in MSDN documentation!

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
' handle it, using GetSystemMetrics (which is publicly declared in PD).
Private Const SM_SWAPBUTTON As Long = 23

'Mouse capturing provides more predictable behavior, particularly on UCs
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long

'This class also handles cursor management, but because cursors can be requested by external functions, they are declared
' publicly in the Icon and Cursor module.  Look there for cursor-related API constants.

'Retrieve the current cursor position, in screen coordinates
Private Declare Function GetCursorPos Lib "user32" (ByRef lpPoint As POINTAPI) As Long

'Set a new cursor for a given class
Private Declare Function GetCursor Lib "user32" () As Long
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function GetClassLong Lib "user32" Alias "GetClassLongW" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetClassLong Lib "user32" Alias "SetClassLongW" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Const GCL_HCURSOR = (-12)

'Reconstruct a full mouse movement history
Private Type MOUSEMOVEPOINT
    x As Long
    y As Long
    ptTime As Long
    dwExtraInfo As Long
End Type

Private Enum mouseResolution
    GMMP_USE_DISPLAY_POINTS = 1
    GMMP_USE_HIGH_RESOLUTION_POINTS = 2
End Enum

#If False Then
    Const GMMP_USE_DISPLAY_POINTS = 1, GMMP_USE_HIGH_RESOLUTION_POINTS = 2
#End If

Private Declare Function GetMouseMovePointsEx Lib "user32" (ByVal sizeOfMouseMoveStruct As Long, ByRef currentMouseMovePoint As MOUSEMOVEPOINT, ByVal ptrToMouseMovePointArray As Long, ByVal numOfPointsToRetrieve As Long, ByVal resolutionFlag As mouseResolution) As Long
Private Declare Function GetMessageTime Lib "user32" () As Long
Private Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

'Weird legacy desktop issues require us to transform points passed to/from GetMouseMovePointsEx, relative to the current monitor system
Private Const SM_XVIRTUALSCREEN = 76
Private Const SM_YVIRTUALSCREEN = 77
Private Const SM_CXVIRTUALSCREEN = 78
Private Const SM_CYVIRTUALSCREEN = 79

'API helper functions for converting between screen and client coordinate spaces
Private Declare Function ClientToScreen Lib "user32" (ByVal hndWindow As Long, ByRef lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hndWindow As Long, ByRef scrPoint As POINTAPI) As Long

'MSDN best practices suggest that we always retrieve the user's setting for scroll wheel sensitivity, as it may change
' while a program is running.  SystemParametersInfo is used for this.  Note that horizontal and vertical settings
' are stored separately.
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uiAction As Long, ByVal uiParam As Long, ByRef pvParam As Long, ByVal fWinIni As Long) As Long
Private Const SPI_GETWHEELSCROLLLINES As Long = &H68
Private Const SPI_GETWHEELSCROLLCHARS As Long = &H6C

'Master subclasser for all input actions
Private cSubclass As cSelfSubHookCallback

'hWnds to track.  At least one entry (position 0) is guaranteed present, but the user may supply more than one.
' When shutting down the class, make sure to unsubclass all entries in the array.
Private m_TrackingData() As Long, m_hWndCount As Long

'Mouse tracking is active for a given hWnd
Private m_MouseTrackingActive() As Boolean

'To prevent multiple hover events from being raised by multiple hWnds (or one reaaaally long hover), we stop raising
' hover events until the mouse moves again
Private m_HoverModeActive As Boolean

'If the user has requested that we force a persistent cursor for this object, the cursor handle will be stored here.
Private m_CursorHandle As Long, m_PrevClassCursorHandle As Long

'If the user wants us to raise additional events using this class - e.g. MouseMove or AppCommands - we will use
' this enum with the multidimensional tracking array to handle events accordingly.
Private Enum PD_CUSTOM_INPUT_TRACKERS
    PDIT_DEFAULT = 0
    PDIT_MOUSE_ENTERLEAVE = 1
    PDIT_STANDARD_MOUSE = 2
    PDIT_APPCOMMAND = 3
    PDIT_CURSOR = 4
End Enum

#If False Then
    Private Const PDIT_DEFAULT = 0, PDIT_MOUSE_ENTERLEAVE = 1, PDIT_STANDARD_MOUSE = 2, PDIT_APPCOMMAND = 3, PDIT_CURSOR = 4
#End If

Private Const NUM_OF_CUSTOM_INPUT_TRACKERS = 5

'This class will generate Click() events using its own algorithm for determining when an action constitutes a "click".
' Basically, a MouseDown/Up combination with two or less MouseMove messages between them constitutes a Click.  This variable
' is used to count MouseMove occurrences.
Private m_MouseMoveCount As Long

'In rare cases (e.g. the layer toolbox), a parent form and child control may have separate cursor tracking activated in two different
' instances of this class.  When this happens, we don't want to automatically handle WM_SETCURSOR messages; otherwise, the parent's
' cursor setting will always override the child.  In this rare instance, this override value will be set to TRUE to notify the class
' that WM_SETCURSOR can be ignored for the parent object.  (Note that this behavior must be manually set via setCursorOverrideState.)
Private m_OverrideSetCursorMessages As Boolean

'While we can't track mouse events any faster than the WM_MOUSEMOVE messages we receive, we can reconstruct missed mouse events by
' using GetMouseMovePointsEx.  To activate this feature, pass TRUE to the setHighResolutionTrackingMode sub, which will in turn
' activate this boolean.  If active, the subclasser will reconstruct any missing mouse events history, and raise each missed event
' in turn.
Private m_HighResModeActive As Boolean

'To reconstruct missing mouse movement points, we must track the last returned mouse point.  Otherwise, we risk repeating points.
Private m_PrevPointStored As Boolean
Private m_LastPointTracked As MOUSEMOVEPOINT

'Because PD sometimes performs energy-intensive actions within mouse events, it is possible for raised messages to become
' increasingly delayed, especially if high-resolution tracking is active.  In an attempt to mitigate this, this class supports
' auto-dropping of messages delayed beyond a certain threshold.  NOTE: this behavior only affects MouseMove events.  All other
' mouse events will be fired regardless of delay.
Private m_DelayTrackingActive As Boolean
Private m_DelayThreshold As Long
Private Const DEFAULT_DELAY_THRESHOLD_MS = 120

'To improve mouse behavior on UCs, this class will auto-capture the mouse when a button is pressed.  Each press of a button increments
' this internal capture counter.  When a button is released, the capture counter is decremented.  When all buttons have been released,
' the mouse capture is released.
Private m_CaptureCounter As Long

'In some cases, the users may want to deal with mouse events only under certain circumstances.  If they test an event and realize that it's
' useless to them, they can set this parameter to TRUE via the ignoreLastMessage() sub.  Upon returning from an event, that parameter
' notifies the subclasser to not mark the event as handled, so DefWindowProc can have a swing at it.
Private m_IgnoreLastMessage As Boolean

'By default, PD forcibly captures the mouse when a button is pressed.  To disable this behavior, set this value to TRUE.
Private m_DoNotCaptureOnButtonPress As Boolean

'If tracking multiple hWnds, the caller may find it helpful to retrieve the relevant hWnd for a given event.  We track the last-handled
' hwnd here.
Private m_LastMsgHWnd As Long

'If the mouse is currently captured, this will be set to TRUE
Private m_MouseHasBeenCaptured As Boolean

Public Function getLastHwnd() As Long
    getLastHwnd = m_LastMsgHWnd
End Function

Public Sub setCaptureOverride(ByVal newCaptureMode As Boolean)
    
    m_DoNotCaptureOnButtonPress = newCaptureMode
    
    'Also, release a current capture, if any
    If newCaptureMode And m_MouseHasBeenCaptured Then
        m_CaptureCounter = 0
        ReleaseCapture
        m_MouseHasBeenCaptured = False
    End If
    
End Sub

'For details, see the declaration for m_IgnoreLastMessage
Public Sub ignoreLastMessage()
    m_IgnoreLastMessage = True
End Sub

'Use this function enable/disable automatic dropping of severely delayed events
Public Sub setAutoDropOfDelayedEvents(ByVal newMode As Boolean)
    m_DelayTrackingActive = newMode
End Sub

'Use this function to set the threshold, in milliseconds, after which a message is considered unacceptably delayed.
' NOTE: system timer resolution limits this value to roughly [10, 16] ms, so setting a threshold below that point may cause
'       nearly all events to be dropped.  Try to limit the requested value to > 100 ms for best results.
Public Sub setAutoDropDelayThreshold(ByVal newThreshold As Long)
    m_DelayThreshold = newThreshold
    If m_DelayThreshold < 16 Then m_DelayThreshold = 16
End Sub

'Use this function to enable/disable high-resolution tracking
Public Sub setHighResolutionTrackingMode(ByVal newMode As Boolean)
    m_HighResModeActive = newMode
End Sub

Public Sub setCursorOverrideState(ByVal newOverride As Boolean)
    m_OverrideSetCursorMessages = newOverride
End Sub

Private Sub Class_Initialize()
    
    'Reset all input tracking variables
    ReDim m_TrackingData(0 To NUM_OF_CUSTOM_INPUT_TRACKERS - 1, 0) As Long
    ReDim m_MouseTrackingActive(0) As Boolean
    
    'Note that we are not currently tracking the mouse for this object
    m_MouseTrackingActive(0) = False
    
    'Initialize the subclasser
    Set cSubclass = New cSelfSubHookCallback
    
    'Reset the cursor handle for this instance
    m_CursorHandle = 0
    
    'Auto-dropping of severely delayed events is activated by default
    m_DelayTrackingActive = True
    m_DelayThreshold = DEFAULT_DELAY_THRESHOLD_MS
    
    'Debugging high-resolution mode is made easier by enabling it automatically.
    m_HighResModeActive = True
    
    'Reset the capture counter
    m_CaptureCounter = 0
    
End Sub

'Initiate input tracking of a given hWnd.  Aside from the required hWnd, a few other, optional parameters exist:
'
' - Whether to track MouseMove events as well.  If this is specified, note that VB's intrinsic "MouseMove" event
'    will not fire, as this function will process the messages and raise its own event, preventing VB from doing so.
'
' - Whether to track AppCommand messages for this hWnd.  This is used for "virtual" functions like forward/back,
'    which are typically generated by mouse buttons 4 and 5, but could also be generated by a multimedia keyboard
'    or gestures on a tablet PC.
'
'BY DEFAULT, all added hWnds will have MouseEnter/Leave/Wheel events tracked "for free".
'
'Note that this function will return TRUE if input tracking was initiated successfully.
Public Function addInputTracker(ByVal targetHwnd As Long, Optional ByVal alsoTrackMouseEnterLeave As Boolean = False, Optional ByVal alsoTrackStandardMouseEvents As Boolean = False, Optional ByVal alsoTrackAppCommands As Boolean = False, Optional ByVal alsoHandleCursorDuties As Boolean = False, Optional ByVal disregardWheelEvents As Boolean = False) As Boolean
        
    'Add the hWnd to our tracking array.  Note that the hWnd in position 0 is treated as the base.
    If targetHwnd <> 0 Then
    
        'Add the hWnd to our tracking array, and increment all subarrays accordingly
        m_hWndCount = m_hWndCount + 1
        ReDim Preserve m_TrackingData(0 To NUM_OF_CUSTOM_INPUT_TRACKERS - 1, 0 To m_hWndCount - 1) As Long
        ReDim Preserve m_MouseTrackingActive(0 To m_hWndCount - 1) As Boolean
        
        m_TrackingData(PDIT_DEFAULT, m_hWndCount - 1) = targetHwnd
        m_TrackingData(PDIT_MOUSE_ENTERLEAVE, m_hWndCount - 1) = alsoTrackMouseEnterLeave
        m_TrackingData(PDIT_STANDARD_MOUSE, m_hWndCount - 1) = alsoTrackStandardMouseEvents
        m_TrackingData(PDIT_APPCOMMAND, m_hWndCount - 1) = alsoTrackAppCommands
        m_TrackingData(PDIT_CURSOR, m_hWndCount - 1) = alsoHandleCursorDuties
    
    'The calling function didn't supply an hWnd.  Warn and exit.
    Else
        Debug.Print "WARNING: a function just requested input tracking, but didn't supply an hWnd.  Fix it!"
        addInputTracker = False
        Exit Function
    End If
    
    'Attach the subclasser to this hWnd
    If Not cSubclass.ssc_Subclass(targetHwnd, , 1, Me) Then
        
        Debug.Print "WARNING: input subclasser failed to subclass hWnd #" & targetHwnd & ".  Fix it!"
        addInputTracker = False
        Exit Function
    
    End If
        
    'Subclassing was successful.  Add all relevant mouse messages.
    ' (Note that we subclass MouseMove events even if the user doesn't want these events returned; this is
    ' used to generate MouseEnter/Leave events.)
    If Not disregardWheelEvents Then cSubclass.ssc_AddMsg targetHwnd, MSG_BEFORE, WM_MOUSEWHEEL, WM_MOUSEHWHEEL
    
    'Set any optional event subclassing as well
    If alsoTrackMouseEnterLeave Then cSubclass.ssc_AddMsg targetHwnd, MSG_BEFORE, WM_MOUSELEAVE, WM_MOUSEMOVE, WM_MOUSEHOVER
    If alsoTrackAppCommands Then cSubclass.ssc_AddMsg targetHwnd, MSG_BEFORE, WM_APPCOMMAND
    If alsoTrackStandardMouseEvents Then
        
        'Mouse down...
        cSubclass.ssc_AddMsg targetHwnd, MSG_BEFORE, WM_LBUTTONDOWN, WM_MBUTTONDOWN, WM_RBUTTONDOWN, WM_XBUTTONDOWN
        
        'Mouse up...
        cSubclass.ssc_AddMsg targetHwnd, MSG_BEFORE, WM_LBUTTONUP, WM_MBUTTONUP, WM_RBUTTONUP, WM_XBUTTONUP
        
        'Double-clicks (including X-buttons)...
        cSubclass.ssc_AddMsg targetHwnd, MSG_BEFORE, WM_LBUTTONDBLCLK, WM_MBUTTONDBLCLK, WM_RBUTTONDBLCLK, WM_XBUTTONDBLCLK
        
    End If
    
    'Cursor duties require response to WM_SETCURSOR
    If alsoHandleCursorDuties Then cSubclass.ssc_AddMsg targetHwnd, MSG_BEFORE, WM_SETCURSOR
            
    'If enter, leave, and hover events are desired, start mouse tracking now
    If alsoTrackMouseEnterLeave Then
    
        If requestMouseTrackingForHwnd(targetHwnd) Then
            m_MouseTrackingActive(m_hWndCount - 1) = True
        Else
            
            Debug.Print "WARNING: couldn't get Windows to track mouse events for hWnd #" & targetHwnd & ".  Fix it!"
            addInputTracker = False
            Exit Function
                
        End If
        
    End If
    
    addInputTracker = True
    
End Function

'Assign a cursor to the hWnd collection.  Call this function without a parameter to reset the cursor to the default arrow.
Public Sub setSystemCursor(Optional ByVal systemCursorType As SystemCursorConstant = IDC_DEFAULT, Optional ByVal confirmHwnd As Long = 0)
    
    'Ignore requests while an override is active
    If Not m_OverrideSetCursorMessages Then
    
        'Load the relevant cursor handle
        Dim tmpCursorHandle As Long
        
        If systemCursorType = IDC_DEFAULT Then
            tmpCursorHandle = LoadCursor(0, IDC_ARROW)
        Else
            tmpCursorHandle = LoadCursor(0, systemCursorType)
        End If
        
        'Check for duplicate cursor requests, and ignore them as necessary
        If (GetCursor() <> tmpCursorHandle) Then
            
            m_CursorHandle = tmpCursorHandle
            
            If confirmHwnd = 0 Then
                applyCursorToHwnds False
            Else
                If getIndexFromHwnd(confirmHwnd) >= 0 Then
                    applyCursorToHwnds False, getIndexFromHwnd(confirmHwnd)
                Else
                    applyCursorToHwnds False
                End If
            End If
            
        End If
        
    End If
        
End Sub

'Assign a PNG cursor to the hWnd collection.  This function leans on the Icon and Cursor module to handle the actual
' retrieval, decompression, and assembly of the cursor.  We handle it there so that requested icons can be cached;
' there is a high probability of PNG icons being requested by more than one class in a given session, so rather than
' repeat all the steps for each request, we only do it once, then cache the results.  PD itself will handle unloading
' custom cursors at exit time, so this class need not concern itself with unloading requested cursors.
Public Sub setPNGCursor(ByVal pngResourceName As String, Optional ByVal cursorHotspotX As Long = 0, Optional ByVal cursorHotspotY As Long = 0)
    
    If (Not m_OverrideSetCursorMessages) Then
    
        Dim tmpCursorHandle As Long
        tmpCursorHandle = Icon_and_Cursor_Handler.requestCustomCursor(pngResourceName, cursorHotspotX, cursorHotspotY)
        
        If m_CursorHandle <> tmpCursorHandle Then
            m_CursorHandle = tmpCursorHandle
            applyCursorToHwnds False
        End If
        
    End If
    
End Sub

'Private helper function used by setSystemCursor and setPNGCursor, above.  When setting the cursor for the first time
' for our base hWnd, it is helpful to reset the class cursor definition so that Windows doesn't automatically raise
' WM_SETCURSOR messages every time the mouse is moved.  Set the optional parameter to TRUE to cache the existing class
' cursor and overwrite it with NULL.
Private Sub applyCursorToHwnds(Optional ByVal removeExistingClassCursor As Boolean = False, Optional ByVal currentIndex As Long = 0)
    
    If (Not m_OverrideSetCursorMessages) Then
    
        'If removeExistingClassCursor is set, use SetClassLong to remove any existing class-level cursor definition
        If removeExistingClassCursor Then m_PrevClassCursorHandle = SetClassLong(m_TrackingData(PDIT_DEFAULT, currentIndex), GCL_HCURSOR, 0&)
        
        'Apply the new cursor using SetCursor
        SetCursor m_CursorHandle
        
    End If

End Sub

'Partner to applyCursorToHwnds, above.  To be used only when the mouse leaves a cursor-controlled area.
Private Sub resetClassCursor(Optional ByVal currentIndex As Long = 0)
    
    If (Not m_OverrideSetCursorMessages) Then
        SetClassLong m_TrackingData(PDIT_DEFAULT, currentIndex), GCL_HCURSOR, LoadCursor(0, IDC_ARROW) 'For now, always reset the cursor to an arrow.  This is necessary until we can fix some bad SetClassLong handling elsewhere in PD (via setHandCursor, when applied to picture boxes) :: m_PrevClassCursorHandle
        SetCursor LoadCursor(0, IDC_ARROW)
    End If
    
End Sub

'Given an hWnd, populate a generic TrackMouseEvent struct and pass it to the TrackMouseEvent function
Private Function requestMouseTrackingForHwnd(ByVal targetHwnd As Long)

    Dim eventTracker As TRACKMOUSEEVENT_STRUCT
    
    With eventTracker
        .cbSize = Len(eventTracker)
        .dwFlags = TME_LEAVE Or TME_HOVER
        .hWndTrack = targetHwnd
        .dwHoverTime = &HEFFFFFFF
    End With
    
    'TrackMouseEvent returns non-zero values for success
    If TrackMouseEvent(eventTracker) <> 0 Then
        requestMouseTrackingForHwnd = True
    Else
        requestMouseTrackingForHwnd = False
    End If
    
End Function

Private Sub Class_Terminate()

    'Release all subclassing and hooking
    cSubclass.shk_TerminateHooks
    cSubclass.ssc_Terminate
    
End Sub

'Note that the vKey constant below is a virtual key mapping, not (necessarily) a standard VB key constant - plan accordingly!
Private Function IsVirtualKeyDown(ByVal vKey As Long) As Boolean
    IsVirtualKeyDown = GetAsyncKeyState(vKey) And &H8000
End Function

'Note that the vKey constant below is a virtual key mapping, not necessarily a standard VB key constant
Private Function isMouseButtonDown(ByVal vKey As Long) As Boolean
    
    'Check for left/right mouse button switching, which is a system-wide user setting
    If GetSystemMetrics(SM_SWAPBUTTON) <> 0 Then
    
        'Reassign left/right mouse buttons as necessary
        If vKey = VK_LBUTTON Then
            vKey = VK_RBUTTON
        Else
            If vKey = VK_RBUTTON Then vKey = VK_LBUTTON
        End If
    
    End If
    
    isMouseButtonDown = IsVirtualKeyDown(vKey)
    
End Function

'Parse out the actual app command from the lParam of a WM_APPCOMMAND message.
' A big thank you to Steve McMahon for translating the GET_APPCOMMAND_LPARAM, GET_DEVICE_LPARAM, and GET_KEYSTATE_LPARAM macros
' into VB; you can see his original work here: http://www.vbaccelerator.com/home/VB/Tips/Responding_to_AppCommands/article.asp
Private Function parseAppCommand(ByVal lParam As Long, ByRef GET_APPCOMMAND_LPARAM As Long, Optional ByRef GET_DEVICE_LPARAM As Long, Optional ByRef GET_KEYSTATE_LPARAM As Long) As Long

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

'Given an hWnd, retrieve its index in our tracking array
Private Function getIndexFromHwnd(ByVal srcHwnd As Long) As Long

    Dim i As Long
    For i = 0 To m_hWndCount - 1
        If srcHwnd = m_TrackingData(PDIT_DEFAULT, i) Then
            getIndexFromHwnd = i
            Exit Function
        End If
    Next i
    
    'If we made it all the way here, the requested hWnd was not found.
    getIndexFromHwnd = -1

End Function

'Use GetCursorPos to retrieve the current mouse pointer coordinates.  Note that getCursorPos always uses screen coordinates,
' so we will manually translate the coords into the space of the passed hWnd.
Private Function getCurrentCursorPosition(ByVal srcHwnd As Long, ByRef controlX As Long, ByRef controlY As Long) As Boolean

    Dim tmpPoint As POINTAPI
    If GetCursorPos(tmpPoint) <> 0 Then
        
        'Convert the screen coordinates into the coordinate space of the supplied hWnd
        If ScreenToClient(srcHwnd, tmpPoint) <> 0 Then
            
            controlX = tmpPoint.x
            controlY = tmpPoint.y
            
            getCurrentCursorPosition = True
            
        Else
            getCurrentCursorPosition = False
        End If
    
    Else
        getCurrentCursorPosition = False
    End If

End Function

'The primary hWnd can use this function to move the mouse to some new position.
Public Sub moveCursorToNewPosition(ByVal xRelativeToHwnd As Double, ByVal yRelativeToHwnd As Double)

    'Retrieve the requested position, in screen coordinates
    Dim tmpPoint As POINTAPI
    tmpPoint.x = xRelativeToHwnd
    tmpPoint.y = yRelativeToHwnd
    
    If ClientToScreen(m_TrackingData(PDIT_DEFAULT, 0), tmpPoint) <> 0 Then
        
        'Cursor trails may cause ghosting, where we move the cursor, but a cursor is still shown at the old mouse position.
        ' To account for that, we must hide the cursor, move it, then re-display it.
        ShowCursor 0
        SetCursorPos tmpPoint.x, tmpPoint.y
        ShowCursor 1
        
    End If
    
End Sub

'Use SystemParametersInfo to retrieve the current user setting for mouse wheel sensitivity
Private Function getUserScrollSetting(ByVal directionIsVertical As Long) As Double

    Dim tmpUserScrollReturn As Long
    
    'Retrieve the system scroll setting
    If directionIsVertical Then
        SystemParametersInfo SPI_GETWHEELSCROLLLINES, 0, tmpUserScrollReturn, 0
    Else
        SystemParametersInfo SPI_GETWHEELSCROLLCHARS, 0, tmpUserScrollReturn, 0
    End If
    
    'Because we're going to use this function to calculate scroll amounts, don't allow it to be zero.
    ' (I don't know if Windows itself allows a 0 value, but it doesn't hurt to check, especially because
    ' the SystemParametersInfo call above could technically fail to return a value.)
    If tmpUserScrollReturn = 0 Then tmpUserScrollReturn = 3
    
    getUserScrollSetting = CDbl(tmpUserScrollReturn) / 120

End Function

'Hi and Lo word values can be retrieved using these helper functions.  They are required because VB doesn't have an unsigned int type,
' so we have to handle the signed bit specially (argh).
Private Function getHiWord(ByVal lParam As Long) As Integer
    If lParam And &H80000000 Then
        getHiWord = (lParam \ 65535) - 1
    Else
        getHiWord = lParam \ 65535
    End If
End Function

Private Function getLoWord(ByVal lParam As Long) As Integer
    If lParam And &H8000& Then
        getLoWord = &H8000 Or (lParam And &H7FFF&)
    Else
        getLoWord = lParam And &HFFFF&
    End If
End Function

'If high-resolution tracking is active, the subclasser will call this function when it receives a WM_MOUSEMOVE message.  This function will
' then use GetMouseMovePointsEx to reconstruct any missing mouse movement history.
Private Sub reconstructMouseMoveHistory(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal msgTimestamp As Long, ByVal srcHwnd As Long)

    'First, note that the passed points are in *window* coordinates.  We need to convert them to *screen* coordinates.
    Dim clientX As Long, clientY As Long
    clientX = x
    clientY = y
    
    Dim tmpPoint As POINTAPI
    tmpPoint.x = clientX
    tmpPoint.y = clientY
    
    ClientToScreen srcHwnd, tmpPoint
    
    Dim screenX As Long, screenY As Long
    screenX = tmpPoint.x
    screenY = tmpPoint.y

    'If we have never tracked a previous point, we don't want to reconstruct a mouse movement history yet (because this is the user's
    ' first click!)  Store the passed point for future reference, then exit immediately.
    If Not m_PrevPointStored Then
    
        m_PrevPointStored = True
        
        With m_LastPointTracked
            .x = screenX And &HFFFF&
            .y = screenY And &HFFFF&
            .ptTime = msgTimestamp
        End With
        
        RaiseEvent MouseMoveCustom(Button, Shift, clientX, clientY)
        Exit Sub
        
    End If
    
    'If we made it all the way here, a previous point has been stored.  We will now reconstruct a full mouse history between that point
    ' and the current one.

    'Current tracking mode.  We can use either high-resolution pen input coordinates, or regular mouse coordinates.
    ' For now, PD only uses regular mouse coordinates, but perhaps we could make use of the pen code in the future,
    ' so I've included it here.
    Dim curTrackingResolution As mouseResolution
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
    Dim mouseHistory(0 To 63) As MOUSEMOVEPOINT
    
    'Retrieve the mouse movement history.  The return value of the function is the number of points filled.
    Dim numPoints As Long
    numPoints = GetMouseMovePointsEx(Len(curMouseMovePoint), curMouseMovePoint, VarPtr(mouseHistory(0)), 64, GMMP_USE_DISPLAY_POINTS)
        
    'If one or more valid points were returned, carry on
    If numPoints > 0 Then
    
        'We now want to search the mouse movement history for the previous point we returned.  All points after that will need to be manually raised.
        Dim i As Long
        For i = 0 To numPoints - 1
        
            'If we reach a point that occurred prior to our last tracked time, exit
            If mouseHistory(i).ptTime < m_LastPointTracked.ptTime Then Exit For
            
            'If we reach our exact last point time, also exit
            If (mouseHistory(i).ptTime = m_LastPointTracked.ptTime) And (mouseHistory(i).x = m_LastPointTracked.x) And (mouseHistory(i).y = m_LastPointTracked.y) Then Exit For
        
        Next i
        
        'The variable "i" now contains the index of the previous point (or one that occurred just prior to it, if the previous point wasn't found).
        ' Iterate through all points following that one, and convert them to valid screen coordinates.
        numPoints = i - 1
        If numPoints < 0 Then numPoints = 0
        
        For i = numPoints To 0 Step -1
        
            'Convert the unsigned 16-bit return values to signed 32-bit values
            Select Case curTrackingResolution
            
                Case GMMP_USE_DISPLAY_POINTS
                
                    If (mouseHistory(i).x > 32767) Then mouseHistory(i).x = mouseHistory(i).x - 65536
                    If (mouseHistory(i).y > 32767) Then mouseHistory(i).y = mouseHistory(i).y - 65536
                
                Case GMMP_USE_HIGH_RESOLUTION_POINTS
                
                    mouseHistory(i).x = ((mouseHistory(i).x * (nVirtualWidth - 1)) - (nVirtualLeft * 65536)) / nVirtualWidth
                    mouseHistory(i).y = ((mouseHistory(i).y * (nVirtualHeight - 1)) - (nVirtualTop * 65536)) / nVirtualHeight
                
            End Select
            
            'With this point successfully transformed, the last thing we need to do is convert it from screen to window coordinates.
            tmpPoint.x = mouseHistory(i).x
            tmpPoint.y = mouseHistory(i).y
            ScreenToClient m_TrackingData(0, 0), tmpPoint
            
            'If automatic message dropping is active, make sure this message falls within the acceptable threshold
            ' before firing it.
            If m_DelayTrackingActive Then
            
                If (getMessageDelay(mouseHistory(i).ptTime) <= m_DelayThreshold) Then
                    RaiseEvent MouseMoveCustom(Button, Shift, tmpPoint.x, tmpPoint.y)
                End If
            
            Else
                RaiseEvent MouseMoveCustom(Button, Shift, tmpPoint.x, tmpPoint.y)
            End If
        
        Next i
        
    'If the current point is also the previous point, GetMouseMovePointsEx may not have any prior points to return.  This is fine -
    ' simply raise the point we received, then exit.
    Else
        
        'If automatic message dropping is active, make sure this message falls within the acceptable threshold
        ' before firing it.
        If m_DelayTrackingActive Then
        
            If (getMessageDelay(msgTimestamp) <= m_DelayThreshold) Then
                RaiseEvent MouseMoveCustom(Button, Shift, clientX, clientY)
            End If
        
        Else
            RaiseEvent MouseMoveCustom(Button, Shift, clientX, clientY)
        End If
        
    End If
    
    'Update the "previous point" tracker with the current point, then exit
    With m_LastPointTracked
        .x = screenX And &HFFFF&
        .y = screenY And &HFFFF&
        .ptTime = msgTimestamp
    End With
    
End Sub

'Given a message time (as retrieved by GetMessageTime(), typically), report the delay in milliseconds between the posted time and the current time.
Private Function getMessageDelay(ByVal srcMessageTime As Long) As Long
    getMessageDelay = GetTickCount() - srcMessageTime
End Function


'All events subclassed by this window are processed here.
Private Sub myWndProc(ByVal bBefore As Boolean, _
                      ByRef bHandled As Boolean, _
                      ByRef lReturn As Long, _
                      ByVal lng_hWnd As Long, _
                      ByVal uMsg As Long, _
                      ByVal wParam As Long, _
                      ByVal lParam As Long, _
                      ByRef lParamUser As Long)
'*************************************************************************************************
'* bBefore    - Indicates whether the callback is before or after the original WndProc. Usually
'*              you will know unless the callback for the uMsg value is specified as
'*              MSG_BEFORE_AFTER (both before and after the original WndProc).
'* bHandled   - In a before original WndProc callback, setting bHandled to True will prevent the
'*              message being passed to the original WndProc and (if set to do so) the after
'*              original WndProc callback.
'* lReturn    - WndProc return value. Set as per the MSDN documentation for the message value,
'*              and/or, in an after the original WndProc callback, act on the return value as set
'*              by the original WndProc.
'* lng_hWnd   - Window handle.
'* uMsg       - Message value.
'* wParam     - Message related data.
'* lParam     - Message related data.
'* lParamUser - User-defined callback parameter. Change vartype as needed (i.e., Object, UDT, etc)
'*************************************************************************************************

    Dim i As Long
    Dim overrideRequired As Boolean
    
    'Before processing anything, retrieve this message's timestamp
    Dim msgTimestamp As Long
    msgTimestamp = GetMessageTime()
    
    'Next, do a failsafe check to make sure the hWnd we've been passed is one we're actually responsible for.
    ' (This should always be TRUE, but better safe than sorry.)
    Dim indexOfHwnd As Long
    indexOfHwnd = getIndexFromHwnd(lng_hWnd)
    
    'If the hWnd was found in our tracking array, start processing the window message.
    If (indexOfHwnd >= 0) Then
        
        'Note this hWnd, in case the user needs to retrieve it mid-event
        m_LastMsgHWnd = lng_hWnd
        
        'As VB already does with its own mouse events, it is helpful to supply key modifiers directly in the event
        ' params.  Unfortunately, the window messages that report key states will only report SHIFT and CTRL masks,
        ' not ALT.  (This is by design, as most windows use the Alt key to forcibly switch focus to the menu, which
        ' in turn makes it irrelevant for mouse modifications.  However, this behavior can be overridden if the
        ' mouse has been captured by a window, which we may end up doing for PD's canvas, and which is why I'm
        ' adding ALT handling here.)
        
        'Anyway, because the Alt key requires special handling, and some messages don't report button state, I just
        ' ignore window message reports entirely and pull key states manually using GetAsyncKeyState.
        Dim retShiftConstants As ShiftConstants
        
        If IsVirtualKeyDown(VK_SHIFT) Then retShiftConstants = retShiftConstants Or vbShiftMask
        If IsVirtualKeyDown(VK_CONTROL) Then retShiftConstants = retShiftConstants Or vbCtrlMask
        
        'NOTE!  MSDN provides some odd instructions for tracking the ALT key (see Remarks here:
        ' http://msdn.microsoft.com/en-us/library/ms646242%28v=vs.85%29.aspx).  They state explicitly to use
        ' use GetKeyState and not GetAsyncKeyState, but they don't provide any rationale for this.  I have no
        ' trouble with GetAsyncKeyState properly reporting Alt status on Win 7, so I'm leaving this strategy
        ' for now, but I remain intrigued by MSDN's oddly specific instructions on the point...
        If IsVirtualKeyDown(VK_ALT) Then retShiftConstants = retShiftConstants Or vbAltMask
        
        'Similarly, always grab mouse buttons independent of the window message
        Dim retMouseButton As PDMouseButtonConstants
        If isMouseButtonDown(VK_LBUTTON) Then retMouseButton = retMouseButton Or pdLeftButton
        If isMouseButtonDown(VK_MBUTTON) Then retMouseButton = retMouseButton Or pdMiddleButton
        If isMouseButtonDown(VK_RBUTTON) Then retMouseButton = retMouseButton Or pdRightButton
        If isMouseButtonDown(VK_XBUTTON1) Then retMouseButton = retMouseButton Or pdXButtonOne
        If isMouseButtonDown(VK_XBUTTON2) Then retMouseButton = retMouseButton Or pdXButtonTwo
        
        'Some functions also supply mouse button and coordinate values; we'll be retrieving these as well.
        ' To improve accuracy of the retrieved coordinates (and to make future support for GetMouseMovePointsEx
        ' easier to implement), we're going to manually retrieve coordinates using GetCursorPos, then translate
        ' these into the coordinate space of the master hWnd given to this class.  This allows for some neat
        ' extendability, like a user control with multiple child controls, but who receives only a singular
        ' MouseMove event for the entire control, with all coordinates automatically transformed.
        Dim mouseX As Long, mouseY As Long
        If Not getCurrentCursorPosition(m_TrackingData(PDIT_DEFAULT, 0), mouseX, mouseY) Then
            Debug.Print "WARNING!  pdInput could not retrieve mouse coordinates relative to hWnd #" & lng_hWnd & ". Fix it!"
        End If
        
        'As a failsafe, crucial mouse events (button down and up, particularly) can never afford to be delayed.  We must retrieve
        ' their values at *exactly* the moment the message arrives.  To that end, we always store copies of these integer-only values,
        ' using the values passed in the message parameters.
        Dim mouseXInt As Integer, mouseYInt As Integer
        mouseXInt = getLoWord(lParam)
        mouseYInt = getHiWord(lParam)
        
        'Regardless of which event we are handling, we now have Shift modifier, mouse button, and mouse coordinate
        ' data (relative to the master hWnd for this class).  This makes individual message parsing much easier, as
        ' some messages don't report this info automatically.
        
        'Parse window messages and raise any standard events; at present, this includes: mouse enter, mouse leave,
        ' and mousewheel.  Other events handled by this class are optional, and will be handled separately.
        Select Case uMsg
        
            'First things first: handle mouse enter and leave events.
            Case WM_MOUSEMOVE
                
                'Regardless of button or hWnd, increment the MouseMove counter
                m_MouseMoveCount = m_MouseMoveCount + 1
                
                'Make sure we are actually tracking mouse enter/leave events for this hWnd
                If m_TrackingData(PDIT_MOUSE_ENTERLEAVE, indexOfHwnd) Then
                
                    'Reset hover mode tracking
                    m_HoverModeActive = False
                
                    'If the mouse is moving, that means it has entered the control.  If we aren't already tracking it,
                    ' do so now, and raise a MouseEnter event to match.
                    If Not m_MouseTrackingActive(indexOfHwnd) Then
                        
                        m_MouseTrackingActive(indexOfHwnd) = True
                        requestMouseTrackingForHwnd lng_hWnd
                        
                        'If cursor handling is enabled, reset the class cursor and re-apply any custom cursor now.
                        If m_TrackingData(PDIT_CURSOR, indexOfHwnd) Then applyCursorToHwnds True, indexOfHwnd
                        
                        'Only raise a MouseEnter event if this event is happening for the primary hWnd; otherwise, track
                        ' mouse events silently.
                        '
                        'NOTE! This will raise more MouseEnter events than we want in the case of a UC with multiple
                        '       tracked controls; I'll investigate a better solution if/when it proves necessary.
                        '
                        'NOTE 2: As of March 2015, fixes to the color selection dialog require me to rework this.  I now
                        ' raise MouseEnter for all tracked hWnds, so that the hWnd can request cursor changes as necessary.
                        ' Hopefully this doesn't break anything elsewhere in the program...
                        RaiseEvent MouseEnter(retMouseButton, retShiftConstants, mouseX, mouseY)
                        
                    End If
                                        
                    'If custom MouseMove tracking has been requested for this hWnd, raise a mouse move event now
                    If m_TrackingData(PDIT_STANDARD_MOUSE, indexOfHwnd) Then
                    
                        'There are two ways to raise MouseMove events for a window.  If high-resolution tracking is active,
                        ' we need to reconstruct any missing mouse events and raise them all in turn.  Otherwise, we can
                        ' simply return this event.
                        If m_HighResModeActive And g_HighResolutionInput Then
                            reconstructMouseMoveHistory retMouseButton, retShiftConstants, mouseX, mouseY, msgTimestamp, lng_hWnd
                        Else
                        
                            'If automatic message dropping is active, make sure this message falls within the acceptable threshold
                            ' before firing it.
                            If m_DelayTrackingActive Then
                            
                                If (getMessageDelay(msgTimestamp) <= m_DelayThreshold) Then
                                    RaiseEvent MouseMoveCustom(retMouseButton, retShiftConstants, mouseX, mouseY)
                                End If
                                
                            Else
                                RaiseEvent MouseMoveCustom(retMouseButton, retShiftConstants, mouseX, mouseY)
                            End If
                            
                            
                        End If
                        
                        'Random fact!  If bHandled is set to TRUE, tooltips will not appear for the window in question.
                        ' I haven't yet discovered a reason for this, but I can reliably reproduce the issue.  My assumption,
                        ' based on the MSDN isntructions of "if an application processes this message, it should return zero,"
                        ' is that DefWindowProc needs to pass the message onto the tooltip object so it can process its own
                        ' internal timer for showing the tooltip; by setting bHandled to True, we prevent that hand-off, so
                        ' the tooltip never gets notification.
                        'bHandled = True
                        
                        lReturn = 0
                        
                    End If
                    
                End If
            
            'Mouse down
            Case WM_LBUTTONDOWN, WM_MBUTTONDOWN, WM_RBUTTONDOWN, WM_XBUTTONDOWN
            
                lReturn = 0
                
                'Increment the capture counter
                m_CaptureCounter = m_CaptureCounter + 1
                
                'If we have not captured the mouse, capture it now
                If (m_CaptureCounter = 1) And (Not m_DoNotCaptureOnButtonPress) And (Not m_MouseHasBeenCaptured) Then
                    SetCapture lng_hWnd
                    m_MouseHasBeenCaptured = True
                End If
                
                'Reset the mouse move counter
                m_MouseMoveCount = 0
                
                'Because mouse up/down events tend to focus on the use of a single button, it makes more sense for us to
                ' manually return a single-button type here (rather than a composite flag).  The calling function can
                ' always query this class if it wants more detailed button up/down state data.
                retMouseButton = 0
                
                If uMsg = WM_LBUTTONDOWN Then
                    retMouseButton = vbLeftButton
                ElseIf uMsg = WM_RBUTTONDOWN Then
                    retMouseButton = vbRightButton
                ElseIf uMsg = WM_MBUTTONDOWN Then
                    retMouseButton = vbMiddleButton
                ElseIf uMsg = WM_XBUTTONDOWN Then
                
                    'In this rare instance, we have to pull data directly from the message, because which XButton triggered
                    ' the message can only be found by checking flags in the high-order word of the wParam.
                    If (wParam \ &H10000) = 1 Then
                        retMouseButton = pdXButtonOne
                    Else
                        retMouseButton = pdXButtonTwo
                    End If
                    
                    'Also strange, X-button events require us to return TRUE!  This only for matters for pre-W2K systems,
                    ' but sticking to the rule shows how carefully we read MSDN docs... ;)
                    lReturn = 1
                    
                End If
                
                RaiseEvent MouseDownCustom(retMouseButton, retShiftConstants, mouseXInt, mouseYInt)
                
                If m_IgnoreLastMessage Then
                    bHandled = False
                    lReturn = DefWindowProc(lng_hWnd, uMsg, wParam, lParam)
                    m_IgnoreLastMessage = False
                End If
            
            'Mouse up
            Case WM_LBUTTONUP, WM_MBUTTONUP, WM_RBUTTONUP, WM_XBUTTONUP
            
                lReturn = 0
                
                'Decrement the capture counter; if all buttons have been released, release our capture entirely
                m_CaptureCounter = m_CaptureCounter - 1
                If (m_CaptureCounter <= 0) And (Not m_DoNotCaptureOnButtonPress) And m_MouseHasBeenCaptured Then
                    m_CaptureCounter = 0
                    ReleaseCapture
                    m_MouseHasBeenCaptured = False
                End If
                
                'Because mouse up/down events tend to focus on the use of a single button, it makes more sense for us to
                ' manually return a single-button type here (rather than a composite flag).  The calling function can
                ' always query this class if it wants more detailed button up/down state data.
                retMouseButton = 0
                
                If uMsg = WM_LBUTTONUP Then
                    retMouseButton = vbLeftButton
                ElseIf uMsg = WM_RBUTTONUP Then
                    retMouseButton = vbRightButton
                ElseIf uMsg = WM_MBUTTONUP Then
                    retMouseButton = vbMiddleButton
                ElseIf uMsg = WM_XBUTTONUP Then
                
                    'In this rare instance, we have to pull data directly from the message, because which XButton triggered
                    ' the message can only be found by checking flags in the high-order word of the wParam.
                    If (wParam \ &H10000) = 1 Then
                        retMouseButton = pdXButtonOne
                    Else
                        retMouseButton = pdXButtonTwo
                    End If
                    
                    'Also strange, X-button events require us to return TRUE!  This only for matters for pre-W2K systems,
                    ' but sticking to the rule shows how carefully we read MSDN docs... ;)
                    lReturn = 1
                    
                End If
                
                'Fun fact!  If PD is busy doing a bunch of processing during MouseMove events (such as drag-resizing a
                ' large layer), the MouseUp event may become artificially delayed.  When it finally triggers, the mouse
                ' may have moved beyond its original point prior to release, causing the MouseUp event to report
                ' inappropriate mouse values.  The best solution to this would be to grab the X/Y coordinates from the
                ' window message, then use GetMouseMovePointsEx to retrieve a high-resolution copy of the coordinates.
                ' Until we implement GetMouseMovePointsEx as a pdInput-wide solution, however, I'm going to simply
                ' take the x/y values from the message and report them for now.
                '
                '(The point of that very long message is to explain why we use mouseXInt and mouseYInt, below.)
                
                'If the user has not moved the mouse an appreciable amount since the last MouseDown event,
                ' call this a click and raise the corresponding Click() event.
                '
                'Note that a MouseUp event is raised either way, and if a Click event is also being raised,
                ' we warn the user in the MouseUp event.  This way, they can choose to abandon MouseUp
                ' processing if the _Click event is implemented.
                If m_MouseMoveCount < 3 Then
                    
                    RaiseEvent MouseUpCustom(retMouseButton, retShiftConstants, mouseXInt, mouseYInt, True)
                    RaiseEvent ClickCustom(retMouseButton, retShiftConstants, mouseXInt, mouseYInt)
                
                'The mouse has moved too much, so a Click event will not be raised.  Only raise a MouseUp event.
                Else
                    RaiseEvent MouseUpCustom(retMouseButton, retShiftConstants, mouseXInt, mouseYInt, False)
                End If
                
                If m_IgnoreLastMessage Then
                    bHandled = False
                    lReturn = DefWindowProc(lng_hWnd, uMsg, wParam, ByVal lParam)
                    m_IgnoreLastMessage = False
                End If
                
            'Double-clicks
            Case WM_LBUTTONDBLCLK, WM_MBUTTONDBLCLK, WM_RBUTTONDBLCLK, WM_XBUTTONDBLCLK
            
                lReturn = 0
                
                'Because double-click events tend to focus on the use of a single button, it makes more sense for us to
                ' manually return a single-button type here (rather than a composite flag).
                retMouseButton = 0
                
                If uMsg = WM_LBUTTONDBLCLK Then
                    retMouseButton = vbLeftButton
                ElseIf uMsg = WM_RBUTTONDBLCLK Then
                    retMouseButton = vbRightButton
                ElseIf uMsg = WM_MBUTTONDBLCLK Then
                    retMouseButton = vbMiddleButton
                ElseIf uMsg = WM_XBUTTONDBLCLK Then
                
                    'In this rare instance, we have to pull data directly from the message, because which XButton triggered
                    ' the message can only be found by checking flags in the high-order word of the wParam.
                    If (wParam \ &H10000) = 1 Then
                        retMouseButton = pdXButtonOne
                    Else
                        retMouseButton = pdXButtonTwo
                    End If
                    
                    'Also strange, X-button events require us to return TRUE!  This only for matters for pre-W2K systems,
                    ' but sticking to the rule shows how carefully we read MSDN docs... ;)
                    lReturn = 1
                    
                End If
                
                RaiseEvent DoubleClickCustom(retMouseButton, retShiftConstants, mouseX, mouseY)
                
                If m_IgnoreLastMessage Then
                    bHandled = False
                    lReturn = DefWindowProc(lng_hWnd, uMsg, wParam, lParam)
                    m_IgnoreLastMessage = False
                End If
            
            'Mouse leaving the client area
            Case WM_MOUSELEAVE
            
                'Make sure we are actually tracking mouse enter/leave events for this hWnd
                If m_TrackingData(PDIT_MOUSE_ENTERLEAVE, indexOfHwnd) Then
            
                    'If we've been tracking mouse movements for this hWnd, raise a MouseLeave event now
                    If m_MouseTrackingActive(indexOfHwnd) Then
                        
                        m_MouseTrackingActive(indexOfHwnd) = False
                        
                        'Only raise a MouseLeave event if this event is happening for the primary hWnd; otherwise, track
                        ' mouse events silently.
                        '
                        'NOTE! This will raise more MouseLeave events than we want in the case of a UC with multiple
                        '       tracked controls; I'll investigate a better solution if/when it proves necessary.
                        If indexOfHwnd = 0 Then RaiseEvent MouseLeave(retMouseButton, retShiftConstants, mouseX, mouseY)
                        
                        'If cursor handling is enabled for this hWnd, reset the cursor to its original value whenever
                        ' the mouse has left the control.  This will prevent us from mistakenly assigning our custom
                        ' cursor to other objects that share this class.
                        If m_TrackingData(PDIT_CURSOR, indexOfHwnd) Then resetClassCursor
                        
                    End If
                    
                    bHandled = True
                    
                End If
                
            'Next, process MouseHover events.  We don't use this at present, but as it's basically the same code as
            ' MouseLeave, I've included it "just in case".
            Case WM_MOUSEHOVER
            
                'Only process hover events if we care about tracking mouse movements for this hWnd
                If m_MouseTrackingActive(indexOfHwnd) And (Not m_HoverModeActive) Then
                    
                    'Note that we don't change the tracking state, because we don't want to generate subsequent
                    ' MouseEnter events after this one!  Instead, silently re-request mouse tracking for this hWnd.
                    requestMouseTrackingForHwnd lng_hWnd
                    
                    'Note that we are currently in hover mode; this won't be reset until some kind of MouseMove occurs
                    m_HoverModeActive = True
                    
                    'Unlike Enter/Leave events, we want to raise hover events regardless of the hWnd; the assumption
                    ' here is that the target window will always want these events, and if it doesn't, it can do its
                    ' own bounds-checking to determine if a hover event is relevant or not.
                    RaiseEvent MouseHover(retMouseButton, retShiftConstants, mouseX, mouseY)
                    
                End If
                
                bHandled = True
                
            'Next, process mousewheel events.  Note that horizontal and vertical scrolls are handled together (for the
            ' most part).
            Case WM_MOUSEWHEEL, WM_MOUSEHWHEEL
            
                'First things first: retrieve the high-word, which contains the change (delta) in mousewheel position
                Dim scrollDelta As Long
                scrollDelta = wParam \ &H10000
                
                'Next, use the user's scroll wheel setting (set via the Control Panel) to calculate a final scroll amount,
                ' in lines or chars depending on whether vertical or horizontal scrolling is active.
                Dim finalScrollAmount As Double
                
                'Vertical scroll only, meaning Vertical Wheel + !Shift and !Ctrl
                If (uMsg = WM_MOUSEWHEEL) And ((retShiftConstants And vbShiftMask) = 0) And ((retShiftConstants And vbCtrlMask) = 0) Then
                    finalScrollAmount = CDbl(scrollDelta) * getUserScrollSetting(True)
                    RaiseEvent MouseWheelVertical(retMouseButton, retShiftConstants, mouseX, mouseY, finalScrollAmount)
                
                'Horizontal scroll, but generated via Vertical Wheel + Shift and !Ctrl
                ElseIf (uMsg = WM_MOUSEWHEEL) And ((retShiftConstants And vbShiftMask) <> 0) And ((retShiftConstants And vbCtrlMask) = 0) Then
                
                    finalScrollAmount = CDbl(scrollDelta) * getUserScrollSetting(True)
                    
                    'Because the user is using the Shift+VerticalWheel combination, reverse the delta; this makes it so
                    ' that shift+up_wheel = left, and shift+down_wheel = right - the idea is that up_wheel and shift+up_wheel
                    ' both target the top-left corner of the image.
                    finalScrollAmount = -1 * finalScrollAmount
                    
                    RaiseEvent MouseWheelHorizontal(retMouseButton, retShiftConstants, mouseX, mouseY, finalScrollAmount)
                    
                'Zoom scroll, generated via Vertical Wheel + Ctrl and !Shift
                ElseIf (uMsg = WM_MOUSEWHEEL) And ((retShiftConstants And vbCtrlMask) <> 0) And ((retShiftConstants And vbShiftMask) = 0) Then
                
                    finalScrollAmount = CDbl(scrollDelta) * getUserScrollSetting(True)
                    RaiseEvent MouseWheelZoom(retMouseButton, retShiftConstants, mouseX, mouseY, finalScrollAmount)
                
                'Horizontal scroll, generated via actual Horizontal Scroll/Tilt message
                ElseIf (uMsg = WM_MOUSEHWHEEL) Then
                
                    finalScrollAmount = CDbl(scrollDelta) * getUserScrollSetting(False)
                    RaiseEvent MouseWheelHorizontal(retMouseButton, retShiftConstants, mouseX, mouseY, finalScrollAmount)
                    
                'The only combination left is Vertical scroll, but with some weird combination of Shift modifiers; return
                ' a generic vertical wheel event as-is.
                Else
                    finalScrollAmount = CDbl(scrollDelta) * getUserScrollSetting(True)
                    RaiseEvent MouseWheelVertical(retMouseButton, retShiftConstants, mouseX, mouseY, finalScrollAmount)
                
                End If
                
                If m_IgnoreLastMessage Then
                    bHandled = False
                    m_IgnoreLastMessage = False
                Else
                    lReturn = 1
                    bHandled = True
                End If
                
            'App commands are meta functions like "back" or "increase volume" or "media rewind".  They can be triggered
            ' by a variety of actions, depending on system configuration: mouse gestures, dedicated keyboard keys,
            ' gestures on a touch screen.  When one occurs, raise an event that the owner can deal with as desired;
            ' we will pass them the app command received, so they can deal with as many (or as few) as they desire.
            Case WM_APPCOMMAND
            
                'At present, we don't process the additional device or key parameters - just the action ID
                If m_TrackingData(PDIT_APPCOMMAND, indexOfHwnd) Then
                    Dim commandID As Long
                    parseAppCommand lParam, commandID
                    RaiseEvent AppCommand(commandID, retShiftConstants, mouseX, mouseY)
                End If
            
            'While handling cursor duties, we must respond to WM_SETCURSOR events.
            Case WM_SETCURSOR
            
                If m_TrackingData(PDIT_CURSOR, indexOfHwnd) And (Not m_OverrideSetCursorMessages) Then
                    applyCursorToHwnds False, indexOfHwnd
                    bHandled = True
                    lReturn = 1
                End If
                
        End Select

    
    'Somehow, the hWnd we've captured is not in our tracking list!  Abandon ship.
    Else
    
        Debug.Print "WARNING!  Irrelevant hWnd intercepted by pdInput: " & lng_hWnd
        bHandled = False
    
    End If



' *************************************************************
' C A U T I O N   C A U T I O N   C A U T I O N   C A U T I O N
' -------------------------------------------------------------
' DO NOT ADD ANY OTHER CODE BELOW THE "END SUB" STATEMENT BELOW
'   add this warning banner to the last routine in your class
' *************************************************************
End Sub


