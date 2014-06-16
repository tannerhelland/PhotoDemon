VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Input Handler (mouse, pen, etc) class
'Copyright ©2013-2014 by Tanner Helland
'Created: 27/May/14 (though many individual parts existed earlier than this!)
'Last updated: 16/June/14
'Last update: start extending key tracking functions so we can add more key bundles to the tracker
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
'   or if you want to handle all mouse events for a collection of objects with a single MouseMove function.  Similarly, if
'   you want to subclass multiple controls within a single parent, that should work just fine, as long as you don't assign
'   the parent hWnd to multiple instances of this class (e.g. create just ONE pdInput copy for the entire collection, rather
'   than different ones for each child).
'
'- As implied above, mouse coordinate, button, and shift key modifiers are processed independent of window messages.  This
'   allows us to supply such data even if a given window message doesn't automatically include them.  It also allows for
'   higher accuracy when tracking mouse move data, as we use 32-bit values instead of 16-bit ones (which seems ridiculous
'   now, but may not be forever).
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
Event MouseLeave(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
Event MouseEnter(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
Event MouseHover(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
Event MouseWheelVertical(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal scrollAmount As Double)
Event MouseWheelHorizontal(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal scrollAmount As Double)
Event MouseWheelZoom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal zoomAmount As Double)
Event ClickCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
Event DoubleClickCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
Event MouseDownCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
Event MouseUpCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal ClickEventAlsoFiring As Boolean)

'The following events are raised only if specifically requested in the addInputTracker function
Event MouseMoveCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
Event AppCommand(ByVal cmdID As AppCommandConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)

'The following events are raised only if specifically requested via the dedicated requestKeyTracking and setKeyTrackers functions
Event KeyDownArrows(ByVal Shift As ShiftConstants, ByVal upArrow As Boolean, ByVal rightArrow As Boolean, ByVal downArrow As Boolean, ByVal leftArrow As Boolean)
Event KeyDownEdits(ByVal Shift As ShiftConstants, ByVal kReturn As Boolean, ByVal kEnter As Boolean, ByVal kSpaceBar As Boolean, ByVal kBackspace As Boolean, ByVal kInsert As Boolean, ByVal kDelete As Boolean, ByVal kTab As Boolean)

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

Private Const WM_KEYDOWN As Long = &H100
Private Const WM_GETDLGCODE As Long = &H87

Private Const WM_ACTIVATE As Long = &H6
Private Const WM_SETFOCUS As Long = &H7
Private Const WM_KILLFOCUS As Long = &H8

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

Private Const VK_LEFT As Long = &H25
Private Const VK_UP As Long = &H26
Private Const VK_RIGHT As Long = &H27
Private Const VK_DOWN As Long = &H28

Private Const VK_NUMLOCK As Long = &H90
Private Const VK_NUMPAD0 As Long = &H60
Private Const VK_NUMPAD1 As Long = &H61
Private Const VK_NUMPAD2 As Long = &H62
Private Const VK_NUMPAD3 As Long = &H63
Private Const VK_NUMPAD4 As Long = &H64
Private Const VK_NUMPAD5 As Long = &H65
Private Const VK_NUMPAD6 As Long = &H66
Private Const VK_NUMPAD7 As Long = &H67
Private Const VK_NUMPAD8 As Long = &H68
Private Const VK_NUMPAD9 As Long = &H69

Private Const VK_BACK As Long = &H8
Private Const VK_TAB As Long = &H9
Private Const VK_RETURN As Long = &HD
Private Const VK_SPACE As Long = &H20
Private Const VK_INSERT As Long = &H2D
Private Const VK_DELETE As Long = &H2E

'In the future, other virtual key codes can be retrieved here:
' http://msdn.microsoft.com/en-us/library/windows/desktop/dd375731%28v=vs.85%29.aspx

'To see if a given key is marked as an "extended key" by the system (e.g. Return vs NumPad Enter key, which have the same
' virtual key code), perform a bitwise AND (&&) of this flag against the lParam of the WM_KEYDOWN event
Private Const KF_EXTENDED As Long = &H1000000

'Mouse buttons can be retrieved from various mouse messages, but for consistency's sake, we pull them straight
' from GetAsyncKeyState.  One thing to note about GetAsyncKeyState is that it returns the physical mouse button
' pressed, *without button swapping* for left-handed mouse users.  We need to check this state and manually
' handle it, using GetSystemMetrics (which is publicly declared in PD).
Private Const SM_SWAPBUTTON As Long = 23

'Cursors can be requested by external functions, so we declare them publicly, in the Icon and Cursor module.

'Retrieve the current cursor position, in screen coordinates
Private Declare Function GetCursorPos Lib "user32" (ByRef lpPoint As POINTAPI) As Long

'Set a new cursor for a given class
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function GetClassLong Lib "user32" Alias "GetClassLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetClassLong Lib "user32" Alias "SetClassLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Const GCL_HCURSOR = (-12)

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
    PDIT_KEYTRACKING = 5
    PDIT_HASFOCUS = 6
    PDIT_ARROWKEYS = 7
    PDIT_EDITKEYS = 8
End Enum

'When tracking arrow keys, we use a few different track conditions (compared to just a binary TRUE/FALSE tracking code)
Private Const ARROW_TRACKING_OBEY_NUMLOCK = 1
Private Const ARROW_TRACKING_IGNORE_NUMLOCK = 2

#If False Then
    Private Const PDIT_DEFAULT = 0, PDIT_MOUSE_ENTERLEAVE = 1, PDIT_STANDARD_MOUSE = 2, PDIT_APPCOMMAND = 3, PDIT_CURSOR = 4, _
    PDIT_KEYTRACKING = 5, PDIT_HASFOCUS = 6, PDIT_ARROWKEYS = 7, PDIT_EDITKEYS = 8
#End If

Private Const NUM_OF_CUSTOM_INPUT_TRACKERS = 9

'This class will generate Click() events using its own algorithm for determining when an action constitutes a "click".
' Basically, a MouseDown/Up combination with two or less MouseMove messages between them constitutes a Click.  This variable
' is used to count MouseMove occurrences.
Private m_MouseMoveCount As Long

'User controls don't handle focus correctly, so manual override control hWnds will be stored here.  (See comments for
' the addOverrideHwnds sub for details)
Private m_numOfOverrideControls As Long
Private m_OverrideHwnds() As Long

'VB is a source of endless focus when it comes to handling focus on UCs.  In VB's mind, user controls should fight
' tooth-and-nail to make sure the UC itself never receives focus, even if we explicitly state that the UC can and should be
' focusable.  Because of this, when the user interacts with a UC, VB will always attempt to pass focus to a child object.
' Unfortunately, which child object gets focus is not predictable, though it seems to be tied to the Tab chain, e.g. the
' next object in the tab line past whichever child control previously had focus will be the next one to get focus this
' time.  This causes all sorts of havoc when attempting to handle unified system key messages (e.g. arrows) for the UC
' as a single unit.
'
'To outsmart VB, UCs can now make use of this addOverrideHwnds function.  Upon initializing key hooking for a UC, call this
' function and supply the hWnd of every focus-irrelevant control that VB might send focus to - picture boxes, etc.  This
' tracker will then track their hWnds in the main hook handler, and treat any key events from those children as belonging to
' the parent UC.  (Note that this still allows for proper arrow key usage within relevant controls, like a combo box.)
'
'To make sure that focus is handled properly for your UC when the user interacts with a non-control region, add something like
' "<PictureBoxName>.SetFocus" to the mouse down event, which will ensure that one of the overridden hWnds gets focus, rather
' than a random combo box or other control.
Public Sub addOverrideHwnds(ParamArray ExtraHWnds() As Variant)

    If Not IsMissing(ExtraHWnds) Then
        
        Dim i As Long
        For i = LBound(ExtraHWnds) To UBound(ExtraHWnds)
            If ExtraHWnds(i) <> 0 Then
                m_numOfOverrideControls = m_numOfOverrideControls + 1
                ReDim Preserve m_OverrideHwnds(0 To m_numOfOverrideControls - 1) As Long
                m_OverrideHwnds(m_numOfOverrideControls - 1) = CLng(ExtraHWnds(i))
                
                cSubclass.ssc_Subclass m_OverrideHwnds(m_numOfOverrideControls - 1), , 1, Me
                cSubclass.ssc_AddMsg m_OverrideHwnds(m_numOfOverrideControls - 1), MSG_BEFORE, WM_KILLFOCUS, WM_SETFOCUS
                
            End If
        Next i
        
    End If

End Sub

Private Sub Class_Initialize()
    
    'Reset all input tracking variables
    ReDim m_TrackingData(0 To NUM_OF_CUSTOM_INPUT_TRACKERS - 1, 0) As Long
    ReDim m_MouseTrackingActive(0) As Boolean
    
    m_numOfOverrideControls = 0
    ReDim m_OverrideHwnds(0) As Long
    
    'Note that we are not currently tracking the mouse for this object
    m_MouseTrackingActive(0) = False
    
    'Initialize the subclasser
    Set cSubclass = New cSelfSubHookCallback
    
    m_CursorHandle = 0
    
End Sub

'Initiate input tracking of one or more hWnds.  *Always pass the master hWnd first*, e.g. the one you want used
' as the basis for x/y positioning.
' Aside from the required hWnd, a few other, optional parameters exist:
' - Whether to track MouseMove events as well.  If this is specified, note that VB's intrinsic "MouseMove" event
'    will not fire, as this function will process the messages and raise its own event, preventing VB from doing so.
' - Whether to track AppCommand messages for this hWnd.  This is used for "virtual" functions like forward/back,
'    which are typically generated by mouse buttons 4 and 5, but could also be generated by a multimedia keyboard
'    or gestures on a tablet PC.
'
'By default, all added hWnds will have MouseEnter/Leave/Wheel events tracked "for free".
'
'Note that this function will return TRUE if input tracking was initiated successfully.
Public Function addInputTracker(ByVal targetHWnd As Long, Optional ByVal alsoTrackMouseEnterLeave As Boolean = False, Optional ByVal alsoTrackStandardMouseEvents As Boolean = False, Optional ByVal alsoTrackAppCommands As Boolean = False, Optional ByVal alsoHandleCursorDuties As Boolean = False) As Boolean

    Dim i As Long
        
    'Add the hWnd to our tracking array.  Note that the hWnd in position 0 is treated as the base.
    If targetHWnd <> 0 Then
    
        'Add the hWnd to our tracking array, and increment all subarrays accordingly
        m_hWndCount = m_hWndCount + 1
        ReDim Preserve m_TrackingData(0 To NUM_OF_CUSTOM_INPUT_TRACKERS - 1, 0 To m_hWndCount - 1) As Long
        ReDim Preserve m_MouseTrackingActive(0 To m_hWndCount - 1) As Boolean
        
        m_TrackingData(PDIT_DEFAULT, m_hWndCount - 1) = targetHWnd
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
    If Not cSubclass.ssc_Subclass(targetHWnd, , 1, Me) Then
        
        Debug.Print "WARNING: input subclasser failed to subclass hWnd #" & targetHWnd & ".  Fix it!"
        addInputTracker = False
        Exit Function
    
    End If
        
    'Subclassing was successful.  Add all relevant mouse messages.
    ' (Note that we subclass MouseMove events even if the user doesn't want these events returned; this is
    ' used to generate MouseEnter/Leave events.)
    cSubclass.ssc_AddMsg targetHWnd, MSG_BEFORE, WM_MOUSEWHEEL, WM_MOUSEHWHEEL
    
    'Set any optional event subclassing as well
    If alsoTrackMouseEnterLeave Then cSubclass.ssc_AddMsg targetHWnd, MSG_BEFORE, WM_MOUSELEAVE, WM_MOUSEMOVE, WM_MOUSEHOVER
    If alsoTrackAppCommands Then cSubclass.ssc_AddMsg targetHWnd, MSG_BEFORE, WM_APPCOMMAND
    If alsoTrackStandardMouseEvents Then
        
        'Mouse down...
        cSubclass.ssc_AddMsg targetHWnd, MSG_BEFORE, WM_LBUTTONDOWN, WM_MBUTTONDOWN, WM_RBUTTONDOWN, WM_XBUTTONDOWN
        
        'Mouse up...
        cSubclass.ssc_AddMsg targetHWnd, MSG_BEFORE, WM_LBUTTONUP, WM_MBUTTONUP, WM_RBUTTONUP, WM_XBUTTONUP
        
        'Double-clicks (including X-buttons)...
        cSubclass.ssc_AddMsg targetHWnd, MSG_BEFORE, WM_LBUTTONDBLCLK, WM_MBUTTONDBLCLK, WM_RBUTTONDBLCLK, WM_XBUTTONDBLCLK
        
    End If
    
    'Cursor duties require response to WM_SETCURSOR
    If alsoHandleCursorDuties Then cSubclass.ssc_AddMsg targetHWnd, MSG_BEFORE, WM_SETCURSOR
            
    'If enter, leave, and hover events are desired, start mouse tracking now
    If alsoTrackMouseEnterLeave Then
    
        If requestMouseTrackingForHwnd(targetHWnd) Then
            m_MouseTrackingActive(m_hWndCount - 1) = True
        Else
            
            Debug.Print "WARNING: couldn't get Windows to track mouse events for hWnd #" & targetHWnd & ".  Fix it!"
            addInputTracker = False
            Exit Function
                
        End If
        
    End If
    
    addInputTracker = True
    
End Function

'Once an hWnd has been added to the tracker, callers can use several other functions to expand the tracking capabilities of the class.
' This function is used to request Arrow Key tracking for a given hWnd.  Note that this function does NOT add a new hWnd to the tracker;
' for that reason, I preface the function with "request" instead of "add", because the hWnd must have already been added.
'
'The optional parameter numPadMatters controls whether numLock status should affect reports from the numPad "arrow" keys.  If
' numLockMatters is set to FALSE, this class will report arrow keys from the number pad *regardless of num lock status*.
Public Function requestKeyTracking(ByVal targetHWnd As Long) As Boolean
    
    'Arrow key tracking is a nasty bit of business.  Let me just state that up front before breaking down how PD handles it.
    
    'By default, Windows generally handles arrow keypresses for applications.  Depending on the context, arrow buttons may
    ' be used to navigate between command buttons, or navigate within a text box or edit control, or modify controls in some
    ' other way.
    
    'This default behavior can be overridden by intercepting the WM_GETDLGCODE message and returning a value that specifies
    ' what default behaviors you want to override.  Sounds easy, right?
    
    'In theory, it should be.  But in practice, it is anything but.  Some element of VB (who knows what) eats the WM_GETDLGCODE
    ' message before it ever arrives at our wndProc.  It is possible that there's a way to override this behavior, or
    ' somehow intercept it, but after a great deal of testing I have not found a reliable way to do it for any possible hWnd.
    ' (This could be due to a whole bunch of factors, including the various window bits VB chooses to use, which I didn't want
    ' to custom-handle as the goal here is to support custom tracking for *any* type of hWnd.)
    
    'So, subclassing is out.  That leaves hooking, which is a more troublesome solution, especially in PD's main screen where
    ' a whole bunch of windows are simultaneously present.
    
    'Because hooking is an energy-intensive project, I don't want to hook anything more than is absolutely necessary.  So
    ' hooking is now handled dynamically as any tracked window(s) gain or lose focus.  To accomplish this, key hooking requests
    ' also generate some subclassing bits as well, so we can track focus of our individual windows.
    
    'Subclass two messages: WM_ACTIVATE (for floating windows) and WM_SETFOCUS/KILLFOCUS (for docked windows)
    cSubclass.ssc_AddMsg targetHWnd, MSG_BEFORE, WM_ACTIVATE, WM_SETFOCUS, WM_KILLFOCUS
    
    'Note that we are tracking arrow key events for this hWnd
    Dim hWndIndex As Long
    hWndIndex = getIndexFromHwnd(targetHWnd)
    
    If hWndIndex >= 0 Then
        
        'Note that this hWnd has requested key tracking
        m_TrackingData(PDIT_KEYTRACKING, hWndIndex) = True
        
        'Also reset the hasFocus parameter for this window; it will be set separately by intercepting the WM_ACTIVATE message
        m_TrackingData(PDIT_HASFOCUS, hWndIndex) = False
        
        requestKeyTracking = True
        
    Else
    
        Debug.Print "Key tracking was requested for an hWnd that doesn't exist in the tracker.  Fix it!"
        requestKeyTracking = False
        
    End If
    
End Function

'Once requestKeyTracking (above) has been used to initialize dynamic key hooking for a given hWnd, this setKeyTrackers function
' can be used to specify which key bundles the client wants tracked.  To allow the system to automatically handle certain keypresses,
' pdInput will restrict its handling to *only the keys requested*, so please keep tracking requests to the bare minimum necessary to
' enable whatever behavior is desired.
Public Function setKeyTrackers(ByVal targetHWnd As Long, Optional ByVal trackArrowKeys As Boolean = False, Optional ByVal numLockMatters As Boolean = True, Optional ByVal trackEditKeys As Boolean = False) As Boolean

    'Make sure we are tracking keys for this hWnd
    Dim hWndIndex As Long
    hWndIndex = getIndexFromHwnd(targetHWnd)
    
    If hWndIndex >= 0 Then
    
        'Track arrow keys
        If trackArrowKeys Then
        
            'Unlike other tracking data, arrow keys are a bit different.  A more detailed track state is required, to reflect
            ' how numLock events are treated.
            If numLockMatters Then
                m_TrackingData(PDIT_ARROWKEYS, hWndIndex) = ARROW_TRACKING_OBEY_NUMLOCK
            Else
                m_TrackingData(PDIT_ARROWKEYS, hWndIndex) = ARROW_TRACKING_IGNORE_NUMLOCK
            End If
            
        End If
        
        'Track edit keys (return, space, backspace, ins/del, etc)
        If trackEditKeys Then m_TrackingData(PDIT_EDITKEYS, hWndIndex) = True
        
        setKeyTrackers = True
    
    Else
    
        Debug.Print "You can't set key trackers for an hWnd that doesn't exist in the tracker.  Fix it!"
        setKeyTrackers = False
    
    End If
    
End Function

'Assign a cursor to the hWnd collection.  Call this function without a parameter to reset the cursor to the default arrow.
Public Sub setSystemCursor(Optional ByVal systemCursorType As SystemCursorConstant = IDC_DEFAULT)
    
    'Load the relevant cursor handle
    Dim tmpCursorHandle As Long
    
    If systemCursorType = IDC_DEFAULT Then
        tmpCursorHandle = LoadCursor(0, IDC_ARROW)
    Else
        tmpCursorHandle = LoadCursor(0, systemCursorType)
    End If
    
    'Check for duplicate cursor requests, and ignore them as necessary
    If m_CursorHandle <> tmpCursorHandle Then
        
        m_CursorHandle = tmpCursorHandle
        applyCursorToHwnds False
        
    End If
        
End Sub

'Assign a PNG cursor to the hWnd collection.  This function leans on the Icon and Cursor module to handle the actual
' retrieval, decompression, and assembly of the cursor.  We handle it there so that requested icons can be cached;
' there is a high probability of PNG icons being requested by more than one class in a given session, so rather than
' repeat all the steps for each request, we only do it once, then cache the results.  PD itself will handle unloading
' custom cursors at exit time, so this class need not concern itself with unloading requested cursors.
Public Sub setPNGCursor(ByVal pngResourceName As String, Optional ByVal cursorHotspotX As Long = 0, Optional ByVal cursorHotspotY As Long = 0)
    
    Dim tmpCursorHandle As Long
    tmpCursorHandle = Icon_and_Cursor_Handler.requestCustomCursor(pngResourceName, cursorHotspotX, cursorHotspotY)
    
    If m_CursorHandle <> tmpCursorHandle Then
        m_CursorHandle = tmpCursorHandle
        applyCursorToHwnds False
    End If
    
End Sub

'Private helper function used by setSystemCursor and setPNGCursor, above.  When setting the cursor for the first time
' for our base hWnd, it is helpful to reset the class cursor definition so that Windows doesn't automatically raise
' WM_SETCURSOR messages every time the mouse is moved.  Set the optional parameter to TRUE to cache the existing class
' cursor and overwrite it with NULL.
Private Sub applyCursorToHwnds(Optional ByVal removeExistingClassCursor As Boolean = False)

    'If removeExistingClassCursor is set, use SetClassLong to remove any existing class-level cursor definition
    If removeExistingClassCursor Then m_PrevClassCursorHandle = SetClassLong(m_TrackingData(PDIT_DEFAULT, 0), GCL_HCURSOR, 0&)

    'Apply the new cursor using SetCursor
    SetCursor m_CursorHandle

End Sub

'Partner to applyCursorToHwnds, above.  To be used only when the mouse leaves a cursor-controlled area.
Private Sub resetClassCursor()
    SetClassLong m_TrackingData(PDIT_DEFAULT, 0), GCL_HCURSOR, m_PrevClassCursorHandle
End Sub

'Given an hWnd, populate a generic TrackMouseEvent struct and pass it to the TrackMouseEvent function
Private Function requestMouseTrackingForHwnd(ByVal targetHWnd As Long)

    Dim eventTracker As TRACKMOUSEEVENT_STRUCT
    
    With eventTracker
        .cbSize = Len(eventTracker)
        .dwFlags = TME_LEAVE Or TME_HOVER
        .hWndTrack = targetHWnd
        .dwHoverTime = &HFFFFFFFF
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

'Note that the vKey constant below is a virtual key mapping, not necessarily a standard VB key constant
Private Function isVirtualKeyDown(ByVal vKey As Long) As Boolean
    isVirtualKeyDown = GetAsyncKeyState(vKey) And &H8000
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
    
    isMouseButtonDown = isVirtualKeyDown(vKey)
    
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

'When an arrow key is pressed for a tracked hWnd that supports it, the WM_KEYDOWN message will call this function.
' This function will check all arrow key states relative to the tracking status for this hWnd (e.g. factoring in numLock when relevant).
Private Function reportArrowKeyUpdate(ByVal indexOfHwnd As Long, ByVal modifierKeys As ShiftConstants) As Boolean
    
    'Start by assembling a list of all pressed default arrow keys
    Dim leftArrow As Boolean, rightArrow As Boolean, upArrow As Boolean, downArrow As Boolean
        
    leftArrow = isVirtualKeyDown(VK_LEFT)
    rightArrow = isVirtualKeyDown(VK_RIGHT)
    upArrow = isVirtualKeyDown(VK_UP)
    downArrow = isVirtualKeyDown(VK_DOWN)
    
    'If numLock status has been marked as irrelevant, and an arrow key is not already pressed, investigate number pad status
    ' Similarly, if numLock is relevant, but numLock is DOWN, track status as well.
    If (m_TrackingData(PDIT_ARROWKEYS, indexOfHwnd) = ARROW_TRACKING_IGNORE_NUMLOCK) Or ((m_TrackingData(PDIT_ARROWKEYS, indexOfHwnd) = ARROW_TRACKING_OBEY_NUMLOCK) And Not isVirtualKeyDown(VK_NUMLOCK)) Then
    
        If Not leftArrow Then leftArrow = (isVirtualKeyDown(VK_NUMPAD1) Or isVirtualKeyDown(VK_NUMPAD4) Or isVirtualKeyDown(VK_NUMPAD7))
        If Not rightArrow Then rightArrow = (isVirtualKeyDown(VK_NUMPAD3) Or isVirtualKeyDown(VK_NUMPAD6) Or isVirtualKeyDown(VK_NUMPAD9))
        If Not upArrow Then upArrow = (isVirtualKeyDown(VK_NUMPAD7) Or isVirtualKeyDown(VK_NUMPAD8) Or isVirtualKeyDown(VK_NUMPAD9))
        If Not downArrow Then downArrow = (isVirtualKeyDown(VK_NUMPAD1) Or isVirtualKeyDown(VK_NUMPAD2) Or isVirtualKeyDown(VK_NUMPAD3))
        
    End If
    
    'Raise a corresponding arrow key event
    If upArrow Or rightArrow Or downArrow Or leftArrow Then
        RaiseEvent KeyDownArrows(modifierKeys, upArrow, rightArrow, downArrow, leftArrow)
        reportArrowKeyUpdate = True
    Else
        reportArrowKeyUpdate = False
    End If

End Function

'When an edit key is pressed for a tracked hWnd that supports it, the WM_KEYDOWN message will call this function.
' This function will check all edit key states relative to the tracking status for this hWnd (e.g. factoring in
' the Extended Key Code bit to determine if numpad key variants are being used, which requires a check of the lParam value
' from the hook proc).
Private Function reportEditKeyUpdate(ByVal indexOfHwnd As Long, ByVal modifierKeys As ShiftConstants, ByVal lParam As Long) As Boolean
    
    'Start by assembling a list of all pressed default edit keys
    Dim keyReturn As Boolean, keyEnter As Boolean, keySpaceBar As Boolean, keyBackspace As Boolean
    Dim keyTab As Boolean, keyInsert As Boolean, keyDelete As Boolean
        
    keyReturn = isVirtualKeyDown(VK_RETURN)
    keyEnter = isVirtualKeyDown(VK_RETURN)
    keySpaceBar = isVirtualKeyDown(VK_SPACE)
    keyBackspace = isVirtualKeyDown(VK_BACK)
    keyTab = isVirtualKeyDown(VK_TAB)
    keyInsert = isVirtualKeyDown(VK_INSERT)
    keyDelete = isVirtualKeyDown(VK_DELETE)
    
    'Return and Enter require special handling, because they share the same virtual keycode (brilliant work, MS engineers).
    ' To distinguish between them, we must check the hook event lParam against the "extended key" bit.
    Dim isExtendedKey As Boolean
    If (lParam And KF_EXTENDED) = 0 Then isExtendedKey = False Else isExtendedKey = True
    If keyReturn And isExtendedKey Then keyReturn = False
    If keyEnter And (Not isExtendedKey) Then keyEnter = False
    
    'With all key states correctly tracked, raise a corresponding edit key event
    If keyReturn Or keyEnter Or keySpaceBar Or keyBackspace Or keyTab Or keyInsert Or keyDelete Then
        RaiseEvent KeyDownEdits(modifierKeys, keyReturn, keyEnter, keySpaceBar, keyBackspace, keyInsert, keyDelete, keyTab)
        reportEditKeyUpdate = True
    Else
        reportEditKeyUpdate = False
    End If
    
End Function

'When processing key hooks, our goal is to strip out any keypresses that we are explicitly handling.  If a handled key is found,
' we prevent the keypress from proceeding to other hooks (e.g. pressing the arrow down key should only affect the control that
' currently has focus).  However, if we are *not* handling a keypress, we need to pass it down the chain so other functions can
' have a go at it.
Private Function isRelevantKeyAffected(ByVal vKeyFromHook As Long, ParamArray listOfVKeys() As Variant) As Boolean

    isRelevantKeyAffected = False

    If Not IsMissing(listOfVKeys) Then
        
        Dim i As Long
        For i = LBound(listOfVKeys) To UBound(listOfVKeys)
            If listOfVKeys(i) <> 0 Then
                If vKeyFromHook = CLng(listOfVKeys(i)) Then
                    isRelevantKeyAffected = True
                    Exit For
                End If
            End If
        Next i
        
    End If

End Function

'This routine MUST BE KEPT as the next-to-last routine for this form. Its ordinal position determines its ability to hook properly.
Private Sub myHookProc(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal lHookType As eHookType, ByRef lParamUser As Long)
'*************************************************************************************************
' http://msdn2.microsoft.com/en-us/library/ms644990.aspx
'* bBefore    - Indicates whether the callback is before or after the next hook in chain.
'* bHandled   - In a before next hook in chain callback, setting bHandled to True will prevent the
'*              message being passed to the next hook in chain and (if set to do so).
'* lReturn    - Return value. For Before messages, set per the MSDN documentation for the hook type
'* nCode      - A code the hook procedure uses to determine how to process the message
'* wParam     - Message related data, hook type specific
'* lParam     - Message related data, hook type specific
'* lHookType  - Type of hook calling this callback
'* lParamUser - User-defined callback parameter. Change vartype as needed (i.e., Object, UDT, etc)
'*************************************************************************************************
    
    
    If lHookType = WH_KEYBOARD Then
    
        bHandled = False
        
        'Manually pull key modifier states (shift, control, alt/menu) in advance; these are standard for all key events
        Dim retShiftConstants As ShiftConstants
        If isVirtualKeyDown(VK_SHIFT) Then retShiftConstants = retShiftConstants Or vbShiftMask
        If isVirtualKeyDown(VK_CONTROL) Then retShiftConstants = retShiftConstants Or vbCtrlMask
        If isVirtualKeyDown(VK_ALT) Then retShiftConstants = retShiftConstants Or vbAltMask
    
        'Check for arrow key handling
        If (m_TrackingData(PDIT_ARROWKEYS, 0) = ARROW_TRACKING_OBEY_NUMLOCK) Or (m_TrackingData(PDIT_ARROWKEYS, 0) = ARROW_TRACKING_IGNORE_NUMLOCK) Then
            bHandled = bHandled Or reportArrowKeyUpdate(0, retShiftConstants)
        End If
        
        'Check for edit key handling
        If m_TrackingData(PDIT_EDITKEYS, 0) Then
            bHandled = bHandled Or reportEditKeyUpdate(0, retShiftConstants, lParam)
        End If
        
        'If bHandled Then Debug.Print "pdInput is handling keypress: " & bHandled & ", " & nCode
        
        'Per MSDN, return the value of CallNextHookEx
        If (Not bHandled) Then
            lReturn = CallNextHookEx(0, nCode, wParam, lParam)
        Else
            lReturn = 1
        End If
            
    End If
    
End Sub


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
    
    'Before processing anything, do a failsafe check to make sure the hWnd we've been passed is one we're
    ' actually responsible for.
    Dim indexOfHwnd As Long
    indexOfHwnd = getIndexFromHwnd(lng_hWnd)
    
    'If the hWnd wasn't found, it's possible that it belongs to the override array used for specialized key handling
    ' on user controls (as user controls *fight* to not receive focus, and thus pass focus to any child control they
    ' can get their hands on).  Check the override array for a match; if one was found, treat this window message as if
    ' it applies to the base hWnd.
    If (indexOfHwnd < 0) And (m_numOfOverrideControls > 0) Then
        For i = 0 To m_numOfOverrideControls - 1
        
            If m_OverrideHwnds(i) = lng_hWnd Then
                indexOfHwnd = 0
                Exit For
            End If
        
        Next i
    End If
    
    'The hWnd was found in our tracking array.  Start processing it according to the events requested when
    ' the hWnd was first supplied to this class.
    If (indexOfHwnd >= 0) Then
    
        'As VB already does with its own mouse events, it is helpful to supply key modifiers directly in the event
        ' params.  Unfortunately, the window messages that report key states will only report SHIFT and CTRL masks,
        ' not ALT.  (This is by design, as most windows use the Alt key to forcibly switch focus to the menu, which
        ' in turn makes it irrelevant for mouse modifications.  However, this behavior can be overridden if the
        ' mouse has been captured by a window, which we may end up doing for PD's canvas, and which is why I'm
        ' adding ALT handling here.)
        
        'Anyway, because the Alt key requires special handling, and some messages don't report button state, I just
        ' ignore window message reports entirely and pull key states manually using GetAsyncKeyState.
        Dim retShiftConstants As ShiftConstants
        
        If isVirtualKeyDown(VK_SHIFT) Then retShiftConstants = retShiftConstants Or vbShiftMask
        If isVirtualKeyDown(VK_CONTROL) Then retShiftConstants = retShiftConstants Or vbCtrlMask
        
        'NOTE!  MSDN provides some odd instructions for tracking the ALT key (see Remarks here:
        ' http://msdn.microsoft.com/en-us/library/ms646242%28v=vs.85%29.aspx).  They state explicitly to use
        ' use GetKeyState and not GetAsyncKeyState, but they don't provide any rationale for this.  I have no
        ' trouble with GetAsyncKeyState properly reporting Alt status on Win 7, so I'm leaving this strategy
        ' for now, but I remain intrigued by MSDN's oddly specific instructions on the point...
        If isVirtualKeyDown(VK_ALT) Then retShiftConstants = retShiftConstants Or vbAltMask
        
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
        
        'Regardless of which event we are handling, we now have Shift modifier, mouse button, and mouse coordinate
        ' data (relative to the master hWnd for this class).  This makes individual message parsing much easier.
        
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
                        If m_TrackingData(PDIT_CURSOR, indexOfHwnd) Then applyCursorToHwnds True
                        
                        'Only raise a MouseEnter event if this event is happening for the primary hWnd; otherwise, track
                        ' mouse events silently.
                        '
                        'NOTE! This will raise more MouseEnter events than we want in the case of a UC with multiple
                        '       tracked controls; I'll investigate a better solution if/when it proves necessary.
                        If indexOfHwnd = 0 Then RaiseEvent MouseEnter(retMouseButton, retShiftConstants, mouseX, mouseY)
                        
                    End If
                                        
                    'If custom MouseMove tracking has been requested for this hWnd, raise a mouse move event now
                    If m_TrackingData(PDIT_STANDARD_MOUSE, indexOfHwnd) Then
                        RaiseEvent MouseMoveCustom(retMouseButton, retShiftConstants, mouseX, mouseY)
                        
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
                
                RaiseEvent MouseDownCustom(retMouseButton, retShiftConstants, mouseX, mouseY)
            
            'Mouse up
            Case WM_LBUTTONUP, WM_MBUTTONUP, WM_RBUTTONUP, WM_XBUTTONUP
            
                lReturn = 0
                
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
                
                'X/Y coords are packed into lParam as hi/lowords
                mouseX = lParam And &HFFFF&
                mouseY = lParam \ &H10000
                
                'If the user has not moved the mouse an appreciable amount since the last MouseDown event,
                ' call this a click and raise the corresponding Click() event.
                '
                'Note that a MouseUp event is raised either way, and if a Click event is also being raised,
                ' we warn the user in the MouseUp event.  This way, they can choose to abandon MouseUp
                ' processing if the _Click event is implemented.
                If m_MouseMoveCount < 2 Then
                    
                    RaiseEvent MouseUpCustom(retMouseButton, retShiftConstants, mouseX, mouseY, True)
                    RaiseEvent ClickCustom(retMouseButton, retShiftConstants, mouseX, mouseY)
                
                'The mouse has moved too much, so a Click event will not be raised.  Only raise a MouseUp event.
                Else
                    RaiseEvent MouseUpCustom(retMouseButton, retShiftConstants, mouseX, mouseY, False)
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
                
                lReturn = 1
                bHandled = True
                
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
            
                If m_TrackingData(PDIT_CURSOR, indexOfHwnd) Then applyCursorToHwnds
                
            'A floating toolbox window is being activated or deactivated.  This is important as pdInput uses
            ' activation status to install and remove keyboard hooks.
            Case WM_ACTIVATE
                
                'See if we're tracking special keypresses for a window
                If m_TrackingData(PDIT_KEYTRACKING, indexOfHwnd) Then
                
                    'We are tracking special keypresses for this hWnd.  Figure out if the window is being activated or deactivated.
                    Const WA_INACTIVE As Long = 0
                    Const WA_ACTIVE As Long = 1
                    Const WA_CLICKACTIVE As Long = 2
                    
                    Dim activationStatus As Long
                    activationStatus = wParam And &HFFFF
                    
                    'This hWnd is being deactivated.  If a hook has been installed, remove it now.
                    If activationStatus = WA_INACTIVE Then
                    
                        'Check for an existing hook
                        If m_TrackingData(PDIT_HASFOCUS, indexOfHwnd) Then
                            
                            overrideRequired = False
                            
                            'If the user has specified override controls (as required by a UC),
                            ' compare the KillFocus wParam, which specifies which window is losing focus,
                            ' to our override array.  If a child control *within the user control* is the
                            ' one taking focus, maintain hooking.
                            If m_numOfOverrideControls > 0 Then
                            
                                For i = 0 To m_numOfOverrideControls - 1
                                    If m_OverrideHwnds(i) = wParam Then
                                        overrideRequired = True
                                        Exit For
                                    End If
                                Next i
                            
                            End If
                            
                            'If no override is required, uninstall the hook now
                            If Not overrideRequired Then
                            
                                Debug.Print "Uninstalling keyboard hook for hWnd #" & m_TrackingData(PDIT_DEFAULT, indexOfHwnd) & " due to WM_ACTIVATE"
                                
                                'A hook was found.  Uninstall it now.
                                cSubclass.shk_UnHook WH_KEYBOARD
                                
                                'Note that this window is now inactive
                                m_TrackingData(PDIT_HASFOCUS, indexOfHwnd) = False
                                
                            End If
                        
                        End If
                    
                    'This hWnd is being activated.  Install a hook now.
                    Else
                        
                        'Check for an existing hook
                        If Not m_TrackingData(PDIT_HASFOCUS, indexOfHwnd) Then
                            
                            Debug.Print "Installing keyboard hook for hWnd #" & m_TrackingData(PDIT_DEFAULT, indexOfHwnd) & " due to WM_ACTIVATE"
                            
                            'No hook exists.  Hook the keyboard for this window.
                            cSubclass.shk_SetHook WH_KEYBOARD, False, MSG_BEFORE, m_TrackingData(PDIT_DEFAULT, indexOfHwnd), 2, Me
                            
                            'Note that this window is now active
                            m_TrackingData(PDIT_HASFOCUS, indexOfHwnd) = True
                            
                        End If
                        
                    End If
                
                End If
                
                'Docked toolbox windows and controls will raise WM_SETFOCUS instead of WM_ACTIVATE
                Case WM_SETFOCUS
                    
                    'See if we're tracking special keypresses for this window
                    If m_TrackingData(PDIT_KEYTRACKING, indexOfHwnd) Then
                    
                        'Check for an existing hook
                        If Not m_TrackingData(PDIT_HASFOCUS, indexOfHwnd) Then
                        
                            Debug.Print "Installing keyboard hook for hWnd #" & m_TrackingData(PDIT_DEFAULT, indexOfHwnd) & " due to WM_SETFOCUS"
                            
                            'No hook exists.  Hook the control now.
                            cSubclass.shk_SetHook WH_KEYBOARD, False, MSG_BEFORE, m_TrackingData(PDIT_DEFAULT, indexOfHwnd), 2, Me
                            
                            'Note that this window is now active
                            m_TrackingData(PDIT_HASFOCUS, indexOfHwnd) = True
                            
                        End If
                        
                    End If
                
                Case WM_KILLFOCUS
                    
                    'See if we're tracking special keypresses for this window
                    If m_TrackingData(PDIT_KEYTRACKING, indexOfHwnd) Then
                    
                        'Check for an existing hook
                        If m_TrackingData(PDIT_HASFOCUS, indexOfHwnd) Then
                            
                            overrideRequired = False
                            
                            'If the user has specified override controls (as required by a UC),
                            ' compare the KillFocus wParam, which specifies which window is taking focus,
                            ' to our override array.  If a child control *within the user control* is the
                            ' one taking focus, maintain hooking.
                            If (m_numOfOverrideControls > 0) And wParam <> m_TrackingData(PDIT_DEFAULT, indexOfHwnd) Then
                            
                                For i = 0 To m_numOfOverrideControls - 1
                                    If m_OverrideHwnds(i) = wParam Then
                                        overrideRequired = True
                                        Exit For
                                    End If
                                Next i
                            
                            End If
                            
                            'If no override is required, uninstall the hook now
                            If Not overrideRequired Then
                            
                                Debug.Print "Uninstalling keyboard hook for hWnd #" & m_TrackingData(PDIT_DEFAULT, indexOfHwnd) & " due to WM_KILLFOCUS"
                                
                                cSubclass.shk_UnHook WH_KEYBOARD
                                
                                m_TrackingData(PDIT_HASFOCUS, indexOfHwnd) = False
                                
                            End If
                        
                        End If
                    
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
