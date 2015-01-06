VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdScrollAPI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon API Scroll bar interface
'Copyright 2014-2015 by Tanner Helland
'Created: 12/February/14
'Last updated: 05/January/15
'Last update: modify some window size decisions to better reflect the container's size and orientation
'
'VB's inherent scroll bar controls are an unpleasant piece of UI, for a whole host of reasons: 16-bit limit,
' terrible "has focus" behavior, frequently losing track of the mouse if other events are processing simultaneously,
' so the bar stays depressed and/or moving despite the user no longer interacting with it.
'
'I've finally had it with these limitations (particularly losing track of the mouse), so I put together this class,
' which wraps an all-API scroll bar control around a container hWnd of your choosing (typically a picture box).
' Note that this is different from simply activating default scrollbars on a DC, as we have full control over the
' creation, destruction, positioning and size of the scroll bar.
'
'The following VB project by William Keith served as a helpful starting point for the project:
' http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=62935&lngWId=1
'
'None of William's original code is in use here, but perhaps that project would be useful to individuals looking
' for traditional Windows-managed client-area scroll bars.
'
'A few notes, just in case I forget in the future:
' 1) The class must be initialized via initializeScrollBarWindow() before attempting to set value, min, or max.
' 2) This class handles both horizontal and vertical scrollbars; see the initializeScrollBarWindow params for details.
' 3) Validation is automatically performed on Value, Min, and Max assignments, so don't worry about out-of-range errors.
' 4) Page size is not currently handled by the class.  If you provide an optional page size to the initialization
'     function, the scroll bar's Max value will not be correct.
' 5) Unlike VB scroll bars, this class exposes only a single "Scroll" event.  This event merges both "Scroll" and
'     "Change"-type events.  Similarly, it will also fire when the Value is changed by external assignment, or if
'     the class is forced to modify the Value property to keep it within the bounds of a dynamically changed max/min.
' 6) SmallChange and LargeChange properties can be set via the initialization command.  I haven't added them as
'     standalone properties, but there's nothing preventing me from doing so (save laziness).
' 7) The scroll bar will automatically resize itself to fit its owner, via subclassing.  I did this specifically as
'     many PD scroll bars have unique positioning and size requirements, particularly when embedded in UCs.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit


'Scroll bar-specific API constants
Private Const SB_VERT As Long = 1
Private Const SB_HORZ As Long = 0

Private Const SB_LINELEFT As Long = 0
Private Const SB_LINERIGHT As Long = 1
Private Const SB_PAGELEFT As Long = 2
Private Const SB_PAGERIGHT As Long = 3

Private Const SB_LINEUP As Long = 0
Private Const SB_LINEDOWN As Long = 1
Private Const SB_PAGEUP As Long = 2
Private Const SB_PAGEDOWN As Long = 3

Private Const SB_THUMBPOSITION As Long = 4
Private Const SB_THUMBTRACK As Long = 5

Private Const SB_LEFT As Long = 6
Private Const SB_RIGHT As Long = 7

Private Const SB_BOTTOM As Long = 6
Private Const SB_TOP As Long = 7

Private Const SB_ENDSCROLL As Long = 8

Private Const ESB_ENABLE_BOTH As Long = &H0
Private Const ESB_DISABLE_BOTH As Long = &H3

'Microsoft provides functions with names like GetScrollPos and SetScrollPos, but they are all deprecated in
' favor of GetScrollInfo and SetScrollInfo, which use a SCROLLINFO struct and flags to determine what values
' are written/read.  As such, this struct is hugely important!
Private Type SCROLLINFO
    cbSize As Long
    fMask As Long
    nMin As Long
    nMax As Long
    nPage As Long
    nPos As Long
    nTrackPos As Long
End Type

'Messages to be subclassed
Private Const WM_SIZING As Long = &H214
Private Const WM_SIZE As Long = &H5

Private Const WM_VSCROLL As Long = &H115
Private Const WM_HSCROLL As Long = &H114

'Window creation constants
Private Const WS_CHILD As Long = &H40000000

'SetWindowPos constants (for resizing the scroll bar window)
Private Const SWP_NOACTIVATE As Long = &H10
Private Const SWP_NOZORDER As Long = &H4

'Scrollbar update constants
Private Const SB_CTL As Long = 2
Private Const SIF_RANGE As Long = &H1
Private Const SIF_PAGE As Long = &H2
Private Const SIF_POS As Long = &H4
Private Const SIF_TRACKPOS As Long = &H10

'Handle to the created scroll bar window
Private m_ScrollHwnd As Long

'Handle to the owner (a picture box or user control, typically)
Private m_OwnerHwnd As Long

'Most API calls handle window position and movement in terms of a rect-type variable; we use it here to
' automatically size the scroll bar to match its owner.
Private Type winRect
    x1 As Long
    y1 As Long
    x2 As Long
    y2 As Long
End Type

'APIs
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Any, lpSrc As Any, ByVal byteLength As Long)

Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long

Private Declare Function GetClientRect Lib "user32" (ByVal hndWindow As Long, ByRef lpRect As winRect) As Long

Private Declare Function EnableScrollBar Lib "user32" (ByVal hWnd As Long, ByVal wSBflags As Long, ByVal wArrows As Long) As Long
Private Declare Function GetScrollInfo Lib "user32" (ByVal hWnd As Long, ByVal n As Long, lpScrollInfo As SCROLLINFO) As Long
Private Declare Function SetScrollInfo Lib "user32" (ByVal hWnd As Long, ByVal n As Long, lpcScrollInfo As SCROLLINFO, ByVal fRedraw As Boolean) As Long
Private Declare Function ShowScrollBar Lib "user32" (ByVal hWnd As Long, ByVal wBar As Long, ByVal bShow As Long) As Long

Private Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal bEnable As Long) As Long
Private Declare Sub SetWindowPos Lib "user32" (ByVal targetHWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)

'Additional APIs and constants used to assign the hand cursor to created scroll bars
Private m_PreviousCursor As Long
Private Const GCL_HCURSOR = (-12)
Private Const IDC_HAND As Long = 32649
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function SetClassLong Lib "user32" Alias "SetClassLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

'Subclasser
Private cSubclass As cSelfSubHookCallback

'Current scroll bar values, range, etc
Private m_Value As Long, m_Min As Long, m_Max As Long, m_SmallChange As Long, m_LargeChange As Long
Private m_Enabled As Boolean

'Scroll event (note that no distinction is made between horizontal and vertical scroll events)
Public Event Scroll()

'LBUTTONUP messages are part of a rather elaborate fix for annoying Windows behavior; see the wndProc for details
Private Const WM_LBUTTONUP As Long = &H202

'The scrollbar's orientation is cached at creation time, in case subsequent functions need it
Private m_OrientationHorizontal As Boolean

'INITIALIZE scroll bar interface.  You must also specify an initial set of min/max/value/small change/large change
' values.  I've done this so that scroll bars can be created using a single line of code, rather than having to
' set all properties individually.
'
'IMPORTANT NOTE! You must initialize the scroll bar before attempting to assign it any values, including min/max ones!
Public Sub initializeScrollBarWindow(ByVal ownerHwnd As Long, ByVal OrientationHorizontal As Boolean, ByVal scrMin As Long, ByVal scrMax As Long, ByVal scrValue As Long, ByVal scrSmallChange As Long, ByVal scrLargeChange As Long, Optional ByVal PageSize As Long = -1)

    'Store a local copy of the owner's hWnd
    m_OwnerHwnd = ownerHwnd
    
    'Always default to Enabled
    m_Enabled = True

    'Start by subclassing the parent
    cSubclass.ssc_Subclass ownerHwnd, , , Me
    cSubclass.ssc_AddMsg ownerHwnd, MSG_AFTER, WM_SIZING, WM_SIZE, WM_LBUTTONUP
    
    'To keep subclassing to a minimum, only request scroll events that we actually need for this type of bar
    m_OrientationHorizontal = OrientationHorizontal
    
    If OrientationHorizontal Then
        cSubclass.ssc_AddMsg ownerHwnd, MSG_AFTER, WM_HSCROLL
    Else
        cSubclass.ssc_AddMsg ownerHwnd, MSG_AFTER, WM_VSCROLL
    End If
    
    'Next, retrieve the client size of the owner.  We need this to know how to size the scroll bar.
    Dim ownerRect As winRect
    GetClientRect ownerHwnd, ownerRect
    
    'Create a new scrollbar window
    If OrientationHorizontal Then
        m_ScrollHwnd = CreateWindowEx(ByVal 0&, "SCROLLBAR", "", WS_CHILD, ownerRect.x1, ownerRect.y1, ownerRect.x2, ownerRect.y2, ownerHwnd, 0&, App.hInstance, 0&)
    Else
        m_ScrollHwnd = CreateWindowEx(ByVal 0&, "SCROLLBAR", "", WS_CHILD Or 1, ownerRect.x1, ownerRect.y1, ownerRect.x2, ownerRect.y2, ownerHwnd, 0&, App.hInstance, 0&)
    End If
    
    'As a failsafe against runaway scroll bars (caused by intense processing going on in the background of a
    ' scroll event), subclass mouse button release messages for the scroll bar window.  We will use these to
    ' forcibly prevent endless scrolling when necessary.
    cSubclass.ssc_Subclass m_ScrollHwnd, , , Me
    cSubclass.ssc_AddMsg m_ScrollHwnd, MSG_BEFORE, WM_LBUTTONUP
    
    'Note that we do not need to subclass scroll message for the new scroll bar hWnd!  Scroll events are sent to
    ' the PARENT of the scroll bar, not the bar itself.
    
    'Make local copies of all the scroll bar attributes
    m_Value = scrValue
    m_Min = scrMin
    m_Max = scrMax
    m_SmallChange = scrSmallChange
    m_LargeChange = scrLargeChange
    
    'Next, prepare the scroll bar range information
    Dim scrollData As SCROLLINFO
    
    With scrollData
        .cbSize = Len(scrollData)
        .fMask = SIF_RANGE Or SIF_POS Or IIf(PageSize > -1, SIF_PAGE, 0)
        .nMax = scrMax
        .nMin = scrMin
        .nPos = scrValue
        .nPage = IIf(PageSize > -1, PageSize, 1)
    End With
    
    'Apply the range information to the newly created bar (which will also trigger a redraw)
    SetScrollInfo m_ScrollHwnd, SB_CTL, scrollData, True
    
    'Show the scroll bar
    ShowScrollBar m_ScrollHwnd, SB_CTL, True
    
    'Apply a hand cursor to the scroll bar
    m_PreviousCursor = SetClassLong(m_ScrollHwnd, GCL_HCURSOR, LoadCursor(0, IDC_HAND))

End Sub

'Enable/disable the bar
Public Property Get Enabled() As Boolean
    Enabled = m_Enabled
End Property

Public Property Let Enabled(NewValue As Boolean)

    m_Enabled = NewValue

    If NewValue Then
        EnableScrollBar m_ScrollHwnd, SB_CTL, ESB_ENABLE_BOTH
    Else
        EnableScrollBar m_ScrollHwnd, SB_CTL, ESB_DISABLE_BOTH
    End If

End Property

'Get/Set scroll bar value
Public Property Get Value() As Long
    Value = m_Value
End Property

Public Property Let Value(NewValue As Long)

    'Check range
    If NewValue < m_Min Then NewValue = m_Min
    If NewValue > m_Max Then NewValue = m_Max

    m_Value = NewValue
    
    'Update the scroll bar as well
    Dim scrollData As SCROLLINFO
    
    With scrollData
        .cbSize = Len(scrollData)
        .fMask = SIF_POS
        .nPos = NewValue
    End With
    
    SetScrollInfo m_ScrollHwnd, SB_CTL, scrollData, True
    
    'Raise a scroll event, so the client can respond (if they want)
    RaiseEvent Scroll

End Property

'Get/Set scroll bar minimum
Public Property Get Min() As Long
    Min = m_Min
End Property

Public Property Let Min(NewValue As Long)

    'Check range
    If NewValue > m_Max Then
        Debug.Print "WARNING! Invalid scroll bar minimum value requested!"
        NewValue = m_Max - 1
    End If
    
    m_Min = NewValue
    
    'Similarly, check to make sure our current value is not less than the new minimum
    Dim forcedToChangeValue As Boolean
    forcedToChangeValue = False
    
    If m_Value < m_Min Then
        m_Value = m_Min
        forcedToChangeValue = True
    End If
    
    'Update the scroll bar as well
    Dim scrollData As SCROLLINFO
    
    With scrollData
        .cbSize = Len(scrollData)
        .fMask = SIF_POS Or SIF_RANGE
        .nPos = m_Value
        .nMin = m_Min
        .nMax = m_Max
    End With
    
    SetScrollInfo m_ScrollHwnd, SB_CTL, scrollData, True
    
    'If we were forced to change the value to fit the new minimum, allow the client to respond to the new value
    If forcedToChangeValue Then RaiseEvent Scroll

End Property

'Get/Set scroll bar maximum
Public Property Get Max() As Long
    Max = m_Max
End Property

Public Property Let Max(NewValue As Long)

    'Check range
    If NewValue < m_Min Then
        Debug.Print "WARNING! Invalid scroll bar maximum value requested!"
        NewValue = m_Min + 1
    End If
    
    m_Max = NewValue
    
    'Similarly, check to make sure our current value is not more than the new maximum
    Dim forcedToChangeValue As Boolean
    forcedToChangeValue = False
    
    If m_Value > m_Max Then
        m_Value = m_Max
        forcedToChangeValue = True
    End If
    
    'Update the scroll bar as well
    Dim scrollData As SCROLLINFO
    
    With scrollData
        .cbSize = Len(scrollData)
        .fMask = SIF_POS Or SIF_RANGE
        .nPos = m_Value
        .nMin = m_Min
        .nMax = m_Max
    End With
    
    SetScrollInfo m_ScrollHwnd, SB_CTL, scrollData, True
    
    'If we were forced to change the value to fit the new maximum, allow the client to respond
    If forcedToChangeValue Then RaiseEvent Scroll

End Property

'Get/Set scroll bar LargeChange
Public Property Get LargeChange() As Long
    LargeChange = m_LargeChange
End Property

Public Property Let LargeChange(NewValue As Long)
    m_LargeChange = NewValue
End Property

'INITIALIZE class
Private Sub Class_Initialize()
    Set cSubclass = New cSelfSubHookCallback
End Sub

'TERMINATE class
Private Sub Class_Terminate()

    If m_ScrollHwnd <> 0 Then
    
        'Restore the original scroll bar class cursor
        SetClassLong m_ScrollHwnd, GCL_HCURSOR, m_PreviousCursor
    
        'Stop subclassing the parent
        cSubclass.ssc_UnSubclass m_OwnerHwnd
        
        'Release the subclasser entirely
        cSubclass.ssc_Terminate
        
        'Destroy the scroll bar window we created
        DestroyWindow m_ScrollHwnd
        
    End If
    
End Sub

'Hi/LoWord functions are needed to extract specific actions taken by the user on a scroll bar
Private Function HiWord(ByVal DWord As Long) As Integer
    HiWord = (DWord And &HFFFF0000) \ &H10000
End Function

Private Function LoWord(ByVal DWord As Long) As Integer
    If (DWord And &H8000&) = 0 Then
        LoWord = DWord And &HFFFF&
    Else
        LoWord = DWord Or &HFFFF0000
    End If
End Function

'All events subclassed by this class are processed here.
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

    
    'Because we are subclassing multiple messages, we must sort and process accordingly
    Select Case uMsg
        
        'If intense actions are processing in the background during a scroll event (as is common in PD), the internal
        ' Windows scroll bar handler may fail to stop scrolling the damn scroll bar!  So the mouse becomes stuck in
        ' some kind of neverland, while the bar continues scrolling all on its own, and the only way to stop it is
        ' clicking the scroll bar again - but that may trigger a new never-ending set of scroll events, ad infinitum.
        '
        'I've finally cracked a way to prevent such events, and it was triggered by discovering that Windows will
        ' successfully return a WM_LBUTTONUP event for the scroll bar window in question.  Whenever a mouse up event
        ' is detected, forcibly halt any scrolling by immediately disabling and re-enabling the scroll bar via two
        ' mechanisms: standard window enabling (which clears the input queue), and scrollbar-specific enabling
        ' (which halts existing input processing).  I tried about a thousand different solutions before settling on
        ' these, and I *hope* they fix the problem once and for all.
        Case WM_LBUTTONUP
        
            'Disable the window and the scroll bar internals
            EnableWindow m_ScrollHwnd, 0
            EnableScrollBar m_ScrollHwnd, SB_CTL, ESB_DISABLE_BOTH
            
            'If the scroll bar state is Enabled, re-enable everything.
            If m_Enabled Then EnableScrollBar m_ScrollHwnd, SB_CTL, ESB_ENABLE_BOTH
            If m_Enabled Then EnableWindow m_ScrollHwnd, 1
            
        'Sizing event is happening on the parent.
        Case WM_SIZE, WM_SIZING
        
            'Retrieve the parent's client rect now
            Dim curWinRect As winRect
            GetClientRect m_OwnerHwnd, curWinRect
            
            'Resize the scroll bar accordingly
            If m_OrientationHorizontal Then
                SetWindowPos m_ScrollHwnd, 0, curWinRect.x1 + 1, curWinRect.y1, curWinRect.x2 - 2, curWinRect.y2, SWP_NOZORDER Or SWP_NOACTIVATE   'Or SWP_NOSENDCHANGING
            Else
                SetWindowPos m_ScrollHwnd, 0, curWinRect.x1, curWinRect.y1 + 1, curWinRect.x2, curWinRect.y2 - 2, SWP_NOZORDER Or SWP_NOACTIVATE   'Or SWP_NOSENDCHANGING
            End If
            
        'Scroll bar event!
        Case WM_HSCROLL, WM_VSCROLL
            
            'Fill a scroll struct with information about the current scroll bar state.
            Dim scrollData As SCROLLINFO
            scrollData.cbSize = Len(scrollData)
            scrollData.fMask = SIF_PAGE Or SIF_POS Or SIF_RANGE Or SIF_TRACKPOS
            GetScrollInfo m_ScrollHwnd, SB_CTL, scrollData
            
            'Some events do not require us to fire a _Scroll event.  Use this to check.
            Dim eventRequired As Boolean
            eventRequired = True
            
            'Check scroll type.  Windows returns a dizzying array of types for a scroll bar action, corresponding
            ' to everything from clicking buttons to dragging the scroll nub to mousing down in the dead region
            ' between nub and button.  In order to retrieve values correctly, we must handle each event with a
            ' specific set of code.
            Select Case LoWord(wParam)
            
                'Thumbtrack is a special case corresponding to the "scroll" event in VB.  It is fired repeatedly
                ' while the user drags the mouse.  For it alone, we use a special GetScrollInfo parameter.
                Case SB_THUMBTRACK
                    m_Value = scrollData.nTrackPos
                
                'Thumbposition is fired when the user drags the nub, then release the mouse (similar to a "change"
                ' event in VB).
                Case SB_THUMBPOSITION
                    m_Value = scrollData.nPos
                    
                'Lineleft/right are single unit movements to the right or left
                Case SB_LINELEFT, SB_LINEUP
                    m_Value = scrollData.nPos - m_SmallChange
                    
                Case SB_LINERIGHT, SB_LINEDOWN
                    m_Value = scrollData.nPos + m_SmallChange
                    
                'Pageleft/right are large unit movements to the right or left
                Case SB_PAGELEFT, SB_PAGEUP
                    m_Value = scrollData.nPos - m_LargeChange
                    
                Case SB_PAGERIGHT, SB_PAGEDOWN
                    m_Value = scrollData.nPos + m_LargeChange
                    
                'Left/right are jumps to min/max.  To be honest, I'm not sure when these events would fire.
                Case SB_LEFT, SB_TOP
                    m_Value = m_Min

                Case SB_RIGHT, SB_BOTTOM
                    m_Value = m_Max

                'Some miscellaneous actions may also fire (e.g. SB_ENDSCROLL).  If that happens, use
                ' the current nub position as the value, but do not raise a corresponding Scroll event.
                Case Else
                    m_Value = scrollData.nPos
                    eventRequired = False
                    
            End Select
            
            'Validate the new value
            If m_Value < m_Min Then m_Value = m_Min
            If m_Value > m_Max Then m_Value = m_Max
            
            'Apply the new value to the scroll bar struct
            scrollData.fMask = SIF_POS
            scrollData.nPos = m_Value
            
            'Submit the struct to the scroll bar so Windows can update its records accordingly
            SetScrollInfo m_ScrollHwnd, SB_CTL, scrollData, True
            
            'Finally, raise a Scroll event so the client can respond accordingly
            If eventRequired Then RaiseEvent Scroll
            
    End Select
            
    

' *************************************************************
' C A U T I O N   C A U T I O N   C A U T I O N   C A U T I O N
' -------------------------------------------------------------
' DO NOT ADD ANY OTHER CODE BELOW THE "END SUB" STATEMENT BELOW
'   add this warning banner to the last routine in your class
' *************************************************************
End Sub







