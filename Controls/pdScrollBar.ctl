VERSION 5.00
Begin VB.UserControl pdScrollBar 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ClipBehavior    =   0  'None
   DrawStyle       =   5  'Transparent
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HasDC           =   0   'False
   HitBehavior     =   0  'None
   PaletteMode     =   4  'None
   ScaleHeight     =   40
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "pdScrollBar.ctx":0000
End
Attribute VB_Name = "pdScrollBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Scrollbar control
'Copyright 2015-2026 by Tanner Helland
'Created: 07/October/15
'Last updated: 29/May/19
'Last update: switch to pdPopupMenu for right-clicks; this allows for localization (finally)
'
'In a surprise to precisely no one, PhotoDemon has some unique needs when it comes to user controls - needs that
' the intrinsic VB controls can't handle.  These range from the obnoxious (lack of an "autosize" property for
' anything but labels) to the critical (no Unicode support).
'
'As such, I've created many of my own UCs for the program.  All are owner-drawn, with the goal of maintaining
' visual fidelity across the program, while also enabling key features like Unicode support.
'
'A few notes on this scroll bar control, specifically:
'
' 1) Unlike traditional scroll bars, only a single "Scroll" event is raised (vs Scroll and Change events).  This event
'    includes a parameter that lets you know whether the event is "crucial", e.g. whether VB would call it a "Change"
'    event instead of a "Scroll" event.
' 2) High DPI settings are handled automatically.
' 3) A hand cursor is automatically applied.
' 4) Visual appearance is automatically handled by PD's central theming engine.
' 5) This control represents both horizontal and vertical orientations.  Set the corresponding property to match,
'     but be forwarned that this does *not* automatically change the control's size to match!  This is by design.
'     (Although I don't know why it would ever be wise to do this, note that you can technically change orientation
'      at run-time, without penalty, as a side-effect of this decision.)
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'This control really only needs one event raised - Scroll.  The "eventIsCritical" parameter can
' optionally be tested; it returns FALSE for events that would be considered a "scroll" by VB
' (e.g. click-dragging), which you could theoretically ignore if you were worried about performance.
' If eventIsCritical is TRUE, however, you *must* respond to the event.
Public Event Scroll(ByVal eventIsCritical As Boolean)

'Because VB focus events are wonky, especially when we use CreateWindow within a UC, this control raises its own
' specialized focus events.  If you need to track focus, use these instead of the default VB functions.
Public Event GotFocusAPI()
Public Event LostFocusAPI()

'To enable specialized behavior of the ESC key on the primary canvas scrollbars, we also raise a special
' "system" key event IFF the scrollbar is not sited on a modal dialog.
' (For additional details, see https://github.com/tannerhelland/PhotoDemon/issues/476 )
Public Event KeyDownSystem(ByVal Shift As ShiftConstants, ByVal whichSysKey As PD_NavigationKey, ByRef markEventHandled As Boolean)

'If the mouse is currently INSIDE the control, this will be set to TRUE; this affects control rendering
Private m_MouseInsideUC As Boolean

'The scrollbar's orientation is cached at creation time, in case subsequent functions need it
Private m_OrientationHorizontal As Boolean

'Current scroll bar values, range, etc.  Note that the "small change" property is not used as
' a raw value - instead, it is modified based on the current significant digit setting.
' (Specifically, "1" maps to "1" when significant digits is 0, "0.1" when significant digits is 1, etc.)
Private m_Value As Double, m_Min As Double, m_Max As Double, m_SmallChange As Double, m_LargeChange As Double

'The number of significant digits for this control.  0 means integer values.
Private m_SignificantDigits As Long

'To simplify mouse_down handling, resize events fill three rects: one for the "up" or "left" scroll button, one for
' the "down" or "right" scroll button, and a third one, for the track rect between the buttons.
Private upLeftRect As RectL, downRightRect As RectL, trackRect As RectL

'Max/min property changes fill a third rect - the "thumb" rect - which is the bar in the middle of the scroll bar.
' Note that the thumb rect is a RECTF, because it supports subpixel positioning.
Private thumbRect As RectF

'To simplify thumb calculations, we calculate its size only when necessary, and cache it.  Note that the size
' is directionless (e.g. it represents the height for a vertical Thumb, and the width for a horizontal one).
Private m_ThumbSize As Single

'Mouse state for the up/down buttons and thumb.  These are important for not just rendering, but in the case of the buttons,
' for controlling a timer that enables "press and hold" behavior.
Private m_MouseDownThumb As Boolean, m_MouseOverThumb As Boolean
Private m_MouseDownTrack As Boolean, m_MouseOverTrack As Boolean
Private m_MouseDownUpButton As Boolean, m_MouseDownDownButton As Boolean
Private m_MouseOverUpButton As Boolean, m_MouseOverDownButton As Boolean

'When the user click-drags the thumb, we need to store a couple of offsets at MouseDown time, to ensure that the thumb moves
' relative to the initial mouse position (as opposed to "snapping" its top-left point to the new mouse coordinates).
Private m_InitMouseX As Single, m_InitMouseY As Single, m_initValue As Double, m_initMouseValue As Double

'Similarly, when the user mouse-downs on the track (but *not* the thumb), we want to cache those values uniquely.
' Once the thumb meets that location, we can turn off the timers.
Private m_TrackX As Single, m_TrackY As Single, m_initTrackValue As Double

'When right-clicking to raise a scrollbar context menu, we need to cache the current (x, y) values for the
' "Scroll here" context menu option.  These are kept separate from the m_InitMouseX and m_InitMouseY values, above,
' for the extremely rare occasion where the user right-clicks when the LMB is already down.
Private m_ContextMenuX As Single, m_ContextMenuY As Single

'The scrollbars around the main canvas are colored a little differently, by design.  Rather than exposing a crapload of
' color properties, a single "VisualStyle" property is exposed, and it controls all colors during painting.
Public Enum ScrollBarVisualStyle
    SBVS_Standard = 0
    SBVS_Canvas = 1
End Enum

#If False Then
    Private Const SBVS_Standard = 0, SBVS_Canvas = 1
#End If

Private m_VisualStyle As ScrollBarVisualStyle

'User control support class.  Historically, many classes (and associated subclassers) were required by each user control,
' but I've since wrapped these into a single central support class.
Private WithEvents ucSupport As pdUCSupport
Attribute ucSupport.VB_VarHelpID = -1

'Local list of themable colors.  This list includes all potential colors used by the control, regardless of state change
' or internal control settings.  The list is updated by calling the UpdateColorList function.
' (Note also that this list does not include variants, e.g. "BorderColor" vs "BorderColor_Hovered".  Variant values are
'  automatically calculated by the color management class, and they are retrieved by passing boolean modifiers to that
'  class, rather than treating every imaginable variant as a separate constant.)
Private Enum PDSCROLL_COLOR_LIST
    [_First] = 0
    PDS_Track = 0
    PDS_ThumbBorder = 1
    PDS_ThumbFill = 2
    PDS_ButtonBorder = 3
    PDS_ButtonFill = 4
    PDS_ButtonArrow = 5
    [_Last] = 5
    [_Count] = 6
End Enum

'Color retrieval and storage is handled by a dedicated class; this allows us to optimize theme interactions,
' without worrying about the details locally.
Private m_Colors As pdThemeColors

'To mimic standard scroll bar behavior, we must fire repeat scroll events when the buttons (or track) are clicked and held.
Private WithEvents m_UpButtonTimer As pdTimer
Attribute m_UpButtonTimer.VB_VarHelpID = -1
Private WithEvents m_DownButtonTimer As pdTimer
Attribute m_DownButtonTimer.VB_VarHelpID = -1

'Popup menu exposed on right-clicks (this is adopted from normal Windows scroll bars)
Private WithEvents m_PopupMenu As pdPopupMenu
Attribute m_PopupMenu.VB_VarHelpID = -1

Public Function GetControlType() As PD_ControlType
    GetControlType = pdct_ScrollBar
End Function

Public Function GetControlName() As String
    GetControlName = UserControl.Extender.Name
End Function

'Container hWnd must be exposed for external tooltip handling
Public Property Get ContainerHwnd() As Long
    ContainerHwnd = UserControl.ContainerHwnd
End Property

'hWnds aren't exposed by default
Public Property Get hWnd() As Long
Attribute hWnd.VB_UserMemId = -515
    hWnd = UserControl.hWnd
End Property

'The Enabled property is a bit unique; see http://msdn.microsoft.com/en-us/library/aa261357%28v=vs.60%29.aspx
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_UserMemId = -514
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal newValue As Boolean)
    UserControl.Enabled = newValue
    PropertyChanged "Enabled"
    RedrawBackBuffer
End Property

Public Property Get LargeChange() As Long
    LargeChange = m_LargeChange
End Property

Public Property Let LargeChange(ByVal newValue As Long)
    m_LargeChange = newValue
    PropertyChanged "LargeChange"
End Property

'Note: the control's maximum value is settable at run-time; the Value property will automatically be brought in-bounds,
' as necessary.
Public Property Get Max() As Double
    Max = m_Max
End Property

Public Property Let Max(ByVal newValue As Double)
        
    m_Max = newValue
    
    'If the current control .Value is greater than the new max, change it to match
    If (m_Value > m_Max) Then
        m_Value = m_Max
        RaiseEvent Scroll(True)
    End If
    
    'Recalculate thumb size and position
    DetermineThumbSize
    If PDMain.IsProgramRunning() Then RedrawBackBuffer True
    
    PropertyChanged "Max"
    
End Property

'Note: the control's minimum value is settable at run-time; the Value property will automatically be brought in-bounds,
' as necessary.
Public Property Get Min() As Double
    Min = m_Min
End Property

Public Property Let Min(ByVal newValue As Double)
        
    m_Min = newValue
    
    'If the current control .Value is less than the new minimum, change it to match
    If (m_Value < m_Min) Then
        m_Value = m_Min
        RaiseEvent Scroll(True)
    End If
    
    'Recalculate thumb size and position, then redraw the button to match
    DetermineThumbSize
    If PDMain.IsProgramRunning() Then RedrawBackBuffer True
    
    PropertyChanged "Min"
    
End Property

'Unlike system scroll bars, PD provides horizontal and visual scrollbars from the same control.  You can change this
' style at run-time, but note that the control does not resize itself, by design.  You must manually resize the control
' to match the new orientation.
Public Property Get OrientationHorizontal() As Boolean
    OrientationHorizontal = m_OrientationHorizontal
End Property

Public Property Let OrientationHorizontal(ByVal newState As Boolean)
    
    If (m_OrientationHorizontal <> newState) Then
        
        m_OrientationHorizontal = newState
        
        'Update the positioning of the buttons, track, thumb, etc
        UpdateControlLayout
        PropertyChanged "OrientationHorizontal"
        
    End If
    
End Property

'Significant digits determines whether the control allows float values or int values (and with how much precision)
Public Property Get SigDigits() As Long
    SigDigits = m_SignificantDigits
End Property

Public Property Let SigDigits(ByVal newValue As Long)
    m_SignificantDigits = newValue
    PropertyChanged "SigDigits"
End Property

Public Property Get SmallChange() As Long
    SmallChange = m_SmallChange
End Property

Public Property Let SmallChange(ByVal newValue As Long)
    m_SmallChange = newValue
    PropertyChanged "SmallChange"
End Property

'Value supports floating-point or integer values, but it is always stored and returned as a Double-type.  PD will automatically
' manage accuracy for you; set the SigDigits property to control the resolution of the scrollbar.
Public Property Get Value() As Double
    Value = m_Value
End Property

Public Property Let Value(ByVal newValue As Double)
    
    'For integer-only scroll bars, clamp values to their integer range
    If (m_SignificantDigits = 0) Then newValue = Int(newValue)
    
    'Don't make any changes unless the new value deviates from the existing one
    If (newValue <> m_Value) Then
        
        m_Value = newValue
        
        'While running, perform bounds-checking.  (It's less important in the designer, as the assumption is that the
        ' developer will momentarily bring everything into order.)
        If PDMain.IsProgramRunning() Then
            
            'To prevent RTEs, perform an additional bounds check.  Clamp the value if it lies outside control boundaries.
            If (m_Value < m_Min) Then m_Value = m_Min
            If (m_Value > m_Max) Then m_Value = m_Max
            
        End If
        
        'Recalculate the current thumb position, then redraw the button (and force an immediate refresh)
        DetermineThumbSize
        RedrawBackBuffer True
        
        'Mark the value property as being changed, and raise the corresponding event.
        PropertyChanged "Value"
        RaiseEvent Scroll(Not m_MouseDownThumb)
        
    End If
                
End Property

'Visual style controls colors (but nothing else, at present)
Public Property Get VisualStyle() As ScrollBarVisualStyle
    VisualStyle = m_VisualStyle
End Property

Public Property Let VisualStyle(ByVal newStyle As ScrollBarVisualStyle)
    
    If (newStyle <> m_VisualStyle) Then
        m_VisualStyle = newStyle
        UpdateColorList
        RedrawBackBuffer
        PropertyChanged "VisualStyle"
    End If
    
End Property

'To support high-DPI settings properly, we expose specialized move+size functions
Public Function GetLeft() As Long
    GetLeft = ucSupport.GetControlLeft
End Function

Public Sub SetLeft(ByVal newLeft As Long)
    ucSupport.RequestNewPosition newLeft, , True
End Sub

Public Function GetTop() As Long
    GetTop = ucSupport.GetControlTop
End Function

Public Sub SetTop(ByVal newTop As Long)
    ucSupport.RequestNewPosition , newTop, True
End Sub

Public Function GetWidth() As Long
    GetWidth = ucSupport.GetControlWidth
End Function

Public Sub SetWidth(ByVal newWidth As Long)
    ucSupport.RequestNewSize newWidth, , True
End Sub

Public Function GetHeight() As Long
    GetHeight = ucSupport.GetControlHeight
End Function

Public Sub SetHeight(ByVal newHeight As Long)
    ucSupport.RequestNewSize , newHeight, True
End Sub

Public Sub SetPositionAndSize(ByVal newLeft As Long, ByVal newTop As Long, ByVal newWidth As Long, ByVal newHeight As Long)
    ucSupport.RequestFullMove newLeft, newTop, newWidth, newHeight, True
End Sub

Public Property Get HasFocus() As Boolean
    HasFocus = ucSupport.DoIHaveFocus()
End Property

'Timers control repeat value changes when the mouse is held down on an up/down button
Private Sub m_DownButtonTimer_Timer()

    'If this is the first time the button is firing, we want to reset the button's interval to the repeat rate instead
    ' of the delay rate.
    If (m_DownButtonTimer.Interval = Interface.GetKeyboardDelay * 1000) Then
        m_DownButtonTimer.Interval = Interface.GetKeyboardRepeatRate * 1000
    End If
    
    'It's a little counter-intuitive, but the DOWN button actually moves the control value UP
    MoveValueUp m_MouseDownTrack
    
    'If the timer was activated because the user is clicking on the mouse track (and not a button), deactivate the
    ' timer once the value equals the value under the mouse.
    If m_MouseDownTrack Then
        If PDMath.IsPointInRectF(m_TrackX, m_TrackY, thumbRect) Or (m_Value > m_initTrackValue) Then m_DownButtonTimer.StopTimer
    End If

End Sub

Private Sub m_PopupMenu_MenuClicked(ByRef clickedMenuID As String, ByVal idxMenuTop As Long, ByVal idxMenuSub As Long)
    
    Select Case idxMenuTop
        
        'Scroll here
        Case 0
            'Change the value to the corresponding value of the context menu position
            Value = GetValueFromMouseCoords(m_ContextMenuX, m_ContextMenuY)
            
        '(separator)
        Case 1
        
        'Top
        Case 2
            Value = Min
        
        'Bottom
        Case 3
            Value = Max
        
        '(separator)
        Case 4
        
        'Page up
        Case 5
            MoveValueDown True
            
        'Page down
        Case 6
            MoveValueUp True
        
        '(separator)
        Case 7
        
        'Scroll up
        Case 8
            MoveValueDown
        
        'Scroll down
        Case 9
            MoveValueUp
    
    End Select
    
End Sub

Private Sub m_UpButtonTimer_Timer()

    'If this is the first time the button is firing, we want to reset the button's interval to the repeat rate instead
    ' of the delay rate.
    If (m_UpButtonTimer.Interval = Interface.GetKeyboardDelay * 1000) Then
        m_UpButtonTimer.Interval = Interface.GetKeyboardRepeatRate * 1000
    End If
    
    'It's a little counter-intuitive, but the UP button actually moves the control value DOWN
    MoveValueDown m_MouseDownTrack
    
    'If the timer was activated because the user is clicking on the mouse track (and not a button), deactivate the
    ' timer once the value equals the value under the mouse.
    If m_MouseDownTrack Then
        If PDMath.IsPointInRectF(m_TrackX, m_TrackY, thumbRect) Or (m_Value < m_initTrackValue) Then m_UpButtonTimer.StopTimer
    End If
    
End Sub

'When the control receives focus, if the focus isn't received via mouse click, display a focus rect
Private Sub ucSupport_GotFocusAPI()
    RedrawBackBuffer
    RaiseEvent GotFocusAPI
End Sub

Private Sub ucSupport_KeyDownSystem(ByVal Shift As ShiftConstants, ByVal whichSysKey As PD_NavigationKey, markEventHandled As Boolean)
    
    'Enter/Esc get reported directly to the system key handler.  Note that we track the return, because TRUE
    ' means the key was successfully forwarded to the relevant handler.  (If FALSE is returned, no control
    ' accepted the keypress, meaning we should forward the event down the line.)
    markEventHandled = NavKey.NotifyNavKeypress(Me, whichSysKey, Shift)
    
    'If the event was *NOT* handled (meaning this scroll bar is *NOT* sited on a modal dialog), allow the owner
    ' to respond to the event in some custom way.
    If (Not markEventHandled) Then RaiseEvent KeyDownSystem(Shift, whichSysKey, markEventHandled)
    
End Sub

'When the control loses focus, erase any focus rects it may have active
Private Sub ucSupport_LostFocusAPI()
    MakeLostFocusUIChanges
    RaiseEvent LostFocusAPI
End Sub

Private Sub MakeLostFocusUIChanges()
    m_MouseInsideUC = False
    m_MouseOverUpButton = False
    m_MouseOverDownButton = False
    m_MouseOverThumb = False
    m_MouseOverTrack = False
    RedrawBackBuffer
End Sub

'A few key events are also handled
Private Sub ucSupport_KeyDownCustom(ByVal Shift As ShiftConstants, ByVal vkCode As Long, markEventHandled As Boolean)
    
    markEventHandled = False
    
    'Only process key events if this control has focus
    If m_MouseInsideUC Or ucSupport.DoIHaveFocus Then
        
        If (vkCode = VK_UP) Or (vkCode = VK_LEFT) Then
            MoveValueDown
            markEventHandled = True
        ElseIf (vkCode = VK_DOWN) Or (vkCode = VK_RIGHT) Then
            MoveValueUp
            markEventHandled = True
        ElseIf (vkCode = VK_PAGEUP) Then
            MoveValueDown True
            markEventHandled = True
        ElseIf (vkCode = VK_PAGEDOWN) Then
            MoveValueUp True
            markEventHandled = True
        ElseIf (vkCode = VK_HOME) Then
            Value = Min
            markEventHandled = True
        ElseIf (vkCode = VK_END) Then
            Value = Max
            markEventHandled = True
        End If
        
    End If

End Sub

'Only left clicks raise Click() events
Private Sub ucSupport_MouseDownCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal timeStamp As Long)
    
    If Me.Enabled Then
        
        'Separate further handling by button
        Select Case Button
        
            Case pdLeftButton
                
                'Determine mouse button state for the up and down button areas
                If PDMath.IsPointInRectL(x, y, upLeftRect) Then
                    m_MouseDownUpButton = True
                    
                    'Adjust the value immediately
                    MoveValueDown
                    
                    'Start the repeat timer as well
                    m_UpButtonTimer.Interval = Interface.GetKeyboardDelay() * 1000
                    m_UpButtonTimer.StartTimer
                    
                Else
                    m_MouseDownUpButton = False
                End If
                
                If PDMath.IsPointInRectL(x, y, downRightRect) Then
                    m_MouseDownDownButton = True
                    MoveValueUp
                    m_DownButtonTimer.Interval = Interface.GetKeyboardDelay() * 1000
                    m_DownButtonTimer.StartTimer
                Else
                    m_MouseDownDownButton = False
                End If
                
                'Determine button state for the thumb
                If PDMath.IsPointInRectF(x, y, thumbRect) Then
                    m_MouseDownThumb = True
                    
                    'Store initial x/y/value values at this location
                    m_InitMouseX = x
                    m_InitMouseY = y
                    m_initValue = m_Value
                    m_initMouseValue = GetValueFromMouseCoords(x, y)
                    
                Else
                
                    m_MouseDownThumb = False
                    
                    'Now we perform a special check for the mouse being inside the track area.  (We do it here so that
                    ' the mouse being over the thumb (which lies *inside* the track) doesn't set this to TRUE.)
                    If PDMath.IsPointInRectL(x, y, trackRect) Then
                        
                        m_MouseDownTrack = True
                        
                        'Cache the mouse positions, so we know when to deactivate the associated timers
                        m_TrackX = x
                        m_TrackY = y
                        m_initTrackValue = GetValueFromMouseCoords(x, y, True)
                        
                        'Activate the auto-scroll timers
                        If m_initTrackValue < m_Value Then
                            MoveValueDown True
                            m_UpButtonTimer.Interval = Interface.GetKeyboardDelay() * 1000
                            m_UpButtonTimer.StartTimer
                        Else
                            MoveValueUp True
                            m_DownButtonTimer.Interval = Interface.GetKeyboardDelay() * 1000
                            m_DownButtonTimer.StartTimer
                        End If
                        
                    Else
                        m_MouseDownTrack = False
                    End If
                    
                End If
                
                'Request a redraw
                RedrawBackBuffer
                    
            'Right button raises the default scroll context menu
            Case pdRightButton
                
                'Cache the current (x, y) values, because the context menu needs them for the "scroll here" option
                m_ContextMenuX = x
                m_ContextMenuY = y
                
                'Make sure translations have been applied to the popup menu captions
                If (Not m_PopupMenu Is Nothing) Then m_PopupMenu.ShowMenu Me.hWnd, m_ContextMenuX, m_ContextMenuY
                
        End Select
        
    'End (If Me.Enabled...)
    End If
    
End Sub

Private Sub ucSupport_MouseEnter(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    m_MouseInsideUC = True
    ucSupport.RequestCursor IDC_HAND
End Sub

'When the mouse leaves the UC, we must repaint the button (as it's no longer hovered)
Private Sub ucSupport_MouseLeave(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    
    If m_MouseInsideUC Then
        
        m_MouseOverUpButton = False
        m_MouseOverDownButton = False
        m_MouseOverThumb = False
        m_MouseOverTrack = False
        
        m_MouseInsideUC = False
        RedrawBackBuffer
        
    End If
    
    'Reset the cursor
    ucSupport.RequestCursor IDC_ARROW
    
End Sub

'When the mouse enters the button, we must initiate a repaint (to reflect its hovered state)
Private Sub ucSupport_MouseMoveCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal timeStamp As Long)
    
    'Reset mouse capture behavior; this greatly simplifies parts of the drawing function
    If (Not m_MouseInsideUC) Then m_MouseInsideUC = True
    
    'If the user is click-dragging the thumb, we give that preferential treatment
    If m_MouseDownThumb Then
        
        'Figure out a new value for the current mouse position
        Dim curValue As Double, valDiff As Double
        curValue = GetValueFromMouseCoords(x, y)
        
        'Solve for the difference between this value and the initial MouseDown value
        valDiff = curValue - m_initMouseValue
        
        'Set the actual control value to match; this assignment will handle redraws as necessary
        Me.Value = m_initValue + valDiff
        
    Else
    
        'Determine mouse hover state for the up and down button areas
        If PDMath.IsPointInRectL(x, y, upLeftRect) Then
            m_MouseOverUpButton = True
            m_MouseOverTrack = False
        Else
            m_MouseOverUpButton = False
        End If
        
        If PDMath.IsPointInRectL(x, y, downRightRect) Then
            m_MouseOverDownButton = True
            m_MouseOverTrack = False
        Else
            m_MouseOverDownButton = False
        End If
            
        If PDMath.IsPointInRectF(x, y, thumbRect) Then
            m_MouseOverThumb = True
            m_MouseOverTrack = False
        Else
            m_MouseOverThumb = False
            
            'Do a special check for the track now
            If PDMath.IsPointInRectL(x, y, trackRect) Then
                m_MouseOverTrack = True
                
                'Cache the mouse positions, so we know where to draw the orientation dot
                m_TrackX = x
                m_TrackY = y
                m_initTrackValue = GetValueFromMouseCoords(x, y, True)
            
            Else
                m_MouseOverTrack = False
            End If
            
        End If
        
        'Repaint the control
        RedrawBackBuffer
        
    End If
    
End Sub

Private Sub ucSupport_MouseUpCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal clickEventAlsoFiring As Boolean, ByVal timeStamp As Long)
    
    If (Button = pdLeftButton) Then
        
        m_MouseDownUpButton = False
        m_MouseDownDownButton = False
        m_MouseDownThumb = False
        m_MouseDownTrack = False
        
        m_UpButtonTimer.StopTimer
        m_DownButtonTimer.StopTimer
        
        'When the mouse is released, raise a final "Scroll" event with the crucial parameter set to TRUE, which lets the
        ' caller know that they can perform any long-running actions now.
        RaiseEvent Scroll(True)
        
        'Request a redraw
        RedrawBackBuffer
        
    End If
    
End Sub

Private Sub ucSupport_MouseWheelHorizontal(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal scrollAmount As Double)
    RelayMouseWheelEvent False, Button, Shift, x, y, scrollAmount
End Sub

Private Sub ucSupport_MouseWheelVertical(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal scrollAmount As Double)
    RelayMouseWheelEvent True, Button, Shift, x, y, scrollAmount
End Sub

'If some external window wants the scrollbar to automatically sync to its own wheel events, it can use this wrapper function.
Public Sub RelayMouseWheelEvent(ByVal wheelIsVertical As Boolean, ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal scrollAmount As Double)
    
    If (scrollAmount <> 0) Then
        
        'For convenience, swap wheel direction for horizontal wheel actions
        If (Not wheelIsVertical) Then scrollAmount = -1 * scrollAmount
        
        If scrollAmount > 0 Then
            MoveValueDown True
        Else
            MoveValueUp True
        End If
        
        'If the mouse is over the scroll bar, wheel actions may cause the thumb to move into (and/or out of) the
        ' cursor's position.  As such, we must update that value here.
        If (m_MouseOverThumb <> IsPointInRectF(x, y, thumbRect)) Then
            m_MouseOverThumb = Not m_MouseOverThumb
            RedrawBackBuffer
        End If
        
    End If
    
End Sub

'Mousewheel zoom (Ctrl+scroll) isn't a relevant scroll bar command.  If we receive a zoom event, assume the user wants it
' relayed to the currently active canvas (with appropriate checks for viewport unavailability - e.g. an active modal dialog).
' This change addresses https://github.com/tannerhelland/PhotoDemon/issues/476
Private Sub ucSupport_MouseWheelZoom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal zoomAmount As Double)
    
    'If a modal dialog is active, disregard
    If (FormMain.MainCanvas(0).IsCanvasInteractionAllowed() And FormMain.MainCanvas(0).IsZoomEnabled And PDImages.IsImageActive()) Then
        
        'Forward the zoom command to the central zoom handler, using the current center point of the canvas
        ' as the zoom focal point.
        If (zoomAmount <> 0) Then Tools_Zoom.RelayCanvasZoom FormMain.MainCanvas(0), PDImages.GetActiveImage(), FormMain.MainCanvas(0).GetCanvasWidth / 2, FormMain.MainCanvas(0).GetCanvasHeight / 2, (zoomAmount > 0)
        
    End If
    
End Sub

Private Sub ucSupport_RepaintRequired(ByVal updateLayoutToo As Boolean)
    If updateLayoutToo Then UpdateControlLayout Else RedrawBackBuffer
End Sub

'INITIALIZE control
Private Sub UserControl_Initialize()
    
    'Initialize a user control support class
    Set ucSupport = New pdUCSupport
    ucSupport.RegisterControl UserControl.hWnd, True
    ucSupport.RequestExtraFunctionality True, True
    ucSupport.SpecifyRequiredKeys VK_UP, VK_DOWN, VK_RIGHT, VK_LEFT, VK_END, VK_HOME, VK_PAGEUP, VK_PAGEDOWN
    
    m_MouseInsideUC = False
    
    'Prep the color manager and load default colors
    Set m_Colors = New pdThemeColors
    Dim colorCount As PDSCROLL_COLOR_LIST: colorCount = [_Count]
    m_Colors.InitializeColorList "PDScrollBar", colorCount
    If Not PDMain.IsProgramRunning() Then UpdateColorList
    
    'Prep timer objects
    If PDMain.IsProgramRunning() Then
        Set m_UpButtonTimer = New pdTimer
        Set m_DownButtonTimer = New pdTimer
    End If
    
End Sub

'Set default properties
Private Sub UserControl_InitProperties()
    Min = 0
    Max = 10
    Value = 0
    LargeChange = 1
    SigDigits = 0
    SmallChange = 1
    OrientationHorizontal = False
    VisualStyle = SBVS_Standard
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    
    With PropBag
        m_Min = .ReadProperty("Min", 0)
        m_Max = .ReadProperty("Max", 10)
        Value = .ReadProperty("Value", 0)
        LargeChange = .ReadProperty("LargeChange", 1)
        SmallChange = .ReadProperty("SmallChange", 1)
        SigDigits = .ReadProperty("SignificantDigits", 0)
        m_OrientationHorizontal = .ReadProperty("OrientationHorizontal", False)
        m_VisualStyle = .ReadProperty("VisualStyle", SBVS_Standard)
    End With

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    With PropBag
        .WriteProperty "Min", m_Min, 0
        .WriteProperty "Max", m_Max, 10
        .WriteProperty "Value", m_Value, 0
        .WriteProperty "LargeChange", m_LargeChange, 1
        .WriteProperty "SmallChange", m_SmallChange, 1
        .WriteProperty "SignificantDigits", m_SignificantDigits, 0
        .WriteProperty "OrientationHorizontal", m_OrientationHorizontal, False
        .WriteProperty "VisualStyle", m_VisualStyle, SBVS_Standard
    End With
    
End Sub

'At run-time, painting is handled by PD's pdWindowPainter class.  In the IDE, however, we must rely on VB's internal paint event.
Private Sub UserControl_Paint()
    ucSupport.RequestIDERepaint UserControl.hDC
End Sub

Private Sub UserControl_Resize()
    If Not PDMain.IsProgramRunning() Then ucSupport.RequestRepaint True
End Sub

'When the control value is INCREASED, this function is called
Private Sub MoveValueUp(Optional ByVal useLargeChange As Boolean = False)
    If useLargeChange Then
        Value = m_Value + m_LargeChange
    Else
        Value = m_Value + (m_SmallChange / (10# ^ m_SignificantDigits))
    End If
End Sub

'When the control value is DECREASED, this function is called
Private Sub MoveValueDown(Optional ByVal useLargeChange As Boolean = False)
    If useLargeChange Then
        Value = m_Value - m_LargeChange
    Else
        Value = m_Value - (m_SmallChange / (10# ^ m_SignificantDigits))
    End If
End Sub

'Any changes to size (or control orientation) must call this function to recalculate the positions of all button and
' slider regions.
Private Sub UpdateControlLayout()
    
    'Retrieve DPI-aware control dimensions from the support class
    Dim bWidth As Long, bHeight As Long
    bWidth = ucSupport.GetBackBufferWidth
    bHeight = ucSupport.GetBackBufferHeight
    
    'We now need to figure out the position of the up and down buttons.  Their position (obviously) changes based on the
    ' scroll bar's orientation.  Also note that at present, PD makes no special allotments for tiny scrollbars.  They will
    ' not look or behave correctly.
    If m_OrientationHorizontal Then
        
        'In horizontal orientation, the buttons are calculated as squares aligned with the left and right edges of the control.
        With upLeftRect
            .Left = 0
            .Top = 0
            .Bottom = bHeight - 1
            .Right = bHeight - 1
        End With
        
        With downRightRect
            .Left = (bWidth - 1) - upLeftRect.Bottom
            .Top = 0
            .Right = bWidth - 1
            .Bottom = bHeight - 1
        End With
        
    Else
        
        'In vertical orientation, the buttons are calculated as squares aligned with the left and right edges of the control.
        With upLeftRect
            .Left = 0
            .Top = 0
            .Bottom = bWidth - 1
            .Right = bWidth - 1
        End With
        
        With downRightRect
            .Left = 0
            .Right = bWidth - 1
            .Bottom = bHeight - 1
            .Top = .Bottom - (bWidth - 1)
        End With
    
    End If
    
    'With the button rects calculated, use the difference between them to calculate a track rect.
    If m_OrientationHorizontal Then
        With trackRect
            .Left = upLeftRect.Right + 1
            .Top = 0
            .Right = downRightRect.Left - 1
            .Bottom = bHeight
        End With
    Else
        With trackRect
            .Left = 0
            .Top = upLeftRect.Bottom + 1
            .Right = bWidth
            .Bottom = downRightRect.Top - 1
        End With
    End If
    
    'Figure out the size of the "thumb" slider.  That function will automatically place a call to determineThumbRect.
    DetermineThumbSize
    
    'With all metrics successfully measured, we can now recreate the back buffer
    If ucSupport.AmIVisible Then RedrawBackBuffer
            
End Sub

'Use this function to completely redraw the back buffer from scratch.  Note that this is computationally expensive compared to just flipping the
' existing buffer to the screen, so only redraw the backbuffer if the control state has somehow changed.
Private Sub RedrawBackBuffer(Optional ByVal redrawImmediately As Boolean = False)
    
    Dim enabledState As Boolean
    enabledState = Me.Enabled
    
    'Request the back buffer DC, and ask the support module to erase any existing rendering for us.
    Dim bufferDC As Long, bWidth As Long, bHeight As Long
    bufferDC = ucSupport.GetBackBufferDC(True, m_Colors.RetrieveColor(PDS_Track, enabledState))
    If (bufferDC = 0) Then Exit Sub
    bWidth = ucSupport.GetBackBufferWidth
    bHeight = ucSupport.GetBackBufferHeight
    
    'Paint all backgrounds and borders first
    If PDMain.IsProgramRunning() Then
    
        'Next, initialize a whole bunch of color values
        Dim thumbBorderColor As Long, thumbFillColor As Long
        Dim upButtonBorderColor As Long, downButtonBorderColor As Long
        Dim upButtonFillColor As Long, downButtonFillColor As Long
        Dim upButtonArrowColor As Long, downButtonArrowColor As Long
        
        thumbBorderColor = m_Colors.RetrieveColor(PDS_ThumbBorder, enabledState, m_MouseDownThumb, m_MouseOverThumb Or ucSupport.DoIHaveFocus)
        thumbFillColor = m_Colors.RetrieveColor(PDS_ThumbFill, enabledState, m_MouseDownThumb, m_MouseOverThumb Or ucSupport.DoIHaveFocus)
        upButtonBorderColor = m_Colors.RetrieveColor(PDS_ButtonBorder, enabledState, m_MouseDownUpButton, m_MouseOverUpButton)
        upButtonFillColor = m_Colors.RetrieveColor(PDS_ButtonFill, enabledState, m_MouseDownUpButton, m_MouseOverUpButton)
        upButtonArrowColor = m_Colors.RetrieveColor(PDS_ButtonArrow, enabledState, m_MouseDownUpButton, m_MouseOverUpButton)
        downButtonBorderColor = m_Colors.RetrieveColor(PDS_ButtonBorder, enabledState, m_MouseDownDownButton, m_MouseOverDownButton)
        downButtonFillColor = m_Colors.RetrieveColor(PDS_ButtonFill, enabledState, m_MouseDownDownButton, m_MouseOverDownButton)
        downButtonArrowColor = m_Colors.RetrieveColor(PDS_ButtonArrow, enabledState, m_MouseDownDownButton, m_MouseOverDownButton)
        
        'With colors decided (finally!), we can actually draw the damn thing.
        
        'pd2D is used for rendering
        Dim cSurface As pd2DSurface, cPen As pd2DPen, cBrush As pd2DBrush
        Drawing2D.QuickCreateSurfaceFromDC cSurface, bufferDC, False
        Drawing2D.QuickCreateSolidBrush cBrush
        Drawing2D.QuickCreateSolidPen cPen, 1!
        
        'The up and down buttons are rendered using integer rendering (because we don't require or want
        ' subpixel positioning - these need to be pixel-aligned).
        
        'Up button
        cBrush.SetBrushColor upButtonFillColor
        PD2D.FillRectangleI_FromRectL cSurface, cBrush, upLeftRect
        cPen.SetPenColor upButtonBorderColor
        PD2D.DrawRectangleI_FromRectL cSurface, cPen, upLeftRect
        
        'Down button
        cBrush.SetBrushColor downButtonFillColor
        PD2D.FillRectangleI_FromRectL cSurface, cBrush, downRightRect
        cPen.SetPenColor downButtonBorderColor
        PD2D.DrawRectangleI_FromRectL cSurface, cPen, downRightRect
        
        'Unlike the up/down buttons, we want the thumb to be antialiased and to use subpixel positioning.
        cSurface.SetSurfaceAntialiasing P2_AA_HighQuality
        If (m_ThumbSize > 0) Then
            cBrush.SetBrushColor thumbFillColor
            PD2D.FillRectangleF_FromRectF cSurface, cBrush, thumbRect
            cPen.SetPenColor thumbBorderColor
            PD2D.DrawRectangleF_FromRectF cSurface, cPen, thumbRect
        End If
        
        'Finally, paint the arrows themselves.  (Note that antialiasing remains on.)
        Dim buttonPt1 As PointFloat, buttonPt2 As PointFloat, buttonPt3 As PointFloat
                    
        'Start with the up/left arrow
        If m_OrientationHorizontal Then
            buttonPt1.x = (upLeftRect.Right - upLeftRect.Left) / 2 + FixDPIFloat(2)
            buttonPt1.y = upLeftRect.Top + FixDPIFloat(5)
            
            buttonPt3.x = buttonPt1.x
            buttonPt3.y = upLeftRect.Bottom - FixDPIFloat(5)
            
            buttonPt2.x = buttonPt1.x - FixDPIFloat(3)
            buttonPt2.y = buttonPt1.y + (buttonPt3.y - buttonPt1.y) / 2
        Else
            buttonPt1.x = upLeftRect.Left + FixDPIFloat(5)
            buttonPt1.y = (upLeftRect.Bottom - upLeftRect.Top) / 2 + FixDPIFloat(2)
            
            buttonPt3.x = upLeftRect.Right - FixDPIFloat(5)
            buttonPt3.y = buttonPt1.y
            
            buttonPt2.x = buttonPt1.x + (buttonPt3.x - buttonPt1.x) / 2
            buttonPt2.y = buttonPt1.y - FixDPIFloat(3)
        End If
        
        cPen.SetPenColor upButtonArrowColor
        cPen.SetPenWidth 2!
        cPen.SetPenLineCap P2_LC_Round
        PD2D.DrawLineF cSurface, cPen, buttonPt1.x, buttonPt1.y, buttonPt2.x, buttonPt2.y
        PD2D.DrawLineF cSurface, cPen, buttonPt2.x, buttonPt2.y, buttonPt3.x, buttonPt3.y
                    
        'Next, the down/right-pointing arrow
        If m_OrientationHorizontal Then
            buttonPt1.x = downRightRect.Left + (downRightRect.Right - downRightRect.Left) / 2 - FixDPIFloat(1)
            buttonPt1.y = downRightRect.Top + FixDPIFloat(5)
            
            buttonPt3.x = buttonPt1.x
            buttonPt3.y = downRightRect.Bottom - FixDPIFloat(5)
            
            buttonPt2.x = buttonPt1.x + FixDPIFloat(3)
            buttonPt2.y = buttonPt1.y + (buttonPt3.y - buttonPt1.y) / 2
        Else
            buttonPt1.x = downRightRect.Left + FixDPIFloat(5)
            buttonPt1.y = downRightRect.Top + (downRightRect.Bottom - downRightRect.Top) / 2 - FixDPIFloat(1)
            
            buttonPt3.x = downRightRect.Right - FixDPIFloat(5)
            buttonPt3.y = buttonPt1.y
            
            buttonPt2.x = buttonPt1.x + (buttonPt3.x - buttonPt1.x) / 2
            buttonPt2.y = buttonPt1.y + FixDPIFloat(3)
        End If
        
        cPen.SetPenColor downButtonArrowColor
        PD2D.DrawLineF cSurface, cPen, buttonPt1.x, buttonPt1.y, buttonPt2.x, buttonPt2.y
        PD2D.DrawLineF cSurface, cPen, buttonPt2.x, buttonPt2.y, buttonPt3.x, buttonPt3.y
        
        Set cBrush = Nothing: Set cPen = Nothing: Set cSurface = Nothing
        
    End If
    
    'Paint the final result to the screen, as relevant
    ucSupport.RequestRepaint redrawImmediately
    
End Sub

'The thumb size is contingent on multiple factors: the size and positioning of the up/down buttons, the available space
' between them, and the range between the control's max/min values.  Call this function to determine a size (BUT NOT A
' POSITION) for the thumb.
Private Sub DetermineThumbSize()
    
    If (Not PDMain.IsProgramRunning()) Then Exit Sub
    
    'Start by determining the maximum available size for the thumb
    Dim maxThumbSize As Single
    
    If m_OrientationHorizontal Then
        maxThumbSize = trackRect.Right - trackRect.Left
    Else
        maxThumbSize = trackRect.Bottom - trackRect.Top
    End If
    
    'If the max size is less than zero, force it to zero and exit
    If (maxThumbSize <= 0) Then
        m_ThumbSize = 0
        
    'If the max size is larger than zero, figure out how many discrete, pixel-sized increments there are between
    ' the up/down (or left/right) buttons.
    Else
    
        Dim totalIncrements As Single
        totalIncrements = Abs(m_Max - m_Min) + 1!
        
        If (totalIncrements <> 0!) Then
        
            m_ThumbSize = maxThumbSize / totalIncrements
            
            'Unlike Windows, we enforce a minimum Thumb size of twice the button size.  This makes the scroll bar a bit
            ' easier to work with, especially on images where the ranges of the scrollbars tend to be *enormous*.
            If m_OrientationHorizontal Then
                If (m_ThumbSize < trackRect.Bottom * 2) Then m_ThumbSize = trackRect.Bottom * 2
            Else
                If (m_ThumbSize < trackRect.Right * 2) Then m_ThumbSize = trackRect.Right * 2
            End If
            
            'Also, don't let the size exceed the trackbar area
            If m_OrientationHorizontal Then
                If (m_ThumbSize > trackRect.Right - trackRect.Left) Then m_ThumbSize = trackRect.Right - trackRect.Left
            Else
                If (m_ThumbSize > trackRect.Bottom - trackRect.Top) Then m_ThumbSize = trackRect.Bottom - trackRect.Top
            End If
            
        Else
            m_ThumbSize = 0
        End If
    
    End If
    
    'After determining a thumb size, we must always recalculate the current position.
    DetermineThumbRect
    
End Sub

'Given the control's value, and the (already) determined thumb size, determine its position.
Private Sub DetermineThumbRect()
    
    'Some coordinates are always the same, regardless of position.
    If m_OrientationHorizontal Then
        thumbRect.Top = 0!
        thumbRect.Height = (trackRect.Bottom - trackRect.Top) - 1
    Else
        thumbRect.Left = 0!
        thumbRect.Width = (trackRect.Right - trackRect.Left) - 1
    End If
    
    'Next, let's calculate a few special circumstances: max and min values, specifically.
    If (m_Value <= m_Min) Then
        
        If m_OrientationHorizontal Then
            thumbRect.Left = trackRect.Left
        Else
            thumbRect.Top = trackRect.Top
        End If
    
    ElseIf (m_Value >= m_Max) Then
        
        If m_OrientationHorizontal Then
            thumbRect.Left = trackRect.Right - m_ThumbSize
        Else
            thumbRect.Top = trackRect.Bottom - m_ThumbSize
        End If
    
    'For any other value, we must calculate the position dynamically
    Else
        
        'Figure out how many pixels we have to work with, and note that we must subtract the size of the thumb from
        ' the available track space.
        Dim availablePixels As Single
        If m_OrientationHorizontal Then
            availablePixels = ((trackRect.Right - trackRect.Left) - m_ThumbSize) - 1
        Else
            availablePixels = ((trackRect.Bottom - trackRect.Top) - m_ThumbSize) - 1
        End If
        
        'Figure out the ratio between the current value and the max/min range
        Dim curPositionRatio As Double
        
        If (m_Max <> m_Min) Then
            curPositionRatio = (m_Value - m_Min) / (m_Max - m_Min)
        Else
            curPositionRatio = 0#
        End If
        
        If m_OrientationHorizontal Then
            thumbRect.Left = trackRect.Left + curPositionRatio * availablePixels
        Else
            thumbRect.Top = trackRect.Top + curPositionRatio * availablePixels
        End If
        
    End If
    
    'With the position calculated, we can plug in the size parameter without any special knowledge
    If m_OrientationHorizontal Then
        thumbRect.Width = m_ThumbSize
    Else
        thumbRect.Height = m_ThumbSize
    End If
    
End Sub

'Given an (x, y) coordinate pair, return the "value" corresponding to that position on the scroll bar track.
Private Function GetValueFromMouseCoords(ByVal x As Single, ByVal y As Single, Optional ByVal padToMaxMinRange As Boolean = False) As Double
    
    'Obviously, value calculations differ depending on the scroll bar's orientation.  Start by figuring out the range afforded
    ' by the track bar portion of the scrollbar.
    Dim availablePixels As Double
    If m_OrientationHorizontal Then
        availablePixels = ((trackRect.Right - trackRect.Left) - m_ThumbSize)
    Else
        availablePixels = ((trackRect.Bottom - trackRect.Top) - m_ThumbSize)
    End If
    
    'Next, figure out where the relevant coordinate (x or y, depending on orientation) lies as a fraction of the total width
    Dim posRatio As Double
    If (availablePixels <> 0#) Then
        If m_OrientationHorizontal Then
            posRatio = (x - trackRect.Left) / availablePixels
        Else
            posRatio = (y - trackRect.Top) / availablePixels
        End If
    End If
    
    'Convert that to a matching position on the control's min/max scale
    If (m_Min = m_Max) Then
        GetValueFromMouseCoords = m_Max
    Else
    
        GetValueFromMouseCoords = m_Min + posRatio * (m_Max - m_Min)
        
        'Clamp output to min/max ranges, as a convenience to the caller
        If padToMaxMinRange Then
            If (GetValueFromMouseCoords < m_Min) Then
                GetValueFromMouseCoords = m_Min
            ElseIf (GetValueFromMouseCoords > m_Max) Then
                GetValueFromMouseCoords = m_Max
            End If
        End If
        
    End If
    
    'Just like the .Value property, for integer-only scroll bars, clamp values to their integer range
    If (m_SignificantDigits = 0) Then GetValueFromMouseCoords = Int(GetValueFromMouseCoords)
    
End Function

'When either...
' 1) the scroll bar orientation changes, or...
' 2) the active translation changes...
' ...the popup menu text needs to be updated to match
Private Sub CreatePopupMenu()
    
    'The text of the scroll bar context menu changes depending on orientation.  We match the verbiage and layout
    ' of the default Windows context menu.
    If PDMain.IsProgramRunning() And (Not g_Language Is Nothing) Then
        
        Set m_PopupMenu = New pdPopupMenu
        m_PopupMenu.AddMenuItem g_Language.TranslateMessage("Scroll here"), "scroll-here"
        m_PopupMenu.AddMenuItem "-"
        
        If m_OrientationHorizontal Then
            m_PopupMenu.AddMenuItem g_Language.TranslateMessage("Left edge"), "top"
            m_PopupMenu.AddMenuItem g_Language.TranslateMessage("Right edge"), "bottom"
        Else
            m_PopupMenu.AddMenuItem g_Language.TranslateMessage("Top"), "top"
            m_PopupMenu.AddMenuItem g_Language.TranslateMessage("Bottom"), "bottom"
        End If
        
        m_PopupMenu.AddMenuItem "-"
        
        If m_OrientationHorizontal Then
            m_PopupMenu.AddMenuItem g_Language.TranslateMessage("Page left"), "page-up"
            m_PopupMenu.AddMenuItem g_Language.TranslateMessage("Page right"), "page-down"
        Else
            m_PopupMenu.AddMenuItem g_Language.TranslateMessage("Page up"), "page-up"
            m_PopupMenu.AddMenuItem g_Language.TranslateMessage("Page down"), "page-down"
        End If
        
        m_PopupMenu.AddMenuItem "-"
        
        If m_OrientationHorizontal Then
            m_PopupMenu.AddMenuItem g_Language.TranslateMessage("Scroll left"), "scroll-up"
            m_PopupMenu.AddMenuItem g_Language.TranslateMessage("Scroll right"), "scroll-down"
        Else
            m_PopupMenu.AddMenuItem g_Language.TranslateMessage("Scroll up"), "scroll-up"
            m_PopupMenu.AddMenuItem g_Language.TranslateMessage("Scroll down"), "scroll-down"
        End If
        
    End If
    
End Sub

'Before the control is rendered, we need to retrieve all painting colors from PD's primary theming class.  Note that this
' step must also be called if/when PD's visual theme settings change.
Private Sub UpdateColorList()
    
    'The color list for this control varies based on the control's "m_VisualStyle" setting.  (Canvas scrollbars are rendered
    ' differently from non-canvas scrollbars.)
    Dim colorTag As String
    If (m_VisualStyle = SBVS_Standard) Then colorTag = vbNullString Else colorTag = "CanvasMode"
    
    With m_Colors
        .LoadThemeColor PDS_Track, "Track" & colorTag, IDE_WHITE
        .LoadThemeColor PDS_ThumbBorder, "ThumbBorder" & colorTag, IDE_BLACK
        .LoadThemeColor PDS_ThumbFill, "ThumbFill" & colorTag, IDE_GRAY
        .LoadThemeColor PDS_ButtonBorder, "ButtonBorder" & colorTag, IDE_BLACK
        .LoadThemeColor PDS_ButtonFill, "ButtonFill" & colorTag, IDE_WHITE
        .LoadThemeColor PDS_ButtonArrow, "ButtonArrow" & colorTag, IDE_GRAY
    End With
    
End Sub

'External functions can call this to request a redraw.  This is helpful for live-updating theme settings, as in the Preferences dialog.
Public Sub UpdateAgainstCurrentTheme(Optional ByVal hostFormhWnd As Long = 0)
    If ucSupport.ThemeUpdateRequired Then
        UpdateColorList
        CreatePopupMenu
        If PDMain.IsProgramRunning() Then NavKey.NotifyControlLoad Me, hostFormhWnd
        If PDMain.IsProgramRunning() Then ucSupport.UpdateAgainstThemeAndLanguage
    End If
End Sub

'By design, PD prefers to not use design-time tooltips.  Apply tooltips at run-time, using this function.
' (IMPORTANT NOTE: translations are handled automatically.  Always pass the original English text!)
Public Sub AssignTooltip(ByRef newTooltip As String, Optional ByRef newTooltipTitle As String = vbNullString, Optional ByVal raiseTipsImmediately As Boolean = False)
    ucSupport.AssignTooltip UserControl.ContainerHwnd, newTooltip, newTooltipTitle, raiseTipsImmediately
End Sub
