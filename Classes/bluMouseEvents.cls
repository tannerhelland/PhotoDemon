VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "bluMouseEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Note: this file has been modified for use within PhotoDemon.

'This class was originally written by Kroc Camen of camendesign.com.

'Many thanks to Kroc for this excellent mouse interfacing routine, which provides important mouse events not provided
' by VB itself.

Option Explicit
'======================================================================================
'MaSS1VE : The Master System Sonic 1 Visual Editor; Copyright (C) Kroc Camen, 2013
'Licenced under a Creative Commons 3.0 Attribution Licence
'--You may use and modify this code how you see fit as long as you give credit
'======================================================================================
'CLASS :: bluMouseEvents

'Subclass an object to track mouse enter / leave / wheel events

'Status             Ready to use
'Last Updated       08-AUG-13
'Dependencies       bluMagic.cls, WIN32.bas

'*****************************************
'Begin Win32 definitions ported by Tanner
' NOTE: these support functions were copied from the original WIN32 module by Kroc Camen, which was supplied with
'       the original bluMouseEvents class.

'In VB6 True is -1 and False is 0, but in the Win32 API it's 1 for True
Private Enum BOOL
    API_TRUE = 1
    API_FALSE = 0
End Enum

'A point _
 <msdn.microsoft.com/en-us/library/windows/desktop/dd162805%28v=vs.85%29.aspx>
Private Type POINT
   X As Long
   Y As Long
End Type

Private Enum IDC
    'This is our own addition to tell us to not change it one way or another
    vbDefault = 0
    
    IDC_APPSTARTING = 32650&
    IDC_ARROW = 32512&
    IDC_CROSS = 32515&
    IDC_HAND = 32649&
    IDC_HELP = 32651&
    IDC_IBEAM = 32513&
    IDC_ICON = 32641&
    IDC_NO = 32648&
    IDC_SIZEALL = 32646&
    IDC_SIZENESW = 32643&
    IDC_SIZENS = 32645&
    IDC_SIZENWSE = 32642&
    IDC_SIZEWE = 32644&
    IDC_UPARROW = 32516&
    IDC_WAIT = 32514&
End Enum

'Get the dimensions of the whole window, including the border area _
 <msdn.microsoft.com/en-us/library/windows/desktop/ms633519%28v=vs.85%29.aspx>
Private Declare Function user32_GetWindowRect Lib "user32" Alias "GetWindowRect" ( _
    ByVal hndWindow As Long, _
    ByRef IntoRECT As RECT _
) As BOOL

'Is a point in the rectangle? e.g. check if mouse is within a window _
 <msdn.microsoft.com/en-us/library/windows/desktop/dd162882%28v=vs.85%29.aspx>
Private Declare Function user32_PtInRect Lib "user32" Alias "PtInRect" ( _
    ByRef InRect As RECT, _
    ByVal X As Long, _
    ByVal Y As Long _
) As BOOL

'Get/set various system configuration info _
 <msdn.microsoft.com/en-us/library/windows/desktop/ms724947%28v=vs.85%29.aspx>
Private Declare Function user32_SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" ( _
    ByVal Action As SPI, _
    ByVal Param As Long, _
    ByRef ParamAny As Any, _
    ByVal WinIni As Long _
) As BOOL

Private Enum SPI
    'If the high contrast mode is enabled
    'NOTE: This is not the same thing as the high contrast theme -- on Windows XP
     'the user might use a high contrast theme without having high contrast mode on.
     'On Vista and above the high contrast mode is automatically enabled when a high
     'contrast theme is selected: <blogs.msdn.com/b/oldnewthing/archive/2008/12/03/9167477.aspx>
    SPI_GETHIGHCONTRAST = &H42
    
    'Number of "lines" to scroll with the mouse wheel
    SPI_GETWHEELSCROLLLINES = &H68
    'Number of "chars" to scroll with a horizontal mouse wheel
    SPI_GETWHEELSCROLLCHARS = &H6C
    
    'Determines whether the drop shadow effect is enabled.
    SPI_GETDROPSHADOW = &H1024
End Enum

'Load a mouse cursor _
 <msdn.microsoft.com/en-us/library/windows/desktop/ms648391%28v=vs.85%29.aspx>
Private Declare Function user32_LoadCursor Lib "user32" Alias "LoadCursorA" ( _
    ByVal hndInstance As Long, _
    ByVal CursorName As IDC _
) As Long

'Sets the screen cursor _
 <msdn.microsoft.com/en-us/library/windows/desktop/ms648393%28v=vs.85%29.aspx>
Private Declare Function user32_SetCursor Lib "user32" Alias "SetCursor" ( _
    ByVal hndCursor As Long _
) As Long

'End Win32 definitions ported by Tanner
'*****************************************


'/// API DEFS /////////////////////////////////////////////////////////////////////////

'A mouse event _
 <msdn.microsoft.com/en-us/library/windows/desktop/ms644968%28v=vs.85%29.aspx>
Private Type MOUSEHOOKSTRUCT
    Pt As POINT
    hndWindow As Long
    HitTestCode As Long
    ExtraInfo As Long
End Type

'Informs the mouse tracking to use the system default for hover time-out
Private Const HOVER_DEFAULT As Long = &HFFFFFFFF

'Structure to call the mouse tracking API _
 <msdn.microsoft.com/en-us/library/windows/desktop/ms645604%28v=vs.85%29.aspx>
Private Type TRACKMOUSEEVENT
    SizeOfMe        As Long         'Size of this structure in bytes
    Flags           As TME_FLAGS    'Various settings
    hndTrack        As Long         'Which window to track
    HoverTime       As Long         'Hover timeout
End Type

Private Enum TME_FLAGS
    TME_CANCEL = &H80000000
    TME_HOVER = &H1&
    TME_LEAVE = &H2&
    TME_NONCLIENT = &H10&
    TME_QUERY = &H40000000
End Enum

'Tap into the mouse movement _
 <msdn.microsoft.com/en-us/library/windows/desktop/ms646265%28v=vs.85%29.aspx>
Private Declare Function user32_TrackMouseEvent Lib "user32" Alias "TrackMouseEvent" ( _
    ByRef ptrTRACKMOUSEEVENT As TRACKMOUSEEVENT _
) As BOOL

'What virtual keys are pressed, i.e. during `WM_MOUSEHOVER` / `WM_MOUSEMOVE`
Private Enum MK
    MK_LBUTTON = &H1            'Left mouse button
    MK_RBUTTON = &H2            'Right mouse button
    MK_SHIFT = &H4              'SHIFT key
    MK_CONTROL = &H8            'CTRL key
    MK_MBUTTON = &H10           'Middle mouse button
    MK_XBUTTON1 = &H20          'PREV mouse button
    MK_XBUTTON2 = &H40          'NEXT mouse button
End Enum

'The mouse events we're subclassing
Private Enum WM
    WM_SETCURSOR = &H20
    WM_MOUSEMOVE = &H200
    WM_MOUSEWHEEL = &H20A
    WM_MOUSEHWHEEL = &H20E
    WM_MOUSEHOVER = &H2A1
    WM_MOUSELEAVE = &H2A3
End Enum

'/// PROPERTY STORAGE /////////////////////////////////////////////////////////////////

'Which mouse pointer to show when the mouse is in the control
Public MousePointer As Long

'/// PRIVATE VARS /////////////////////////////////////////////////////////////////////

'NOTE: changed by Tanner to use cSelfSubHookCallback, which is already present in PD
'This will handle the subclassing safely
Private Magic As cSelfSubHookCallback

'The control we are tracking
Private hndTarget As Long
'Optionally, the parent form of the target so as to track mouse wheel events without _
 the control having to be focused. We listen into wheel events across the whole form _
 and test to see if the mouse is within the target control
Private hndParentForm As Long

'This strucutre will be used to start the mouse tracking, we have to keep hold of it _
 as the mouse tracking stops after every event fires
Private TME As TRACKMOUSEEVENT
'Whether the mouse tracking is running
Private IsTracking As Boolean
'Whether the mouse has left the window; _
 we use this to ensure that the MouseIn event only fires once
Private IsOut As Boolean

'/// EVENTS ///////////////////////////////////////////////////////////////////////////

'When the mouse leaves the window
Event MouseOut()
'And when it enters. This will fire just once instead of continuously like MouseMove
Event MouseIn()
'When the mouse stays in place for a brief period of time. _
 This is used for tooltips, for example
Event MouseHover( _
    ByVal Button As VBRUN.MouseButtonConstants, _
    ByVal Shift As VBRUN.ShiftConstants, _
    ByVal X As Single, ByVal Y As Single _
)
'Mouse wheel scroll
Event MouseVScroll( _
    ByVal LinesScrolled As Single, _
    ByVal Button As VBRUN.MouseButtonConstants, _
    ByVal Shift As VBRUN.ShiftConstants, _
    ByVal X As Single, ByVal Y As Single _
)
Event MouseHScroll( _
    ByVal CharsScrolled As Single, _
    ByVal Button As VBRUN.MouseButtonConstants, _
    ByVal Shift As VBRUN.ShiftConstants, _
    ByVal X As Single, ByVal Y As Single _
)

'*******************************************
'Begin support functions ported by Tanner
' NOTE: these support functions were copied from the original WIN32 module by Kroc Camen, which was supplied with
'       the original bluMouseEvents class.

'PROPERTY WheelScrollLines : The number of lines to scroll when the mouse wheel rolls _
 ======================================================================================
Public Property Get WheelScrollLines() As Long
    Call user32_SystemParametersInfo(SPI_GETWHEELSCROLLLINES, 0, WheelScrollLines, 0)
    If WheelScrollLines <= 0 Then WheelScrollLines = 3
End Property

'PROPERTY WheelScrollChars : The number of characters to scroll with horizontal wheel _
 ======================================================================================
Public Property Get WheelScrollChars() As Long
    Call user32_SystemParametersInfo(SPI_GETWHEELSCROLLCHARS, 0, WheelScrollChars, 0)
    If WheelScrollChars <= 0 Then WheelScrollChars = 3
End Property

'End support functions ported by Tanner
'*****************************************

'CLASS Initialize _
 ======================================================================================
Private Sub Class_Initialize()
    Let IsTracking = False
    Let IsOut = False
End Sub

'CLASS Terminate _
 ======================================================================================
Private Sub Class_Terminate()
    'Reset state
    Let IsTracking = False
    Let IsOut = False
    'Clean up the subclassing object
    If Not Magic Is Nothing Then
        If hndParentForm = 0 Then
            Call Magic.ssc_DelMsg( _
                hndTarget, MSG_BEFORE, _
                WM_MOUSEWHEEL, WM_MOUSEHWHEEL _
            )
        Else
            Call Magic.ssc_DelMsg( _
                hndParentForm, MSG_BEFORE, _
                WM_MOUSEWHEEL, WM_MOUSEHWHEEL _
            )
            Call Magic.ssc_UnSubclass(hndParentForm)
        End If
        
        Call Magic.ssc_DelMsg( _
            hndTarget, MSG_BEFORE, _
            WM_SETCURSOR, WM_MOUSEMOVE, WM_MOUSEHOVER, WM_MOUSELEAVE _
        )
        Call Magic.ssc_UnSubclass(hndTarget)
        Call Magic.ssc_Terminate
        Set Magic = Nothing
    End If
End Sub

'/// PUBLIC PROCEDURES ////////////////////////////////////////////////////////////////

'Attach _
 ======================================================================================
Public Sub Attach( _
    ByVal Target_hWnd As Long, _
    Optional ByVal ParentForm_hWnd As Long = 0 _
)
    'TODO: Test for MDIForm and fail (can use `FindWindowEx` with "MDIClient")
    If Target_hWnd = 0 Then Exit Sub
    
    'Keep hold of the window handles
    Let hndTarget = Target_hWnd
    Let hndParentForm = ParentForm_hWnd
    
    'Clean up any current tracking
    Call Class_Terminate
    
    'Start subclassing it
    Set Magic = New cSelfSubHookCallback
    If Magic.ssc_Subclass(hndTarget, , , Me) Then
        With TME
            Let .SizeOfMe = LenB(TME)
            Let .Flags = TME_HOVER Or TME_LEAVE
            Let .HoverTime = HOVER_DEFAULT
            Let .hndTrack = hndTarget
        End With
        If user32_TrackMouseEvent(TME) = API_TRUE Then
            Let IsTracking = True
            Let IsOut = False
            'Listen into the mouse events
            Call Magic.ssc_AddMsg( _
                hndTarget, MSG_BEFORE, _
                WM_SETCURSOR, WM_MOUSEMOVE, WM_MOUSEMOVE, WM_MOUSEHOVER, WM_MOUSELEAVE _
            )
        End If
        
        If hndParentForm = 0 Then
            Call Magic.ssc_AddMsg( _
                hndTarget, MSG_BEFORE, _
                WM_MOUSEWHEEL, WM_MOUSEHWHEEL _
            )
        Else
            Call Magic.ssc_Subclass(hndParentForm, , , Me)
            Call Magic.ssc_AddMsg( _
                hndParentForm, MSG_BEFORE, _
                WM_MOUSEWHEEL, WM_MOUSEHWHEEL _
            )
        End If
    End If
End Sub

'/// PRIVATE PROCEDURES ///////////////////////////////////////////////////////////////

'GetMouseStateFromParam : Extract the mouse / key state from a message param _
 ======================================================================================
Private Sub GetMouseStateFromParam( _
    ByVal wParam As Long, ByVal lParam As Long, _
    ByRef Button As VBRUN.ButtonConstants, _
    ByRef Shift As VBRUN.ShiftConstants, _
    ByRef X As Single, ByRef Y As Single _
)
    'TODO: Find out what numbers the X mouse buttons are
    'TODO: Alt mask?
    If (wParam And MK.MK_LBUTTON) <> 0 Then Let Button = Button Or vbLeftButton
    If (wParam And MK.MK_MBUTTON) <> 0 Then Let Button = Button Or vbMiddleButton
    If (wParam And MK.MK_RBUTTON) <> 0 Then Let Button = Button Or vbRightButton
    If (wParam And MK.MK_CONTROL) <> 0 Then Let Shift = Shift Or vbCtrlMask
    If (wParam And MK.MK_SHIFT) <> 0 Then Let Shift = Shift Or vbShiftMask
    
    'The X and Y mouse coords are in the Long value as upper and lower 16-bits
    Let X = lParam And &HFFFF&
    Let Y = lParam \ &H10000
End Sub

'/// SUBCLASS /////////////////////////////////////////////////////////////////////////
'bluMagic helps us tap into the Windows message stream going on in the background _
 so that we can trap mouse / window events and a whole lot more. This works using _
 "function ordinals", therefore this procedure has to be the last one on the page

'SubclassWindowProcedure : THIS MUST BE THE LAST PROCEDURE ON THIS PAGE _
 ======================================================================================
Private Sub SubclassWindowProcedure( _
    ByVal Before As Boolean, _
    ByRef Handled As Boolean, _
    ByRef ReturnValue As Long, _
    ByVal hndWindow As Long, _
    ByVal Message As WM, _
    ByVal wParam As Long, _
    ByVal lParam As Long, _
    ByRef UserParam As Long _
)
    'For the `WM_MOUSEWHEEL` and `WM_MOUSEHOVER` events we'll provide the mouse and _
     key state (mouse button, ctrl/shift key state and mouse X / Y)
    Dim Button As VBRUN.MouseButtonConstants
    Dim Shift As VBRUN.ShiftConstants
    Dim X As Single, Y As Single
    
    '`WM_MOUSEWHEEL`: _
     <msdn.microsoft.com/en-us/library/windows/desktop/ms645617%28v=vs.85%29.aspx>
    '`WM_MOUSEHWHEEL`: _
     <msdn.microsoft.com/en-us/library/windows/desktop/ms645614%28v=vs.85%29.aspx>
    If Message = WM_MOUSEWHEEL _
    Or Message = WM_MOUSEHWHEEL Then '-------------------------------------------------
        'Since mouse wheel events are usually only sent to the window in focus, and _
         we don't want the user to have to click the control before being able to _
         scroll, the parent form is (optionally) subclassed so that we can hear _
         mouse wheel events all the time and we only need check if the mouse is _
         within the control
        
        'Get the location of the control in the screen _
         (including the non-client area, such as the scrollbars)
        Dim WindowRECT As RECT
        Call user32_GetWindowRect(hndTarget, WindowRECT)
        'Determine which mouse button / meta keys are pressed
        Call GetMouseStateFromParam(wParam, lParam, Button, Shift, X, Y)
        If user32_PtInRect(WindowRECT, X, Y) = API_FALSE Then Exit Sub
        
        'The upper bytes of the `wParam` contain the amount the wheel scrolled, _
         which will be some fraction or multiple of 120 _
         <blogs.msdn.com/b/oldnewthing/archive/2013/01/23/10387366.aspx>
        Dim Delta As Long
        Let Delta = wParam \ &H10000
        
        'WARNING: Most code examples out there integer divide the delta (amount _
         scrolled) by 120 to determine the number of "detents" (notches) of the wheel _
         have been scrolled, but newer wheel mice have no notches and scroll smoothly _
         sending multiple `WM_MOUSEWHEEL` messages with smaller delta values, _
         therefore these examples break for absolute delta values that are not _
         multiples of 120. Instead, we will work on partial or multiple "lines" / _
         "chars" being scrolled. The height of a "line" / "char" is up to the _
         controller to convert into px
        
        'Horizontal or vertical scroll? _
         (holding shift converts a vertical scroll to horizontal)
        If Message = WM_MOUSEHWHEEL Or (Shift And vbShiftMask) > 0 Then
            Dim CharsScrolled As Single
            Let CharsScrolled = WheelScrollChars * (Delta / 120)
            'Alert the controller so it can do the scrolling
            RaiseEvent MouseHScroll(CharsScrolled, Button, Shift, X, Y)
        Else
            Dim LinesScrolled As Single
            Let LinesScrolled = WheelScrollLines * (Delta / 120)
            'Alert the controller so it can do the scrolling
            RaiseEvent MouseVScroll(LinesScrolled, Button, Shift, X, Y)
        End If
        
        Let ReturnValue = 0
        Let Handled = True
    
    '`WM_MOUSEHOVER` _
     <msdn.microsoft.com/en-us/library/windows/desktop/ms645613%28v=vs.85%29.aspx>
    ElseIf Message = WM_MOUSEHOVER Then '----------------------------------------------
        Let IsTracking = False
        
        'Determine which mouse button / meta keys are pressed
        Call GetMouseStateFromParam(wParam, lParam, Button, Shift, X, Y)
        RaiseEvent MouseHover(Button, Shift, X, Y)
        
        Let ReturnValue = 0
        Let Handled = True
        
    ElseIf Message = WM_MOUSELEAVE Then '----------------------------------------------
        RaiseEvent MouseOut
        Let IsTracking = False
        Let IsOut = True
        
        Let ReturnValue = 0
        Let Handled = True
    
    ElseIf Message = WM_MOUSEMOVE Then '-----------------------------------------------
        If IsTracking = False Then
            Let IsTracking = True
            Call user32_TrackMouseEvent(TME)
        End If
        If IsOut = True Then
            IsOut = False
            RaiseEvent MouseIn
        End If
        Let ReturnValue = 1
        
        'Tanner's fix: allow VB's internal MouseMove events to still fire
        'Let Handled = True
        Handled = False
        
    ElseIf Message = WM_SETCURSOR Then '-----------------------------------------------
        If IsOut = False And MousePointer > 0 Then
            Call user32_SetCursor( _
                user32_LoadCursor(0, MousePointer) _
            )
            Let ReturnValue = 1
            Let Handled = True
        End If
    
    End If
    
'======================================================================================
'    C A U T I O N   C A U T I O N   C A U T I O N   C A U T I O N   C A U T I O N
'--------------------------------------------------------------------------------------
'           DO NOT ADD ANY OTHER CODE BELOW THE "END SUB" STATEMENT BELOW
'======================================================================================
End Sub
