VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdWindowManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Window Manager class
'Copyright ©2012-2013 by Tanner Helland
'Created: 02/October/13
'Last updated: 02/October/13
'Last update: initial build
'
'PhotoDemon first began as a single-pane, single-image editor.  About a decade ago, I rewrote it as an MDI project, and in
' 2013, it became time to rewrite its window method once again - this time, to a true multi-window project, separate from
' MDI.  (MDI has a lot of interaction problems, not to mention a host of known bugs that will never be fixed because
' Microsoft is - rightfully - moving away from the MDI model.)
'
'One unfortunate result of this change is that PhotoDemon must now be responsible for much of its own window handling.  In
' particular, child windows must move with the parent window, and all document windows must be managed intelligently.
'
'Thus this class.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Subclassed messages related to window movement
Private Const WM_SYSCOMMAND As Long = &H112
Private Const WM_MOUSEMOVE As Long = &H200
Private Const WM_MOVING As Long = &H216
Private Const WM_MOUSELEAVE As Long = &H2A3

'System message that show up as additional params when subclassing WM_SYSCOMMAND
Private Const SC_MAXIMIZE As Long = &HF030&
Private Const SC_RESTORE As Long = &HF120&

'Constant for tracking MouseLeave events
Private Const TME_LEAVE = &H2&

'Constants for setting window transparency
Private Const LWA_ALPHA As Long = &H2&
Private Const WS_EX_LAYERED As Long = &H80000
Private Const GWL_EXSTYLE As Long = (-20)

'Most API calls handle window position and movement in terms of a rect-type variable
Private Type winRect
    x1 As Long
    y1 As Long
    x2 As Long
    y2 As Long
End Type

Private Type toolbarData
    offsetX As Long
    offsetY As Long
    hasBeenMadeTransparent As Boolean
End Type

Private Type tMouseEvent
    SizeOfMe        As Long
    Flags           As Long
    hndTrack        As Long
    HoverTime       As Long
End Type

'API calls for window positioning and tracking
Private Declare Function ScreenToClient Lib "user32" (ByVal hndWindow As Long, ByRef scrPoint As POINTAPI) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hndWindow As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hndWindow As Long, ByRef lpRect As winRect) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hndWindow As Long, ByRef lpRect As winRect) As Long
Private Declare Function TrackMouseEvent Lib "user32" (ByRef mEventsWanted As tMouseEvent) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
      
'API calls for changing toolbar window opacity
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal targetHwnd As Long, ByVal TransColor As Long, ByVal newAlpha As Byte, ByVal fFlags As Long) As Boolean
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal targetHwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal targetHwnd As Long, ByVal nIndex As Long) As Long


'Toolbar tracking
Private toolbarHwnds() As Long
Private toolbarInfo() As toolbarData
Private numOfToolbars As Long

'Primary window tracking
Private mainWindowHwnd As Long

'We must subclass various window types in order to manage them properly
Private cSubclass As cSelfSubHookCallback

'This class provides some helper functions to intelligently position child windows, typically when the program is run for the
' first time and custom positions have not yet been established.  Constants are used to specify where the window should be placed.
Public Sub requestIdealPosition(ByVal childHwnd As Long, Optional ByVal desiredPosition As Long = 0)

    'Before doing anything else, get the current size and location of child and parent windows
    Dim childRect As winRect, parentRect As winRect, parentClientRect As winRect
    GetWindowRect childHwnd, childRect
    GetWindowRect mainWindowHwnd, parentRect
    GetClientRect mainWindowHwnd, parentClientRect

    'By comparing the parent window's client area and actual area, we can determine how much is "chrome", e.g. window borders, etc.
    Dim parentChromeWidth As Long, parentChromeHeight As Long
    parentChromeWidth = (parentRect.x2 - parentRect.x1) - parentClientRect.x2
    parentChromeHeight = (parentRect.y2 - parentRect.y1) - parentClientRect.y2
    
    'Right now, we only support two "desired positions": top-left and top-right corners.  I may add more after additional testing.
    Select Case desiredPosition
    
        'Top-left
        Case 0
            MoveWindow childHwnd, parentRect.x1 + parentChromeWidth \ 2, parentRect.y1 + parentChromeHeight, childRect.x2 - childRect.x1, childRect.y2 - childRect.y1, 1
            
        'Top-right
        Case 1
            MoveWindow childHwnd, parentRect.x2 - parentChromeWidth \ 2 - (childRect.x2 - childRect.x1), parentRect.y1 + parentChromeHeight, childRect.x2 - childRect.x1, childRect.y2 - childRect.y1, 1
    
    End Select

End Sub

'The primary PhotoDemon form must register its hWnd, so we can track its movement and move any children windows accordingly.
Public Sub registerParentHwnd(ByVal parentHwnd As Long)
    mainWindowHwnd = parentHwnd
    cSubclass.ssc_Subclass parentHwnd, , , Me
    cSubclass.ssc_AddMsg parentHwnd, MSG_AFTER, WM_MOVING, WM_SYSCOMMAND, WM_MOUSEMOVE
End Sub

'Each toolbar window must be registered through this class, using its hWnd.  This class will then use that hWnd to do things
' like move child windows proportionally to their parent.
Public Sub registerToolbarHwnd(ByVal childHwnd As Long)

    'Has this toolbar already been registered?
    Dim i As Long
    For i = 0 To numOfToolbars
        If toolbarHwnds(i) = childHwnd Then Exit Sub
    Next i
    
    'Add this toolbar to the collection
    toolbarHwnds(numOfToolbars) = childHwnd
    
    'Calculate the initial offset of the window and store it to the toolbar position array
    Dim childRect As winRect, parentRect As winRect
    GetWindowRect childHwnd, childRect
    GetWindowRect mainWindowHwnd, parentRect
    
    With toolbarInfo(numOfToolbars)
        .offsetX = childRect.x1 - parentRect.x1
        .offsetY = childRect.y1 - parentRect.y1
        .hasBeenMadeTransparent = False
    End With
    
    'Resize the toolbar array for future additions
    numOfToolbars = numOfToolbars + 1
    ReDim Preserve toolbarHwnds(0 To numOfToolbars) As Long
    ReDim Preserve toolbarInfo(0 To numOfToolbars) As toolbarData
    
    'Subclass this toolbar window, and add a user param of 64 + "index in array".  We can use this value to recognize this hWnd as belonging
    ' to a toolbar window, and quickly look it up in our toolbar tracking array.
    ' look up its as belonging to a toolbar window.
    cSubclass.ssc_Subclass childHwnd, 64 + (numOfToolbars - 1), , Me
    cSubclass.ssc_AddMsg childHwnd, MSG_AFTER, WM_MOVING, WM_MOUSEMOVE
    
End Sub

Private Sub Class_Initialize()

    'Reset all toolbar tracking variables
    numOfToolbars = 0
    ReDim toolbarHwnds(0) As Long
    ReDim toolbarInfo(0) As toolbarData
        
    'Prepare the subclasser
    Set cSubclass = New cSelfSubHookCallback

End Sub

Private Sub Class_Terminate()

    'Release all subclassing
    cSubclass.ssc_Terminate

End Sub

'Set a toolbar window's translucency.  We do this to make them semi-transparent when they do not have the mouse over them.
Private Sub SetWindowTranslucency(ByVal srcHwnd As Long, Optional ByVal nTransparency As Long = 255)

    'SetWindowLong and SetLayeredWindowAttributes are API functions, see MSDN for details
    Dim tmpAttributes As Long
    tmpAttributes = GetWindowLong(srcHwnd, GWL_EXSTYLE)
    
    If nTransparency = 255 Then
        SetLayeredWindowAttributes srcHwnd, 0, nTransparency, LWA_ALPHA
        SetWindowLong srcHwnd, GWL_EXSTYLE, tmpAttributes And Not WS_EX_LAYERED
    Else
        SetWindowLong srcHwnd, GWL_EXSTYLE, tmpAttributes Or WS_EX_LAYERED
        SetLayeredWindowAttributes srcHwnd, 0, nTransparency, LWA_ALPHA
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
    
    Dim curWinRect As winRect
    Dim curWinPoint As POINTAPI
    Dim tmpMouseTracker As tMouseEvent
    
    'Before doing anything else, check the user-supplied param.  If it is >= 64, we know this is a child (toolbar) window.
    If lParamUser >= 64 Then
    
        'Retrieve the toolbar's index in our master tracking array
        Dim winIndex As Long
        winIndex = lParamUser - 64
    
        'Check child toolbar window messages
        Select Case uMsg
        
            'The window is currently being moved
            Case WM_MOVING
                
                'lParam contains a pointer to the window's rect location.  Retrieve it now.
                Call CopyMemory(curWinRect, ByVal lParam, Len(curWinRect))
                
                'Calculate this window's rectangle as an offset of its parent.
                Dim parentRect As winRect
                GetWindowRect mainWindowHwnd, parentRect
                
                'Store the new offset values
                toolbarInfo(winIndex).offsetX = curWinRect.x1 - parentRect.x1
                toolbarInfo(winIndex).offsetY = curWinRect.y1 - parentRect.y1
                
            'The mouse is moving inside this window.  We must manually request notification of MouseLeave events before Windows will
            ' send them to this window, so submit a request (if we haven't already).
            Case WM_MOUSEMOVE
                If toolbarInfo(winIndex).hasBeenMadeTransparent Then
                    toolbarInfo(winIndex).hasBeenMadeTransparent = False
                    SetWindowTranslucency lng_hWnd, 255
                End If
                        
        End Select
    
    'The user param is not >= 64, meaning it references the primary window (FormMain).  Move all child windows to match the
    ' parent form's new position.
    Else
    
        Dim childRect As winRect
        Dim i As Long
    
        'Check parent window messages
        Select Case uMsg
        
            'The parent window is currently being moved
            Case WM_MOVING
                
                'lParam contains a pointer to the window's rect location.  Retrieve it now.
                Call CopyMemory(curWinRect, ByVal lParam, Len(curWinRect))
                
                'Move each child window, while retaining the current offset
                For i = 0 To numOfToolbars - 1
                    GetWindowRect toolbarHwnds(i), childRect
                    MoveWindow toolbarHwnds(i), curWinRect.x1 + toolbarInfo(i).offsetX, curWinRect.y1 + toolbarInfo(i).offsetY, childRect.x2 - childRect.x1, childRect.y2 - childRect.y1, 1
                Next i
            
            'The parent window has just been maximized
            Case WM_SYSCOMMAND
            
                If (wParam = SC_MAXIMIZE) Or (wParam = SC_RESTORE) Then
                
                    'Manually retrieve the new window position.
                    GetWindowRect lng_hWnd, curWinRect
                    
                    'Move each child window, while retaining the current offset
                    For i = 0 To numOfToolbars - 1
                        GetWindowRect toolbarHwnds(i), childRect
                        MoveWindow toolbarHwnds(i), curWinRect.x1 + toolbarInfo(i).offsetX, curWinRect.y1 + toolbarInfo(i).offsetY, childRect.x2 - childRect.x1, childRect.y2 - childRect.y1, 1
                    Next i
                    
                End If
                
            'Mouse has been moved.  (NOTE: MDI forms do not receive mouse move events.  ARGH.  I am leaving this so I can fix it
            ' once we have removed all MDI code from the program.)
            Case WM_MOUSEMOVE
                Debug.Print "event received"
                For i = 0 To numOfToolbars - 1
                    If Not toolbarInfo(i).hasBeenMadeTransparent Then
                        toolbarInfo(i).hasBeenMadeTransparent = True
                        SetWindowTranslucency toolbarHwnds(i), 160
                    End If
                Next i
                        
        End Select
    
    End If




' *************************************************************
' C A U T I O N   C A U T I O N   C A U T I O N   C A U T I O N
' -------------------------------------------------------------
' DO NOT ADD ANY OTHER CODE BELOW THE "END SUB" STATEMENT BELOW
'   add this warning banner to the last routine in your class
' *************************************************************
End Sub


