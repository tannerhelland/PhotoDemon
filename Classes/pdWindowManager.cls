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
'Copyright ©2013-2014 by Tanner Helland
'Created: 02/October/13
'Last updated: 25/March/14
'Last update: add Layers toolbox to the list of windows with specialized handling
'
'PhotoDemon first began as a single-pane, single-image editor.  About a decade ago, I rewrote it as an MDI project, and in
' 2014, it became time to rewrite its window method once again - this time, to a tabbed interface in line with the major
' UI trends and research of the past decade.  (MDI has a lot of interaction problems, not to mention a host of known bugs
' that will never be fixed because Microsoft is - rightfully - moving away from the MDI model.)
'
'One unfortunate result of this change is that PhotoDemon must now be responsible for much of its own window handling.  In
' particular, child windows must move with the parent window, and all document windows must be managed intelligently.
' Top-most behavior must be carefully handled to ensure proper order of the main form, toolboxes, image windows, and various
' modal and non-modal dialogs.
'
'Thus this class.  The main form, toolbar forms, and any special forms (like progress bars) are all registered with this handler
' when they are created. This class manages their placement, styles, opacity, and a host of other minor details.  Additionally,
' other windows (such as tool dialogs) can be registered on-the-fly if they require special behavior.
'
'One thing I am particularly adamant about is properly deactivating top-most windows when the user switches to another
' application.  I believe I have covered every possible contingency, because errant top-most windows are never something I
' want PD to cause.
'
'In order to perform all this detailed window management, this class subclasses a lot of forms and a lot of messages.  That
' said, it should be wholly IDE-safe, but consider yourself warned.  (Also, some window order API calls do not function
' properly in the IDE per http://support.microsoft.com/kb/192254.  Compile for best results.)
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Opacity of toolbar windows when a child window is active
Private Const INACTIVE_TOOLBOX_OPACITY As Byte = 170

'Subclassed messages related to window movement
Private Const WM_SYSCOMMAND As Long = &H112
Private Const WM_SYSCOMMAND_WPARAM_MASK As Long = &HFFF0
Private Const WM_MOUSEMOVE As Long = &H200
Private Const WM_MOVE As Long = &H3
Private Const WM_MOVING As Long = &H216
Private Const WM_MOUSELEAVE As Long = &H2A3
Private Const WM_NCMOUSEMOVE As Long = &HA0
Private Const WM_SIZING As Long = &H214
Private Const WM_SIZE As Long = &H5
Private Const WM_ERASEBKGND As Long = &H14

'Subclassed messages for detecting when PD is losing focus to another app, or regaining it from another app
Private Const WM_ACTIVATE      As Long = &H6
Private Const WM_ACTIVATEAPP   As Long = &H1C
Private Const WA_INACTIVE      As Long = 0
Private Const WA_ACTIVE        As Long = 1
Private Const WA_CLICKACTIVE   As Long = 2

'System message that show up as additional params when subclassing WM_SYSCOMMAND
Private Const SC_MAXIMIZE As Long = &HF030&
Private Const SC_MINIMIZE As Long = &HF020&
Private Const SC_RESTORE As Long = &HF120&
Private Const SC_SIZE As Long = &HF000&
Private Const SC_MOVE As Long = &HF010&

'Constants for changing window style (including transparency)
Private Const LWA_ALPHA As Long = &H2&
Private Const WS_EX_LAYERED As Long = &H80000
Private Const WS_EX_COMPOSITED As Long = &H2000000
Private Const GWL_EXSTYLE As Long = (-20)
Private Const GWL_STYLE As Long = (-16)
Private Const GWL_HWNDPARENT As Long = (-8)

Private Const SWP_ASYNCWINDOWPOS As Long = &H4000
Private Const SWP_FRAMECHANGED As Long = &H20
Private Const SWP_NOACTIVATE As Long = &H10
Private Const SWP_NOMOVE As Long = &H2
Private Const SWP_NOOWNERZORDER As Long = &H200
Private Const SWP_NOREDRAW As Long = &H8
Private Const SWP_NOSENDCHANGING As Long = &H400
Private Const SWP_NOSIZE As Long = &H1
Private Const SWP_HIDEWINDOW As Long = &H80
Private Const SWP_SHOWWINDOW As Long = &H40
Private Const SWP_NOZORDER As Long = &H4
Private Const SWP_DRAWFRAME As Long = &H20
Private Const SWP_NOCOPYBITS As Long = &H100

'Experimental messages related to Win 7 taskbar stuff
Private Const WM_CHILDACTIVATE As Long = &H22

'AnimateWindows decs
Private Const AW_ACTIVATE As Long = &H20000
Private Const AW_BLEND As Long = &H80000
Private Const AW_CENTER As Long = &H10
Private Const AW_HIDE As Long = &H10000
Private Const AW_HOR_POSITIVE As Long = &H1
Private Const AW_HOR_NEGATIVE As Long = &H2
Private Const AW_SLIDE As Long = &H40000
Private Const AW_VER_POSITIVE As Long = &H4
Private Const AW_VER_NEGATIVE As Long = &H8

Private Enum enWindowStyles
    WS_BORDER = &H800000
    WS_CAPTION = &HC00000
    WS_CHILD = &H40000000
    WS_CLIPCHILDREN = &H2000000
    WS_CLIPSIBLINGS = &H4000000
    WS_DISABLED = &H8000000
    WS_DLGFRAME = &H400000
    WS_GROUP = &H20000
    WS_HSCROLL = &H100000
    WS_MAXIMIZE = &H1000000
    WS_MAXIMIZEBOX = &H10000
    WS_MINIMIZE = &H20000000
    WS_MINIMIZEBOX = &H20000
    WS_OVERLAPPED = &H0&
    WS_POPUP = &H80000000
    WS_SYSMENU = &H80000
    WS_TABSTOP = &H10000
    WS_THICKFRAME = &H40000
    WS_VISIBLE = &H10000000
    WS_VSCROLL = &H200000
    WS_EX_ACCEPTFILES = &H10&
    WS_EX_DLGMODALFRAME = &H1&
    WS_EX_NOACTIVATE = &H8000000
    WS_EX_NOPARENTNOTIFY = &H4&
    WS_EX_TOPMOST = &H8&
    WS_EX_TRANSPARENT = &H20&
    WS_EX_TOOLWINDOW = &H80&
    WS_EX_MDICHILD = &H40
    WS_EX_WINDOWEDGE = &H100
    WS_EX_CLIENTEDGE = &H200
    WS_EX_CONTEXTHELP = &H400
    WS_EX_RIGHT = &H1000
    WS_EX_LEFT = &H0
    WS_EX_RTLREADING = &H2000
    WS_EX_LTRREADING = &H0
    WS_EX_LEFTSCROLLBAR = &H4000
    WS_EX_RIGHTSCROLLBAR = &H0
    WS_EX_CONTROLPARENT = &H10000
    WS_EX_STATICEDGE = &H20000
    WS_EX_APPWINDOW = &H40000
    WS_EX_OVERLAPPEDWINDOW = (WS_EX_WINDOWEDGE Or WS_EX_CLIENTEDGE)
    WS_EX_PALETTEWINDOW = (WS_EX_WINDOWEDGE Or WS_EX_TOOLWINDOW Or WS_EX_TOPMOST)
End Enum

#If False Then
    Private Const WS_BORDER = &H800000, WS_CAPTION = &HC00000, WS_CHILD = &H40000000, WS_CLIPCHILDREN = &H2000000, WS_CLIPSIBLINGS = &H4000000, WS_DISABLED = &H8000000, WS_DLGFRAME = &H400000, WS_EX_ACCEPTFILES = &H10&, WS_EX_DLGMODALFRAME = &H1&, WS_EX_NOPARENTNOTIFY = &H4&, WS_EX_TOPMOST = &H8&, WS_EX_TRANSPARENT = &H20&, WS_EX_TOOLWINDOW = &H80&, WS_GROUP = &H20000, WS_HSCROLL = &H100000, WS_MAXIMIZE = &H1000000, WS_MAXIMIZEBOX = &H10000, WS_MINIMIZE = &H20000000, WS_MINIMIZEBOX = &H20000, WS_OVERLAPPED = &H0&, WS_POPUP = &H80000000, WS_SYSMENU = &H80000, WS_TABSTOP = &H10000, WS_THICKFRAME = &H40000, WS_VISIBLE = &H10000000, WS_VSCROLL = &H200000, WS_EX_MDICHILD = &H40, WS_EX_WINDOWEDGE = &H100, WS_EX_CLIENTEDGE = &H200, WS_EX_CONTEXTHELP = &H400, WS_EX_RIGHT = &H1000, WS_EX_LEFT = &H0, WS_EX_RTLREADING = &H2000, WS_EX_LTRREADING = &H0, WS_EX_LEFTSCROLLBAR = &H4000, WS_EX_RIGHTSCROLLBAR = &H0, WS_EX_CONTROLPARENT = &H10000, WS_EX_STATICEDGE = &H20000, WS_EX_APPWINDOW = &H40000
    Private Const WS_EX_OVERLAPPEDWINDOW = (WS_EX_WINDOWEDGE Or WS_EX_CLIENTEDGE), WS_EX_PALETTEWINDOW = (WS_EX_WINDOWEDGE Or WS_EX_TOOLWINDOW Or WS_EX_TOPMOST)
#End If

'Constants for window z-order
Private Const HWND_NOTOPMOST As Long = -2
Private Const HWND_TOPMOST As Long = -1
Private Const HWND_TOP As Long = 0
Private Const HWND_BOTTOM As Long = 1

'Windows can be aligned in a variety of ways, and it's helpful to track each direction independently.
Private Type pdWindowAlignment
    alignLeft As Boolean
    alignTop As Boolean
    alignRight As Boolean
    alignBottom As Boolean
End Type

'Each window that the manager tracks stores some basic metadata in a variable of this type.
Private Type pdWindowData
    alignmentOptions As pdWindowAlignment
    forciblyMinimized As Boolean
    formReference As Form
    hWnd As Long
    hasBeenMadeTransparent As Boolean
    isVisible As Boolean
    lastPosition As winRect
    lastWindowState As Long
    lastWindowBorder As Long
    minimumWidth As Long
    minimumHeight As Long
    offsetX As Long
    offsetY As Long
    optionalToolbarType As pdToolbarType
    originalWidthClient As Long
    originalHeightClient As Long
    originalWidthFull As Long
    originalHeightFull As Long
    previousRenderStyle As renderedBorderStyle
    subclassingReleased As Boolean
    typeOfWindow As pdWindowType
    windowName As String
End Type

'This class tracks a lot of different window types.  Each gets special treatment, so we track their type as follows:
Public Enum pdWindowType
    MAIN_WINDOW = 0
    TOOLBAR_WINDOW = 1
    IMAGE_WINDOW = 2
    IMAGE_TABSTRIP = 3
End Enum

#If False Then
    Private Const MAIN_WINDOW = 0, TOOLBAR_WINDOW = 1, IMAGE_WINDOW = 2, IMAGE_TABSTRIP = 3
#End If

Public Enum pdToolbarType
    NOT_A_TOOLBAR = 0
    FILE_TOOLBOX = 1
    LAYER_TOOLBOX = 2
    TOOLS_TOOLBOX = 3
End Enum

#If False Then
    Private Const NOT_A_TOOLBAR = 0, FILE_TOOLBOX = 1, LAYER_TOOLBOX = 2, TOOLS_TOOLBOX = 3
#End If

Private Enum renderedBorderStyle
    NO_STYLE = 0
    FLOAT_STYLE = 1
    DOCK_STYLE = 2
End Enum

#If False Then
    Private Const NO_STYLE = 0, FLOAT_STYLE = 1, DOCK_STYLE = 2
#End If

'When a window is being resized, it is helpful to know which edge is being resized.  Windows uses the following system, which we mimic:
Private Enum winResizeEdge
    WMSZ_LEFT = 1
    WMSZ_RIGHT = 2
    WMSZ_TOP = 3
    WMSZ_TOPLEFT = 4
    WMSZ_TOPRIGHT = 5
    WMSZ_BOTTOM = 6
    WMSZ_BOTTOMLEFT = 7
    WMSZ_BOTTOMRIGHT = 8
End Enum

#If False Then
    Private Const WMSZ_LEFT = 1, WMSZ_RIGHT = 2, WMSZ_TOP = 3, WMSZ_TOPLEFT = 4, WMSZ_TOPRIGHT = 5, WMSZ_BOTTOM = 6, WMSZ_BOTTOMLEFT = 7, WMSZ_BOTTOMRIGHT = 8
#End If

'These constants can be used as the second parameter of the ShowWindow API function
Private Enum showWindowOptions
    SW_HIDE = 0
    SW_SHOWNORMAL = 1
    SW_SHOWMINIMIZED = 2
    SW_SHOWMAXIMIZED = 3
    SW_SHOWNOACTIVATE = 4
    SW_SHOW = 5
    SW_MINIMIZE = 6
    SW_SHOWMINNOACTIVE = 7
    SW_SHOWNA = 8
    SW_RESTORE = 9
    SW_SHOWDEFAULT = 10
    SW_FORCEMINIMIZE = 11
End Enum

#If False Then
    Private Const SW_HIDE = 0, SW_SHOWNORMAL = 1, SW_SHOWMINIMIZED = 2, SW_SHOWMAXIMIZED = 3, SW_SHOWNOACTIVATE = 4, SW_SHOW = 5, SW_MINIMIZE = 6, SW_SHOWMINNOACTIVE = 7, SW_SHOWNA = 8, SW_RESTORE = 9, SW_SHOWDEFAULT = 10, SW_FORCEMINIMIZE = 11
#End If

'GetSystemMetrics constants
Private Const SM_CYCAPTION As Long = 4
Private Const SM_CYSIZEFRAME As Long = 33

'API calls for window positioning and tracking
Private Declare Function ClientToScreen Lib "user32" (ByVal hndWindow As Long, ByRef lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hndWindow As Long, ByRef scrPoint As POINTAPI) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hndWindow As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hndWindow As Long, ByRef lpRect As winRect) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hndWindow As Long, ByRef lpRect As winRect) As Long
Private Declare Function GetCursorPos Lib "user32" (ByRef lpPoint As POINTAPI) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hndWindow As Long, ByVal nCmdShow As showWindowOptions) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hndWindow As Long) As Long

'API calls for changing window style
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal targetHWnd As Long, ByVal TransColor As Long, ByVal newAlpha As Byte, ByVal fFlags As Long) As Boolean
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal targetHWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal targetHWnd As Long, ByVal nIndex As Long) As Long
Private Declare Sub SetWindowPos Lib "user32" (ByVal targetHWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function BeginDeferWindowPos Lib "user32" (ByVal nNumWindows As Long) As Long
Private Declare Function DeferWindowPos Lib "user32" (ByVal hWinPosInfo As Long, ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function EndDeferWindowPos Lib "user32" (ByVal hWinPosInfo As Long) As Long
Private Declare Function InvalidateRect Lib "user32" (ByVal targetHWnd As Long, ByRef lpRect As Any, ByVal bErase As Long) As Long
Private Declare Function UpdateWindow Lib "user32" (ByVal targetHWnd As Long) As Long
Private Declare Function AdjustWindowRect Lib "user32" (ByRef lpRect As winRect, ByVal dwStyle As Long, ByVal bMenu As Long) As Long
Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal targetHWnd As Long, ByVal lpString As String) As Long
Private Declare Function AnimateWindow Lib "user32" (ByVal targetHWnd As Long, ByVal dwTime As Long, ByVal dwFlags As Long) As Long
'Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function SetFocus Lib "user32" (ByVal targetHWnd As Long) As Long

'Child window tracking
Private pdWindows() As pdWindowData
Private numOfPDWindows As Long

'Primary window tracking; while the main window's data is stored just like the data for any other window, it is helpful to maintain dedicated
' links to its hWnd and position in the tracking array, to prevent us having to constantly search for it.
Private mainWindowHwnd As Long, mainWindowIndex As Long

'The custom image tabstrip is another specialized window that is internally tracked like any other, but because it has unique needs, it is
' helpful to maintain a dedicated link to it.
Private imageTabstripHwnd As Long, imageTabstripIndex As Long

'When PD loses focus, we automatically revert all top-level toolbar windows to normal style.  When PD regains focus, we restore them as top-most.
' Activation/deactivation messages are received multiple times, but we only want to do activation/deactivation stuff once, so our current state
' is tracked by this variable.
Private isPDDeactivated As Boolean

'We must subclass various window types in order to manage them properly
Private cSubclass As cSelfSubHookCallback

'XML handling (used to save/load window locations) is handled through a specialized class
Private xmlEngine As pdXML
Private windowDataPath As String

'Toolbars windows can be floated or docked (as a group - not individually!  e.g. all toolbars can be docked, or all can be floated)
Private toolbarsFloating As Boolean

'When the main form is being moved or resized, it will trigger the movement of all child forms to match.  We set this variable to TRUE
' so that child forms don't waste time and energy handling window move messages that have already been handled.
Private mainFormTriggeredChange As Boolean

'If the program has registered an extra window that it wants made top-most, any hWnds will be stored in this array.
Private extraWindowHwnds() As Long
Private numOfExtraWindows As Long

'When the main form is minimized, a flurry of window messages are sent to all child windows, including instructions to resize to (0, 0).
' To prevent this size from being synched as the "last-known window size", we track minimize mode, and disregard all location and
' size synching while it is active.
Private insideMinimizeMode As Boolean

'Retrieve the current tabstrip alignment
Public Function getImageTabstripAlignment() As AlignConstants

    With pdWindows(imageTabstripIndex).alignmentOptions
    
        If .alignLeft Then getImageTabstripAlignment = vbAlignLeft
        If .alignTop Then getImageTabstripAlignment = vbAlignTop
        If .alignRight Then getImageTabstripAlignment = vbAlignRight
        If .alignBottom Then getImageTabstripAlignment = vbAlignBottom
    
    End With

End Function

'When the user changes the image tabstrip's alignment, this function must be called.  It will make all necessary adjustments
' inside the window manager and automatically reposition the tabstrip.
Public Sub setImageTabstripAlignment(ByVal newAlignment As AlignConstants)

    'Make a copy of the current alignment
    Dim prevAlign As pdWindowAlignment
    prevAlign = pdWindows(imageTabstripIndex).alignmentOptions
    
    'Reset all current alignment values to FALSE
    With pdWindows(imageTabstripIndex)
    
        .alignmentOptions.alignLeft = False
        .alignmentOptions.alignTop = False
        .alignmentOptions.alignRight = False
        .alignmentOptions.alignBottom = False
        
        Select Case newAlignment
        
            Case vbAlignTop
                .alignmentOptions.alignTop = True
            
            Case vbAlignBottom
                .alignmentOptions.alignBottom = True
            
            Case vbAlignLeft
                .alignmentOptions.alignLeft = True
            
            Case vbAlignRight
                .alignmentOptions.alignRight = True
            
        End Select
    
    End With
    
    'If we are changing from a vertical to horizontal alignment (or vice-versa), the window dimensions need to be swapped.
    Dim curRect As winRect
    GetWindowRect imageTabstripHwnd, curRect
    
    Select Case newAlignment
    
        'Check for left/right alignment changing to top/bottom alignment
        Case vbAlignTop, vbAlignBottom
            If prevAlign.alignLeft Or prevAlign.alignRight Then
                SetWindowPos imageTabstripHwnd, 0, curRect.x1, curRect.y1, curRect.x2 - curRect.x1, curRect.x2 - curRect.x1, SWP_NOREDRAW Or SWP_NOACTIVATE Or SWP_NOMOVE Or SWP_NOZORDER Or SWP_NOSENDCHANGING
            Else
                SetWindowPos imageTabstripHwnd, 0, curRect.x1, curRect.y1, curRect.y2 - curRect.y1, curRect.y2 - curRect.y1, SWP_NOREDRAW Or SWP_NOACTIVATE Or SWP_NOMOVE Or SWP_NOZORDER Or SWP_NOSENDCHANGING
            End If
            
        'Check for top/bottom alignment changing to left/right alignment
        Case vbAlignLeft, vbAlignRight
            If prevAlign.alignTop Or prevAlign.alignBottom Then
                SetWindowPos imageTabstripHwnd, 0, curRect.x1, curRect.y1, curRect.y2 - curRect.y1, curRect.y2 - curRect.y1, SWP_NOREDRAW Or SWP_NOACTIVATE Or SWP_NOMOVE Or SWP_NOZORDER Or SWP_NOSENDCHANGING
            Else
                SetWindowPos imageTabstripHwnd, 0, curRect.x1, curRect.y1, curRect.x2 - curRect.x1, curRect.x2 - curRect.x1, SWP_NOREDRAW Or SWP_NOACTIVATE Or SWP_NOMOVE Or SWP_NOZORDER Or SWP_NOSENDCHANGING
            End If
    
    End Select
    
    'Request a redraw of the window, which will handle any additional resizing of the tabstrip
    Dim curMainFormRect As winRect
    GetWindowRect pdWindows(mainWindowIndex).hWnd, curMainFormRect
    
    requestWindowRedraw imageTabstripIndex, curMainFormRect
    
    'Request a redraw of the main window; because it clips children, it needs to be notified of non-standard redraws like this to
    ' avoid leftover artifacts along the tabstrip's borders.
    requestWindowRedraw mainWindowIndex, curMainFormRect
    
    'The available client area for the main form may have changed.  Refresh all relevant image windows.
    Dim i As Long
    For i = 0 To numOfPDWindows - 1
        
        If pdWindows(i).typeOfWindow <> MAIN_WINDOW Then updateWindowStyle i
    
    Next i

End Sub

'Though officially discouraged by Microsoft (see http://msdn.microsoft.com/en-us/library/windows/desktop/ms633591%28v=vs.85%29.aspx), it is
' sometimes useful to dynamically reassign ownership of a window at run-time.  In fact, this is a necessity in VB, because we cannot physically
' assign an owner at creation time, as VB doesn't let us intermix form references and CreateWindow API calls.  (For the record, I am almost
' certain that VB uses this exact same code to assign ownership when it encounters code like FormChild.Show vbModeless, FormOwner.)
' Technically we should also track the previous window owner and restore it upon form exiting, but VB makes that unnecessary due to the way it
' couples owner/child unload behavior.
'
'NOTE: I have left this code here because it may be helpful in the future, but at present, it is not used (due to the removal of
'       floating image windows).
'Public Sub requestNewOwner(ByVal childHwnd As Long, ByVal ownerhWnd As Long)
'    SetWindowLong childHwnd, GWL_HWNDPARENT, ownerhWnd
'End Sub

'If external functions want to know if PD has focus or not, they can use this function.  It will return TRUE if PhotoDemon is the active window.
Public Function doesProgramHaveFocus() As Boolean
    doesProgramHaveFocus = Not isPDDeactivated
End Function

'Sometimes, PhotoDemon needs to display a modal dialog (most tools are modal).  When it does so, this sub should be called so that the
' window engine can make the modal dialog top-most, and deactivate the top-most parameter if the user switches to another application.
Public Sub requestTopmostWindow(ByVal srcWindowHwnd As Long, ByVal ownerHwnd As Long, Optional ByVal windowUnloading As Boolean = False)

    'By default, windows are ADDED to our collection, not removed.
    If Not windowUnloading Then
    
        'Store the window's hWnd.  We need to remember it in case the user switches to another program while this window is still active;
        ' in that case, we will remove the window's top-most status, to prevent the window from obscuring other programs.
        extraWindowHwnds(numOfExtraWindows) = srcWindowHwnd
        numOfExtraWindows = numOfExtraWindows + 1
        ReDim Preserve extraWindowHwnds(0 To numOfExtraWindows) As Long
        
        'Make the window top-most, but do it without moving, sizing, or activating the window.  (Those actions are left up to the caller.)
        SetWindowPos srcWindowHwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOSENDCHANGING 'Or SWP_SHOWWINDOW 'or SWP_NOACTIVATE
        
        'Dynamically set the window's owner to match
        SetWindowLong srcWindowHwnd, GWL_HWNDPARENT, ownerHwnd
        
        'NOW we can display the window.  (Doing this via anything besides VB's in-built .Show method breaks stuff.  I'll need
        ' to investigate further to see if it's possible to mimic the .Show event via API.)
        'ShowWindow srcWindowHwnd, SW_SHOW
    
    'This function should be notified whenever a previously added window has been hidden or unloaded.  We'll remove it from our collection,
    ' and stop activating/deactivating its top-most condition when the program gains or loses focus.
    Else
    
        'Find the hWnd in our tracking array
        Dim i As Long, srcWindowIndex As Long
        srcWindowIndex = -1
        For i = 0 To numOfExtraWindows
            If extraWindowHwnds(i) = srcWindowHwnd Then
                srcWindowIndex = i
                Exit For
            End If
        Next i
        
        'If we found this hWnd, remove it from the collection
        If srcWindowIndex >= 0 Then
        
            'Shift all existing entries down by one
            For i = srcWindowIndex To numOfExtraWindows - 1
                extraWindowHwnds(i) = extraWindowHwnds(i + 1)
            Next i
            
            'Decrement our spare window counter
            numOfExtraWindows = numOfExtraWindows - 1
            
            'Shrink the array by one
            ReDim Preserve extraWindowHwnds(0 To numOfExtraWindows) As Long
            
        End If
    
    End If

End Sub

'Using a "Form.Caption = <string>" statement on a docked child form will cause the window to un-dock.  Thus, all caption requests
' must be passed through the window manager, who will then use the API to change the text.
Public Sub requestWindowCaptionChange(ByRef targetForm As Form, ByVal newCaption As String)
    SetWindowText targetForm.hWnd, newCaption
End Sub

'Toolbox windows can be hidden (but not unloaded!) via this sub.  It's important to note that toolbox windows should *not* be
' unloaded, because many aspects of PD interact with the toolboxes - for example, en/disabling buttons on the file toolbar as
' necessary - and if the toolbox is unloaded, those statements will cause errors.
Public Sub setWindowVisibility(ByVal winHwnd As Long, ByVal newVisibility As Boolean, Optional ByVal suppressImageRedraws As Boolean = False)
    
    'Find this entry in the windows array, and mark its visibility status
    Dim winIndex As Long
    
    Dim i As Long
    For i = 0 To numOfPDWindows - 1
        If pdWindows(i).hWnd = winHwnd Then
        
            winIndex = i
        
            'If this window's visibility already matches what we have stored, ignore the request and exit
            If pdWindows(i).isVisible = newVisibility Then
                Exit Sub
            Else
                pdWindows(i).isVisible = newVisibility
                Exit For
            End If
            
        End If
    Next i
    
    'Use the ShowWindow API to show or hide the window
    If newVisibility Then
        ShowWindow winHwnd, SW_RESTORE
    Else
        ShowWindow winHwnd, SW_HIDE
    End If
    
    'The available client area for the main form may have changed.  Refresh all relevant windows.
    If Not suppressImageRedraws Then
    
        For i = 0 To numOfPDWindows - 1
            If pdWindows(i).typeOfWindow <> MAIN_WINDOW Then updateWindowStyle i
        Next i
    
    End If
     
    'Because the main image canvas area may be modified due to window visibility, redraw it now
    FormMain.refreshAllCanvases
    
End Sub

'If external (or internal) functions move windows around on their own, they can manually request resyncing the window's offset.
Public Sub requestWindowResync(ByVal winIndex As Long)
    
    If insideMinimizeMode Then Exit Sub
    
    'Get the parent's window rect
    Dim parentRect As winRect
    GetWindowRect mainWindowHwnd, parentRect
    
    'Get the child's window rect
    Dim childRect As winRect
    GetWindowRect pdWindows(winIndex).hWnd, childRect
    
    'Store the new offset values
    pdWindows(winIndex).offsetX = childRect.x1 - parentRect.x1
    pdWindows(winIndex).offsetY = childRect.y1 - parentRect.y1
    
    'Also, update the last-known location properly
    updateStoredWindowLocation winIndex
    
End Sub

'Similar to the functions below, this sub can be used to retrieve a rect that contains (in *screen coordinates*) the position of the main
' form's "actual" client rect, by which I mean - the rect accounting for docked toolbars, as necessary.
Friend Sub getActualMainFormClientRect(ByRef modifiedRect As winRect, Optional ByVal ignoreTabstrip As Boolean = False, Optional ByVal useScreenCoordinates As Boolean = True)

    'Start by retrieving the main form's client rect, in screen coordinates as desired.
    Dim clientRect As winRect, toolbarRect As winRect
    
    If useScreenCoordinates Then
        getClientCoordsRelativeToScreen mainWindowHwnd, clientRect
    Else
        GetClientRect mainWindowHwnd, clientRect
    End If
        
    'If toolbars are docked, subtract their width and/or height from the rect, contingent on their visibility and alignment
    If (Not toolbarsFloating) Then
        
        Dim i As Long
        For i = 0 To numOfPDWindows - 1
            If (pdWindows(i).typeOfWindow = TOOLBAR_WINDOW) And (pdWindows(i).isVisible) Then
                
                GetWindowRect pdWindows(i).hWnd, toolbarRect
            
                'Use the toolbar's alignment to determine if its width should be subtracted from the left or right of the client area
                If pdWindows(i).alignmentOptions.alignLeft Then
                    clientRect.x1 = clientRect.x1 + (toolbarRect.x2 - toolbarRect.x1)
                ElseIf pdWindows(i).alignmentOptions.alignRight Then
                    clientRect.x2 = clientRect.x2 - (toolbarRect.x2 - toolbarRect.x1)
                ElseIf pdWindows(i).alignmentOptions.alignTop Then
                    clientRect.y1 = clientRect.y1 + (toolbarRect.y2 - toolbarRect.y1)
                ElseIf pdWindows(i).alignmentOptions.alignBottom Then
                    clientRect.y2 = clientRect.y2 - (toolbarRect.y2 - toolbarRect.y1)
                End If
                
            End If
        Next i
        
    End If
        
    'If the image tabstrip is visible, remove its size as well
    If pdWindows(imageTabstripIndex).isVisible And (Not ignoreTabstrip) Then
    
        With pdWindows(imageTabstripIndex)
        
            If .alignmentOptions.alignLeft Then clientRect.x1 = clientRect.x1 + getWindowWidth(.hWnd)
            If .alignmentOptions.alignTop Then clientRect.y1 = clientRect.y1 + getWindowHeight(.hWnd)
            If .alignmentOptions.alignRight Then clientRect.x2 = clientRect.x2 - getWindowWidth(.hWnd)
            If .alignmentOptions.alignBottom Then clientRect.y2 = clientRect.y2 - getWindowHeight(.hWnd)
        
        End With
        
    End If
    
    'Provide a 1px border around the client area
    'NOTE: this has been removed as of 20 Dec 13, but I'm leaving the code as comments in case it ever becomes prudent
    '      to reinstate the border.
    'clientRect.x1 = clientRect.x1 + 1
    'clientRect.x2 = clientRect.x2 - 1
    'clientRect.y1 = clientRect.y1 + 1
    'clientRect.y2 = clientRect.y2 - 1
    
    modifiedRect = clientRect

End Sub

'External functions (such as those that auto-zoom an image to fit it on-screen) need to know the size of the client area on the main form.
' That client size will change based on the docking status of toolbars, so we must use the window manager to calculate it.
Public Function requestActualMainFormClientWidth() As Long

    Dim clientRect As winRect
    GetClientRect mainWindowHwnd, clientRect
    requestActualMainFormClientWidth = clientRect.x2 - clientRect.x1
    
    'If the image tabstrip is visible and left- or right-aligned, remove its width as well
    If pdWindows(imageTabstripIndex).isVisible Then
        If pdWindows(imageTabstripIndex).alignmentOptions.alignLeft Or pdWindows(imageTabstripIndex).alignmentOptions.alignRight Then
            requestActualMainFormClientWidth = requestActualMainFormClientWidth - getWindowWidth(pdWindows(imageTabstripIndex).hWnd)
        End If
    End If
    
    'When toolbars are docked, we must remove all toolbar widths from the client size.
    If Not toolbarsFloating Then
        
        Dim toolbarRect As winRect
        
        'Remove all vertical toolbar widths from the client rect
        Dim i As Long
        For i = 0 To numOfPDWindows - 1
            If (pdWindows(i).typeOfWindow = TOOLBAR_WINDOW) And (pdWindows(i).isVisible) Then
                If pdWindows(i).alignmentOptions.alignLeft Or pdWindows(i).alignmentOptions.alignRight Then
                    GetWindowRect pdWindows(i).hWnd, toolbarRect
                    requestActualMainFormClientWidth = requestActualMainFormClientWidth - (toolbarRect.x2 - toolbarRect.x1)
                End If
            End If
        Next i
        
    End If

End Function

'At present, the docking status of toolbars doesn't affect the main form's client height, making this calculation quite easy.
Public Function requestActualMainFormClientHeight() As Long

    Dim clientRect As winRect
    GetClientRect mainWindowHwnd, clientRect
    requestActualMainFormClientHeight = clientRect.y2
    
    'If the image tabstrip is visible and top- or bottom-aligned, remove its height as well
    If pdWindows(imageTabstripIndex).isVisible Then
        If pdWindows(imageTabstripIndex).alignmentOptions.alignTop Or pdWindows(imageTabstripIndex).alignmentOptions.alignBottom Then
            requestActualMainFormClientHeight = requestActualMainFormClientHeight - getWindowHeight(pdWindows(imageTabstripIndex).hWnd)
        End If
    End If
    
    'When toolbars are docked, we must remove all horizontal toolbar heights from the client size.
    If Not toolbarsFloating Then
    
        Dim toolbarRect As winRect
        
        'Remove all vertical toolbar widths from the client rect
        Dim i As Long
        For i = 0 To numOfPDWindows - 1
            If (pdWindows(i).typeOfWindow = TOOLBAR_WINDOW) And (pdWindows(i).isVisible) Then
                If pdWindows(i).alignmentOptions.alignTop Or pdWindows(i).alignmentOptions.alignBottom Then
                    GetWindowRect pdWindows(i).hWnd, toolbarRect
                    requestActualMainFormClientHeight = requestActualMainFormClientHeight - (toolbarRect.y2 - toolbarRect.y1)
                End If
            End If
        Next i
        
    End If

End Function

'If a toolbar is docked, it will also be aligned.  We can use alignment data to position it instead of an x/y offset.
Private Sub setWindowPosByAlignment(ByVal winIndex As Long, Optional ByVal suspendRedraw As Boolean = False)

    Dim newPositionRect As winRect
        
    'Retrieve the window's current size
    Dim curWinRect As winRect
    GetClientRect pdWindows(winIndex).hWnd, curWinRect
    
    'Retrieve the client area of the main form
    Dim mainClientRect As winRect
    GetClientRect mainWindowHwnd, mainClientRect
        
    'PhotoDemon toolboxes have a sort of hierarchy in how they are arranged.  In order, the preference for sizing goes:
    ' 1) File toolbox (on left)
    ' 2) Tools toolbox (on bottom)
    ' 3) Layers toolbox (on right)
    Dim i As Long, fileToolboxIndex As Long, toolsToolboxIndex As Long
    For i = 0 To numOfPDWindows - 1
        If pdWindows(i).optionalToolbarType = FILE_TOOLBOX Then fileToolboxIndex = i
        If pdWindows(i).optionalToolbarType = TOOLS_TOOLBOX Then toolsToolboxIndex = i
    Next i
    
    'For the tools toolbox, add the width of the file toolbox to the x1 position
    If (pdWindows(winIndex).optionalToolbarType = TOOLS_TOOLBOX) And (pdWindows(fileToolboxIndex).isVisible) Then
        mainClientRect.x1 = mainClientRect.x1 + getClientWidth(toolbar_File.hWnd)
    End If
    
    'For the layers toolbox, subtract the height of the tools toolbox from the y2 position
    If (pdWindows(winIndex).optionalToolbarType = LAYER_TOOLBOX) And (pdWindows(toolsToolboxIndex).isVisible) Then
        mainClientRect.y2 = mainClientRect.y2 - getClientHeight(toolbar_Tools.hWnd)
    End If
    
    'At present, we assume that a window is either aligned top or bottom, right or left.  (e.g. it must be aligned to a corner)
    ' This makes positioning much simpler.
    With pdWindows(winIndex)
    
        'Horizontal alignment
        If .alignmentOptions.alignLeft Then
            newPositionRect.x1 = mainClientRect.x1
            newPositionRect.y1 = mainClientRect.y1
        End If
        If .alignmentOptions.alignRight Then
            newPositionRect.x1 = mainClientRect.x2 - (curWinRect.x2 - curWinRect.x1)
            newPositionRect.y1 = mainClientRect.y1
        End If
        
        'Vertical alignment
        If .alignmentOptions.alignTop Then
            newPositionRect.x1 = mainClientRect.x1
            newPositionRect.y1 = mainClientRect.y1
        End If
        If .alignmentOptions.alignBottom Then
            newPositionRect.x1 = mainClientRect.x1
            newPositionRect.y1 = mainClientRect.y2 - (curWinRect.y2 - curWinRect.y1)
        End If
    
    End With
    
    'Calculate the new position x2 and y2 to match the window's current width and height
    newPositionRect.x2 = newPositionRect.x1 + (curWinRect.x2 - curWinRect.x1)
    newPositionRect.y2 = newPositionRect.y1 + (curWinRect.y2 - curWinRect.y1)
    
    'Move the window into place
    MoveWindow pdWindows(winIndex).hWnd, newPositionRect.x1, newPositionRect.y1, newPositionRect.x2 - newPositionRect.x1, newPositionRect.y2 - newPositionRect.y1, IIf(suspendRedraw, 0, 1)
    
    'Resync the window's offset
    requestWindowResync winIndex

End Sub

'If toolbars are docked, their height (and/or width, if I ever add horizontal toolbars) needs to be kept in sync with the main form.
' After the main form is resized, this function can be called to resync all toolbar width/heights.
Private Sub fixDockedToolbarSize(ByVal winIndex As Long)

    'If toolbars are not docked, exit.
    If toolbarsFloating Then Exit Sub

    'Resize the window to fit inside the parent form, but not overlap the progress bar at the bottom
    Dim clientWidth As Long, clientHeight As Long
    clientHeight = getClientHeight(mainWindowHwnd)
    clientWidth = getClientWidth(mainWindowHwnd)
    
    'In rare circumstances, it's possible for the user to shrink the main form horizontally to the point that the main menu overflows
    ' into a second row.  To accomodate this case, also recalculate the proper top client coordinate.
    Dim curMainFormClient As winRect
    GetClientRect mainWindowHwnd, curMainFormClient
        
    'PhotoDemon toolboxes have a sort of hierarchy in how they are arranged.  In order, the preference for sizing goes:
    ' 1) File toolbox (on left)
    ' 2) Tools toolbox (on bottom)
    ' 3) Layers toolbox (on right)
    Dim i As Long, fileToolboxIndex As Long, toolsToolboxIndex As Long
    For i = 0 To numOfPDWindows - 1
        If pdWindows(i).optionalToolbarType = FILE_TOOLBOX Then fileToolboxIndex = i
        If pdWindows(i).optionalToolbarType = TOOLS_TOOLBOX Then toolsToolboxIndex = i
    Next i
    
    'For the tools toolbox, add the width of the file toolbox to the x1 position
    If (pdWindows(winIndex).optionalToolbarType = TOOLS_TOOLBOX) And (pdWindows(fileToolboxIndex).isVisible) Then
        clientWidth = clientWidth - getClientWidth(toolbar_File.hWnd)
        curMainFormClient.x1 = curMainFormClient.x1 + getClientWidth(toolbar_File.hWnd)
    End If
    
    'For the layers toolbox, subtract the height of the tools toolbox from the y2 position
    If (pdWindows(winIndex).optionalToolbarType = LAYER_TOOLBOX) And (pdWindows(toolsToolboxIndex).isVisible) Then
        clientHeight = clientHeight - getClientHeight(toolbar_Tools.hWnd)
        curMainFormClient.y2 = curMainFormClient.y2 - getClientHeight(toolbar_Tools.hWnd)
    End If
    
    'Get this toolbar window's current location and size.  (We only care about size, as location will have already been set
    ' in a prior step.)
    Dim childRect As winRect
    GetWindowRect pdWindows(winIndex).hWnd, childRect
        
    'Depending on the toolbar's alignment, we will either change the height or the width
    If pdWindows(winIndex).alignmentOptions.alignLeft Or pdWindows(winIndex).alignmentOptions.alignRight Then
        childRect.y1 = curMainFormClient.y1
        childRect.y2 = curMainFormClient.y1 + clientHeight
    ElseIf pdWindows(winIndex).alignmentOptions.alignBottom Or pdWindows(winIndex).alignmentOptions.alignTop Then
        childRect.x1 = curMainFormClient.x1
        childRect.x2 = curMainFormClient.x1 + clientWidth
    End If
    
    'Reposition the window, but note that we do not want it moved - only resized
    SetWindowPos pdWindows(winIndex).hWnd, getDesiredZOrder(winIndex), 0, 0, childRect.x2 - childRect.x1, childRect.y2 - childRect.y1, SWP_NOMOVE Or SWP_NOACTIVATE Or SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_NOREDRAW
    UpdateWindow pdWindows(winIndex).hWnd
    
    'SOME THOUGHTS:
    ' The code above still works, but it shouldn't - SetWindowPos is not requesting a redraw, so why is the window being redrawn?
    ' I need to track down the source of these frequent redraws.
    '
    ' Assuming other sources can be removed, it may work better to calculate InvalidateRect ourself prior to SetWindowPos, as we know
    ' the window's before and after positions...
    
End Sub

'Given a window, return its horizontal (or vertical) client size, in pixels.
Public Function getClientWidth(ByVal targetHWnd As Long) As Long
    Dim tmpRect As winRect
    GetClientRect targetHWnd, tmpRect
    getClientWidth = tmpRect.x2
End Function

Public Function getClientHeight(ByVal targetHWnd As Long) As Long
    Dim tmpRect As winRect
    GetClientRect targetHWnd, tmpRect
    getClientHeight = tmpRect.y2
End Function

'Given a window, return its horizontal (or vertical) size, in pixels.
Private Function getWindowWidth(ByVal targetHWnd As Long) As Long
    Dim tmpRect As winRect
    GetWindowRect targetHWnd, tmpRect
    getWindowWidth = tmpRect.x2 - tmpRect.x1
End Function

Private Function getWindowHeight(ByVal targetHWnd As Long) As Long
    Dim tmpRect As winRect
    GetWindowRect targetHWnd, tmpRect
    getWindowHeight = tmpRect.y2 - tmpRect.y1
End Function

'Given a window, find the size differential between the window and client areas.
' These two functions *should* return accurate results regardless of OS visual settings.
Public Function getHorizontalChromeSize(ByVal targetHWnd As Long) As Long

    'Ask the OS to report window and client size
    Dim WindowRECT As winRect, clientRect As winRect
    GetWindowRect targetHWnd, WindowRECT
    GetClientRect targetHWnd, clientRect
    
    getHorizontalChromeSize = (WindowRECT.x2 - WindowRECT.x1) - clientRect.x2
    
End Function

Public Function getVerticalChromeSize(ByVal targetHWnd As Long) As Long

    'Ask the OS to report window and client size
    Dim WindowRECT As winRect, clientRect As winRect
    GetWindowRect targetHWnd, WindowRECT
    GetClientRect targetHWnd, clientRect
    
    getVerticalChromeSize = (WindowRECT.y2 - WindowRECT.y1) - clientRect.y2
    
End Function

'Force a full window refresh (invalidate + update)
Private Sub forceFullWindowRefresh(ByVal winIndex As Long)
    InvalidateRect pdWindows(winIndex).hWnd, ByVal 0&, 0
    UpdateWindow pdWindows(winIndex).hWnd
End Sub

'Retrieve the client coordinates of a window, but relative to the *screen* (e.g. with the top-left set to screen coords instead of the default (0,0)
Private Sub getClientCoordsRelativeToScreen(ByVal targetHWnd As Long, ByRef rectToFill As winRect)

    Dim curClientRect As winRect
    GetClientRect targetHWnd, curClientRect
    
    'Client rect is always based at (0,0); to convert it to screen coordinates, we use the ClientToScreen API function.
    Dim clientLeft As POINTAPI, clientRight As POINTAPI
    
    'Retrieve the top-left corner first.  Note that we don't have to fill clientLeft, because the top-left client coordinates are always (0,0)
    ClientToScreen targetHWnd, clientLeft
    
    'Fill the bottom-right coord object, then convert it to screen coords
    clientRight.x = curClientRect.x2
    clientRight.y = curClientRect.y2
    ClientToScreen targetHWnd, clientRight
    
    'Fill the rect object we were passed.
    rectToFill.x1 = clientLeft.x
    rectToFill.y1 = clientLeft.y
    rectToFill.x2 = clientRight.x
    rectToFill.y2 = clientRight.y

End Sub

'When a window is floated or docked, certain appearance changes must be made (such as changing its window border).  Because multiple actions
' trigger such a change, we provide the ability to update a window separately, via this function.
Private Sub updateWindowStyle(ByVal winIndex As Long, Optional ByVal isFirstLoad As Boolean = False)

    Dim mainFormRect As winRect, mainFormClientRect As winRect, childRect As winRect
    Dim newWinStyle As Long
    
    'Because we will access the main form's location and client location under a variety of circumstances, retrieve their values now
    GetWindowRect mainWindowHwnd, mainFormRect
    getClientCoordsRelativeToScreen mainWindowHwnd, mainFormClientRect
    
    Select Case pdWindows(winIndex).typeOfWindow
    
        'Toolbar windows
        Case TOOLBAR_WINDOW
                    
            'FLOATING STYLE
            If toolbarsFloating Then
            
                GetWindowRect pdWindows(winIndex).hWnd, childRect
                
                If pdWindows(winIndex).previousRenderStyle <> FLOAT_STYLE Then
                
                    'Remove the main window parent reference (if present)
                    SetParent pdWindows(winIndex).hWnd, 0
                    
                    'Set the border style to fixed toolbar
                    newWinStyle = WS_BORDER Or WS_CAPTION Or WS_SYSMENU
                    If pdWindows(winIndex).isVisible Then newWinStyle = newWinStyle Or WS_VISIBLE
                    SetWindowLong pdWindows(winIndex).hWnd, GWL_STYLE, newWinStyle
                    SetWindowLong pdWindows(winIndex).hWnd, GWL_EXSTYLE, WS_EX_TOOLWINDOW 'Or WS_EX_LAYERED
                    
                    'Remember the window's current style.  This saves us having to redraw it in the future if its style is acceptable.
                    pdWindows(winIndex).previousRenderStyle = FLOAT_STYLE
                    
                    'Using this new border style, calculate a desired window rectangle while forcing the client size to the window's
                    ' original dimensions.  (Note that this does not matter the *first* time a floating window is rendered, because
                    ' the window will already be in floating-style.)
                    If Not isFirstLoad Then
                        childRect.x2 = childRect.x1 + pdWindows(winIndex).originalWidthClient
                        childRect.y2 = childRect.y1 + pdWindows(winIndex).originalHeightClient
                        AdjustWindowRect childRect, GetWindowLong(pdWindows(winIndex).hWnd, GWL_STYLE), 0
                    End If
                    
                    'Erase any alignment data, as floating windows cannot currently be aligned
                    pdWindows(winIndex).alignmentOptions.alignTop = False
                    pdWindows(winIndex).alignmentOptions.alignBottom = False
                    pdWindows(winIndex).alignmentOptions.alignLeft = False
                    pdWindows(winIndex).alignmentOptions.alignRight = False
                    
                    'Immediately synchronize the toolbar's initial position
                    requestWindowResync winIndex
                
                End If
                
                'Use that new rect to resize the image
                SetWindowPos pdWindows(winIndex).hWnd, getDesiredZOrder(winIndex), childRect.x1, childRect.y1, childRect.x2 - childRect.x1, childRect.y2 - childRect.y1, SWP_FRAMECHANGED Or SWP_DRAWFRAME Or SWP_NOOWNERZORDER
                
                'Force a full repaint (to make sure children controls are also repainted)
                forceFullWindowRefresh winIndex
                
            'DOCKED STYLE
            Else
            
                'Before changing the border style, determine the client area.  We will use this to resize the window after removing its border.
                Dim clientDimensions As winRect
                GetClientRect pdWindows(winIndex).hWnd, clientDimensions
                
                If pdWindows(winIndex).previousRenderStyle <> DOCK_STYLE Then
                    
                    'Remove the window border
                    newWinStyle = WS_CHILD
                    If pdWindows(winIndex).isVisible Then newWinStyle = newWinStyle Or WS_VISIBLE
                    SetWindowLong pdWindows(winIndex).hWnd, GWL_STYLE, newWinStyle
                    SetWindowLong pdWindows(winIndex).hWnd, GWL_EXSTYLE, WS_EX_NOACTIVATE 'Or WS_EX_LAYERED
                    SetWindowPos pdWindows(winIndex).hWnd, getDesiredZOrder(winIndex), 0, 0, pdWindows(winIndex).originalWidthClient, pdWindows(winIndex).originalHeightClient, SWP_NOMOVE Or SWP_FRAMECHANGED Or SWP_NOOWNERZORDER Or SWP_NOCOPYBITS
                    
                    'Specify the main window as the parent
                    SetParent pdWindows(winIndex).hWnd, mainWindowHwnd
                    
                    'Remember the window's current style.  This saves us having to redraw it in the future if its style is acceptable.
                    pdWindows(winIndex).previousRenderStyle = DOCK_STYLE
                    
                End If
                                    
                'Restore full window opacity
                setWindowTranslucency pdWindows(winIndex).hWnd, 255
                
                'Get an updated window rect
                GetWindowRect pdWindows(winIndex).hWnd, childRect
                
                'Because these are now children of the main window (per their "docked" status"), all movement is relative
                ' to the parent's client area - NOT the screen.  Thus we need to retrieve the parent's rect in client coords.
                GetClientRect mainWindowHwnd, mainFormClientRect
                
                'Embed the toolbar in a specific section of the main form
                Select Case pdWindows(winIndex).optionalToolbarType
                
                    'Main toolbar gets embedded in the top-left
                    Case FILE_TOOLBOX
                        pdWindows(winIndex).alignmentOptions.alignLeft = True
                        
                    'Selection toolbar gets embedded at the bottom
                    Case TOOLS_TOOLBOX
                        pdWindows(winIndex).alignmentOptions.alignBottom = True
                        
                    'Layers toolbar gets embedded at the right
                    Case LAYER_TOOLBOX
                        pdWindows(winIndex).alignmentOptions.alignRight = True
                        
                End Select
                
                'Use each toolbox's alignment to set its new position and size
                setWindowPosByAlignment winIndex
                fixDockedToolbarSize winIndex
                    
                'Force a full repaint (to make sure children controls are also repainted)
                forceFullWindowRefresh winIndex
                
            End If
            
        'Image tabstrip
        Case IMAGE_TABSTRIP
        
            'Remove the window border
            newWinStyle = WS_CHILD
            If pdWindows(winIndex).isVisible Then newWinStyle = newWinStyle Or WS_VISIBLE
            SetWindowLong pdWindows(winIndex).hWnd, GWL_STYLE, newWinStyle
            'SetWindowLong pdWindows(winIndex).hWnd, GWL_EXSTYLE, WS_EX_LAYERED
            SetWindowPos pdWindows(winIndex).hWnd, getDesiredZOrder(winIndex), 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_FRAMECHANGED Or SWP_NOACTIVATE
            
            'Specify the main form as the tabstrip's parent
            SetParent pdWindows(winIndex).hWnd, mainWindowHwnd
            
            'Force the window to properly align within the client area of the parent
            getActualMainFormClientRect childRect, True, False
            
            With pdWindows(winIndex)
            
                If .alignmentOptions.alignLeft Then MoveWindow .hWnd, childRect.x1, childRect.y1, getWindowWidth(.hWnd), childRect.y2 - childRect.y1, 1 'IIf(.isVisible, 1, 0)
                If .alignmentOptions.alignTop Then MoveWindow .hWnd, childRect.x1, childRect.y1, childRect.x2 - childRect.x1, getWindowHeight(.hWnd), 1 'IIf(.isVisible, 1, 0)
                If .alignmentOptions.alignRight Then MoveWindow .hWnd, childRect.x2 - getWindowWidth(.hWnd), childRect.y1, getWindowWidth(.hWnd), childRect.y2 - childRect.y1, 1 'IIf(.isVisible, 1, 0)
                If .alignmentOptions.alignBottom Then MoveWindow .hWnd, childRect.x1, childRect.y2 - getWindowHeight(.hWnd), childRect.x2 - childRect.x1, getWindowHeight(.hWnd), 1 'IIf(.isVisible, 1, 0)
            
            End With
            
            'Force a full repaint (to make sure children controls are also repainted)
            If pdWindows(winIndex).isVisible Then forceFullWindowRefresh winIndex
            
    End Select
    
    'Finally, update this window's rect, and store its offset (so that the children windows will stay in their proper location)
    requestWindowResync winIndex

End Sub

'Get floating/docked state of toolbars or image windows
Public Function getFloatState(ByVal winType As pdWindowType) As Boolean

    Select Case winType
    
        Case TOOLBAR_WINDOW
            getFloatState = toolbarsFloating
        
    End Select

End Function

'Set floating/docked state of toolbars or image windows
Public Sub setFloatState(ByVal winType As pdWindowType, ByVal isFloating As Boolean)

    Dim i As Long
    
    'Before doing anything else, reset all z-order data
    For i = 0 To numOfPDWindows - 1
        If pdWindows(i).typeOfWindow <> MAIN_WINDOW Then
            SetWindowPos pdWindows(i).hWnd, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOACTIVATE
        End If
    Next i
    
    Select Case winType
    
        'Toolbars
        Case TOOLBAR_WINDOW
                    
            'Set the main tracking variable
            toolbarsFloating = isFloating
            
    End Select
    
    'Whenever float status is changed, all windows need to be redrawn (as their layering and/or position may have changed)
    If (numOfPDWindows > 0) Then
        For i = 0 To numOfPDWindows - 1
            
            'Request a window style update and/or redraw
            updateWindowStyle i
            
        Next i
    End If
    
End Sub

'If a window had location data previously stored, this function will retrieve that data and move the window into place.
' If window data cannot be found, the form will be moved to the requested pseudo-location supplied by the second parameter.
Public Function restoreWindowLocation(ByVal windowIndex As Long, ByVal preferredLocationIfNecessary As Long) As Boolean

    'Start by looking for this form's location data in the XML engine.
    
    'If an entry is found, restore the window to that location.
    If xmlEngine.doesTagExist("windowEntry", "id", pdWindows(windowIndex).windowName) Then
    
        'Retrieve this window's location data from the XML file.
        With pdWindows(windowIndex)
        
            'Window rect values
            .lastPosition.x1 = CLng(xmlEngine.getUniqueTag_String("windowLeft", 0, , "windowEntry", "id", .windowName))
            .lastPosition.y1 = CLng(xmlEngine.getUniqueTag_String("windowTop", 0, , "windowEntry", "id", .windowName))
            .lastPosition.x2 = CLng(xmlEngine.getUniqueTag_String("windowRight", 0, , "windowEntry", "id", .windowName))
            .lastPosition.y2 = CLng(xmlEngine.getUniqueTag_String("windowBottom", 0, , "windowEntry", "id", .windowName))
            
            'Certain toolbars are not resizable.  Instead of relying on previous rect values, supply our own.
            If (.typeOfWindow = TOOLBAR_WINDOW) And (.optionalToolbarType = FILE_TOOLBOX) Then
                If toolbarsFloating Then
                    .lastPosition.x2 = .lastPosition.x1 + pdWindows(windowIndex).originalWidthFull
                Else
                    .lastPosition.x2 = .lastPosition.x1 + pdWindows(windowIndex).originalWidthClient
                End If
            ElseIf (.typeOfWindow = TOOLBAR_WINDOW) And (.optionalToolbarType = TOOLS_TOOLBOX) Then
                If toolbarsFloating Then
                    .lastPosition.y2 = .lastPosition.y1 + pdWindows(windowIndex).originalHeightFull
                Else
                    .lastPosition.y2 = .lastPosition.y1 + pdWindows(windowIndex).originalHeightClient
                End If
            ElseIf (.typeOfWindow = TOOLBAR_WINDOW) And (.optionalToolbarType = LAYER_TOOLBOX) Then
                If toolbarsFloating Then
                    .lastPosition.x2 = .lastPosition.x1 + pdWindows(windowIndex).originalWidthFull
                Else
                    .lastPosition.x2 = .lastPosition.x1 + pdWindows(windowIndex).originalWidthClient
                End If
            End If
            
            'Window and border state
            .lastWindowState = CLng(xmlEngine.getUniqueTag_String("windowState", 0, , "windowEntry", "id", .windowName))
            .lastWindowBorder = CLng(xmlEngine.getUniqueTag_String("windowStyle", 0, , "windowEntry", "id", .windowName))
            
            'Apply those values to the window, starting with border style, then state, then coordinates
            .formReference.BorderStyle = .lastWindowBorder
            .formReference.WindowState = .lastWindowState
            
            'Alignment data is currently only used for the image tabstrip.  In the future, it may be enabled for other toolbars.
            If .typeOfWindow = IMAGE_TABSTRIP Then
                .alignmentOptions.alignLeft = CBool(xmlEngine.getUniqueTag_String("alignLeft", "False", , "windowEntry", "id", .windowName))
                .alignmentOptions.alignTop = CBool(xmlEngine.getUniqueTag_String("alignTop", "True", , "windowEntry", "id", .windowName))
                .alignmentOptions.alignRight = CBool(xmlEngine.getUniqueTag_String("alignRight", "False", , "windowEntry", "id", .windowName))
                .alignmentOptions.alignBottom = CBool(xmlEngine.getUniqueTag_String("alignBottom", "False", , "windowEntry", "id", .windowName))
            End If
            
            'Make sure the location values will result in an on-screen form.  If they will not (for example, if the user detached a
            ' secondary monitor on which PhotoDemon was being used), change the values to ensure this window appears on-screen.
            
            'Note that this check relies on access to PD's g_cMonitors class, which returns the full virtual desktop dimensions, and
            ' not just the primary monitor's (as VB's Screen object does).
            Dim winWidth As Long, winHeight As Long
            winWidth = .lastPosition.x2 - .lastPosition.x1
            winHeight = .lastPosition.y2 - .lastPosition.y1
            
            If (.lastPosition.x1 + winWidth) < g_cMonitors.DesktopLeft Then
                .lastPosition.x1 = g_cMonitors.DesktopLeft
                .lastPosition.x2 = .lastPosition.x1 + winWidth
            End If
            
            If .lastPosition.x1 > (g_cMonitors.DesktopLeft + g_cMonitors.DesktopWidth) Then
                .lastPosition.x1 = (g_cMonitors.DesktopWidth - winWidth)
                .lastPosition.x2 = .lastPosition.x1 + winWidth
            End If
            
            If .lastPosition.y1 < g_cMonitors.DesktopTop Then
                .lastPosition.y1 = g_cMonitors.DesktopTop
                .lastPosition.y2 = .lastPosition.y1 + winHeight
            End If
            
            If .lastPosition.y1 > g_cMonitors.DesktopHeight Then
                .lastPosition.y1 = (g_cMonitors.DesktopHeight - winHeight)
                .lastPosition.y2 = .lastPosition.y1 + winHeight
            End If
            
            MoveWindow .hWnd, .lastPosition.x1, .lastPosition.y1, .lastPosition.x2 - .lastPosition.x1, .lastPosition.y2 - .lastPosition.y1, 1
            
        End With
        
    'Window data was not found.  Move the window to a new location using the preferred location param.
    Else
    
        If pdWindows(windowIndex).typeOfWindow = IMAGE_TABSTRIP Then
            pdWindows(windowIndex).alignmentOptions.alignTop = True
        Else
            requestIdealPosition pdWindows(windowIndex).hWnd, preferredLocationIfNecessary
        End If
    
    End If

End Function

'Load previous window locations from file.
Public Function loadAllWindowLocations() As Boolean
    
    If FileExist(windowDataPath) Then
            
        'Attempt to load and validate the relevant preset file; if we can't, create a new, blank XML object
        If (Not xmlEngine.loadXMLFile(windowDataPath)) Or Not (xmlEngine.isPDDataType("Window locations")) Then
            Debug.Print "No window location data found.  A new window location file has been created."
            resetXMLData
        End If
            
    Else
        resetXMLData
    End If
    
    'We don't actually load window locations now.  Now that the XML data is safely inside our XML engine, we load window data from it
    ' on-demand as windows are added to the window manager.

End Function

'Write the current locations of all windows to the XML engine.  (These will be used to restore the window location on subsequent loads.)
Public Function saveAllWindowLocations() As Boolean
    
    Dim i As Long
    For i = 0 To numOfPDWindows - 1
    
        'We save locations for everything but image windows
        If Not (pdWindows(i).typeOfWindow = IMAGE_WINDOW) Then
    
            'Check for this window preset in the file.  If it does not exist, add it now.
            If Not xmlEngine.doesTagExist("windowEntry", "id", pdWindows(i).windowName) Then
            
                xmlEngine.writeTagWithAttribute "windowEntry", "id", pdWindows(i).windowName, "", True
                xmlEngine.closeTag "windowEntry"
                xmlEngine.writeBlankLine
            
            End If
            
            'Write this window's location data into the XML file.
            With pdWindows(i)
            
                'Window rect values
                xmlEngine.updateTag "windowLeft", .lastPosition.x1, "windowEntry", "id", .windowName
                xmlEngine.updateTag "windowTop", .lastPosition.y1, "windowEntry", "id", .windowName
                xmlEngine.updateTag "windowRight", .lastPosition.x2, "windowEntry", "id", .windowName
                xmlEngine.updateTag "windowBottom", .lastPosition.y2, "windowEntry", "id", .windowName
                
                'Window and border state
                xmlEngine.updateTag "windowState", .lastWindowState, "windowEntry", "id", .windowName
                xmlEngine.updateTag "windowStyle", .lastWindowBorder, "windowEntry", "id", .windowName
                
                'Alignment; this is only used by the image tabstrip at present, but may be enabled for other windows
                ' in the future.
                xmlEngine.updateTag "alignLeft", CStr(.alignmentOptions.alignLeft), "windowEntry", "id", .windowName
                xmlEngine.updateTag "alignTop", CStr(.alignmentOptions.alignTop), "windowEntry", "id", .windowName
                xmlEngine.updateTag "alignRight", CStr(.alignmentOptions.alignRight), "windowEntry", "id", .windowName
                xmlEngine.updateTag "alignBottom", CStr(.alignmentOptions.alignBottom), "windowEntry", "id", .windowName
                
            End With
        
        End If
    
    Next i
    
    'Write the data out to file
    xmlEngine.writeXMLToFile windowDataPath

End Function

'Reset the XML engine.  Note that the XML object SHOULD ALREADY BE INSTANTIATED before calling this function.
Private Function resetXMLData()

    xmlEngine.prepareNewXML "Window locations"
    xmlEngine.writeBlankLine
    xmlEngine.writeComment "Everything past this point is window location data for various PhotoDemon dialogs."
    xmlEngine.writeBlankLine
    
End Function

'This class provides some helper functions to intelligently position child windows, typically only used if toolbars are in float
' mode and we can't locate a saved window location file.  Constants are used to specify where the window should be placed.
Public Sub requestIdealPosition(ByVal childHwnd As Long, Optional ByVal desiredPosition As Long = 1)

    'Before doing anything else, get the current size and location of child and parent windows
    Dim childRect As winRect, parentRect As winRect, parentClientRect As winRect
    GetWindowRect childHwnd, childRect
    GetWindowRect mainWindowHwnd, parentRect
    GetClientRect mainWindowHwnd, parentClientRect

    'By comparing the parent window's client area and actual area, we can determine how much is "chrome", e.g. window borders, etc.
    Dim parentChromeWidth As Long, parentChromeHeight As Long
    parentChromeWidth = (parentRect.x2 - parentRect.x1) - parentClientRect.x2
    parentChromeHeight = (parentRect.y2 - parentRect.y1) - parentClientRect.y2
    
    'Right now, we only support a handful of "desired positions".  I may add more after additional testing.
    Select Case desiredPosition
    
        'Do nothing
        Case 0
    
        'Top-left
        Case 1
            MoveWindow childHwnd, parentRect.x1 + parentChromeWidth \ 2, parentRect.y1 + parentChromeHeight, childRect.x2 - childRect.x1, childRect.y2 - childRect.y1, 1
            
        'Top-right
        Case 2
            MoveWindow childHwnd, parentRect.x2 - parentChromeWidth \ 2 - (childRect.x2 - childRect.x1), parentRect.y1 + parentChromeHeight, childRect.x2 - childRect.x1, childRect.y2 - childRect.y1, 1
            
        'Bottom-left
        Case 3
            MoveWindow childHwnd, parentRect.x1 + parentChromeWidth \ 2, parentRect.y2 - (childRect.y2 - childRect.y1), childRect.x2 - childRect.x1, childRect.y2 - childRect.y1, 1
        
        'Bottom-right
        Case 4
            MoveWindow childHwnd, parentRect.x2 - parentChromeWidth \ 2 - (childRect.x2 - childRect.x1), parentRect.y2 - (childRect.y2 - childRect.y1), childRect.x2 - childRect.x1, childRect.y2 - childRect.y1, 1
    
    End Select

End Sub

'The window manager keeps a persistent list of last-known locations for every window it accesses.  Window locations are updated when:
' - They're first loaded (and no previous position has been stored)
' - They're moved
' - They're closed
'Any of those actions just need to send a request to this sub, which will handle the actual location storage process.
Private Sub updateStoredWindowLocation(ByVal indexInWindowArray As Long)

    Dim newWindowRect As winRect
    
    With pdWindows(indexInWindowArray)
    
        GetWindowRect .hWnd, newWindowRect
        .lastPosition = newWindowRect
        If Not (.formReference Is Nothing) Then
            .lastWindowBorder = .formReference.BorderStyle
            .lastWindowState = .formReference.WindowState
        End If
    
    End With

End Sub

'The primary PhotoDemon form must register its hWnd, so we can track its movement and move any children windows accordingly.
Public Sub registerParentForm(ByRef parentForm As Form)

    'Has this window already been registered?  If so, don't re-register it.
    Dim i As Long
    For i = 0 To numOfPDWindows
        If pdWindows(i).hWnd = parentForm.hWnd Then Exit Sub
    Next i
    
    'Add this window to the collection, and cache its hWnd (because we reference the main window handle frequently)
    Set pdWindows(numOfPDWindows).formReference = parentForm
    pdWindows(numOfPDWindows).hWnd = parentForm.hWnd
    mainWindowHwnd = parentForm.hWnd
    mainWindowIndex = numOfPDWindows
    pdWindows(numOfPDWindows).typeOfWindow = MAIN_WINDOW
    pdWindows(numOfPDWindows).windowName = parentForm.Name
    
    'Apply any unique styles to the parent window
    Dim newWinStyle As Long
    newWinStyle = GetWindowLong(parentForm.hWnd, GWL_STYLE)
    SetWindowLong parentForm.hWnd, GWL_STYLE, newWinStyle Or WS_CLIPCHILDREN
    newWinStyle = GetWindowLong(parentForm.hWnd, GWL_EXSTYLE)
    SetWindowLong parentForm.hWnd, GWL_EXSTYLE, newWinStyle Or WS_EX_COMPOSITED
    
    'Subclass the main window
    cSubclass.ssc_Subclass mainWindowHwnd, , , Me
    cSubclass.ssc_AddMsg mainWindowHwnd, MSG_AFTER, WM_MOUSEMOVE, WM_NCMOUSEMOVE, WM_SIZE  ', WM_MOVE - adding WM_MOVE causes freezing at present, and I'm not sure why, so leave it unsubclassed for the time being
    cSubclass.ssc_AddMsg mainWindowHwnd, MSG_BEFORE, WM_SIZING, WM_MOVING
    cSubclass.ssc_AddMsg mainWindowHwnd, MSG_BEFORE_AFTER, WM_SYSCOMMAND
    
    'Because I care about PD being a good "desktop citizen", I deactivate top-most status of all toolbars (and other windows, as necessary),
    ' when PD loses focus.  To do that successfully, we must intercept app activation messages.
    cSubclass.ssc_AddMsg mainWindowHwnd, MSG_BEFORE, WM_ACTIVATEAPP, WM_ACTIVATE
    
    'Look for previous location data in the window location file.  If said data exists, load it and move the window to that location.
    restoreWindowLocation numOfPDWindows, 0
    
    'Resize the window array for future additions
    numOfPDWindows = numOfPDWindows + 1
    ReDim Preserve pdWindows(0 To numOfPDWindows) As pdWindowData
        
End Sub

'Each child window must be registered through this class, using its hWnd.  This class will then use that hWnd to do things
' like move child windows proportionally to their parent.
Public Sub registerChildForm(ByRef childForm As Form, ByVal whatWindowType As pdWindowType, Optional ByVal preferredLocationIfNecessary As Long = 1, Optional ByVal toolbarType As pdToolbarType = NOT_A_TOOLBAR, Optional imageWindowIndex As Long = -1, Optional ByVal minWidth As Long = -1, Optional ByVal minHeight As Long = -1)

    'Has this window already been registered?  If so, don't re-register it.
    Dim i As Long
    For i = 0 To numOfPDWindows
        If pdWindows(i).hWnd = childForm.hWnd Then Exit Sub
    Next i
    
    'Add this window to the collection
    Set pdWindows(numOfPDWindows).formReference = childForm
    pdWindows(numOfPDWindows).hWnd = childForm.hWnd
    
    'Calculate the initial offset between this window and the main window, and store that offset
    Dim childRect As winRect, parentRect As winRect
    GetWindowRect pdWindows(numOfPDWindows).hWnd, childRect
    GetWindowRect mainWindowHwnd, parentRect
    
    With pdWindows(numOfPDWindows)
        .isVisible = True
        .minimumWidth = minWidth
        .minimumHeight = minHeight
        .offsetX = childRect.x1 - parentRect.x1
        .offsetY = childRect.y1 - parentRect.y1
        .optionalToolbarType = toolbarType
        
        'The width and height of toolbars is handled differently.  Instead of storing last-used width and height
        ' (which may be incorrect if monitor DPI has changed since the last session), force them to their initial
        ' width/height, which VB will have automatically resolved to match the current DPI setting.
        .originalWidthFull = getWindowWidth(pdWindows(numOfPDWindows).hWnd)
        .originalHeightFull = getWindowHeight(pdWindows(numOfPDWindows).hWnd)
        .originalWidthClient = getClientWidth(pdWindows(numOfPDWindows).hWnd)
        .originalHeightClient = getClientHeight(pdWindows(numOfPDWindows).hWnd)
        
        .hasBeenMadeTransparent = False
        .typeOfWindow = whatWindowType
        .windowName = childForm.Name
    End With
        
    'Subclass this window, and add a user param of 64 + "index in array".  We can use this value to recognize this hWnd as belonging
    ' to a child window, and quickly look it up in our window tracking array.
    cSubclass.ssc_Subclass pdWindows(numOfPDWindows).hWnd, 64 + numOfPDWindows, , Me
    cSubclass.ssc_AddMsg pdWindows(numOfPDWindows).hWnd, MSG_AFTER, WM_MOUSEMOVE, WM_NCMOUSEMOVE, WM_ACTIVATE, WM_MOVE, WM_SIZE
    cSubclass.ssc_AddMsg pdWindows(numOfPDWindows).hWnd, MSG_BEFORE, WM_MOVING, WM_SIZING
    
    'Paint this window with the proper visual style (based on whether it's docked or floating)
    updateWindowStyle numOfPDWindows, True
    
    'Move the window into place using either last-known location data, or the supplied "preferred location"
    If pdWindows(numOfPDWindows).typeOfWindow <> IMAGE_WINDOW Then restoreWindowLocation numOfPDWindows, preferredLocationIfNecessary
    
    'Update this window's offset (because it may have been moved by the above function)
    requestWindowResync numOfPDWindows
    
    'If this is the custom image tabstrip window, make a note of its hWnd and array location, because we will refer to it frequently
    ' and this is faster than constantly searching the window array for it.
    If pdWindows(numOfPDWindows).typeOfWindow = IMAGE_TABSTRIP Then
        imageTabstripHwnd = pdWindows(numOfPDWindows).hWnd
        imageTabstripIndex = numOfPDWindows
    End If
    
    'Resize the window array for future additions
    numOfPDWindows = numOfPDWindows + 1
    ReDim Preserve pdWindows(0 To numOfPDWindows) As pdWindowData
    
End Sub

'When a window is unloaded, call this function so that we can stop subclassing in a safe and predictable way, and also track the
' last-known location of this window.
Public Sub unregisterForm(ByRef srcForm As Form)

    'Find this window in the array.  (Note that if it cannot be found, nothing will happen.)
    Dim i As Long
    For i = 0 To numOfPDWindows
        If pdWindows(i).hWnd = srcForm.hWnd Then
            
            'Update the last-known window location before the form closes
            updateStoredWindowLocation i
            
            'Mark the window as invisible
            pdWindows(i).isVisible = False
            
            'Release this window's subclassing
            pdWindows(i).subclassingReleased = True
            cSubclass.ssc_UnSubclass pdWindows(i).hWnd
            pdWindows(i).hWnd = 0
            
            'Remove our reference to this form (if we don't, it will remain in memory)
            Set pdWindows(i).formReference = Nothing
            
            Exit For
            
        End If
    Next i

End Sub

'At present, we handle all window maintenance on-the-fly, so nothing special needs to be done when all images are unloaded.
' If extra work does need to be performed in the future, the program already ties into this function.
Public Sub allImageWindowsUnloaded()
    
End Sub

Public Sub requestActivation(ByVal hWndToActivate As Long)
    SetWindowPos hWndToActivate, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_SHOWWINDOW Or SWP_NOOWNERZORDER
End Sub

Private Sub Class_Initialize()

    'Reset all tracking variables
    numOfPDWindows = 0
    ReDim pdWindows(0) As pdWindowData
    
    numOfExtraWindows = 0
    ReDim extraWindowHwnds(0) As Long
    
    'Prepare the subclasser
    Set cSubclass = New cSelfSubHookCallback
    
    'Prepare the XML handler, and retrieve window location data from file (if it exists)
    Set xmlEngine = New pdXML
    windowDataPath = g_UserPreferences.getPresetPath & "Program_WindowLocations.xml"
    loadAllWindowLocations
    
    'To prevent circular redraws, we disable certain subclassing when we know we are responsible for triggering the window messages.  Initialize
    ' that tracking variable to FALSE.
    mainFormTriggeredChange = False
    
End Sub

Private Sub Class_Terminate()

    'Release all subclassing
    cSubclass.ssc_Terminate
    
    'Write window location data out to file.
    saveAllWindowLocations

End Sub

'Set a toolbar window's translucency.  We do this to make them semi-transparent when they do not have the mouse over them.
Private Sub setWindowTranslucency(ByVal srcHwnd As Long, Optional ByVal nTransparency As Long = 255)

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

'Window z-order is a difficult thing to get right.  It's a complicated interplay of what is floating, what is not, and it must
' all be disabled when PD loses focus.  Outside functions can call this with any window index to retrieve the proper hWndBefore
' parameter for a SetWindowPos call.
Private Function getDesiredZOrder(ByVal winIndex As Long, Optional ByVal pdLosingFocus As Boolean = False) As Long

     'A window can receive four z-order instructions.  Unfortunately, Windows makes it very difficult to provide detailed z-order instructions.
    ' Instead, we are limited to very broad classes of "top-most", "not-top-most", and "bottom".
    ' Descriptions are taken directly from this MSDN page: http://msdn.microsoft.com/en-us/library/windows/desktop/ms633545%28v=vs.85%29.aspx
    
    'Place the window at the top of the Z order.
    'Const HWND_TOP As Long = 0

    'Place the window above all non-topmost windows. The window maintains its topmost position even when it is deactivated.
    'Const HWND_TOPMOST As Long = -1
    
    'Place the window above all non-topmost windows but behind all topmost windows. This flag has no effect if the window is already
    ' a non-topmost window.
    'Const HWND_NOTOPMOST As Long = -2
    
    'Place the window at the bottom of the Z order. If the hWnd parameter identifies a topmost window, the window loses its topmost status
    ' and is placed at the bottom of all other windows.
    'Const HWND_BOTTOM As Long = 1

    'If PhotoDemon is losing focus, this step is simple - restore a normal z-order and exit
    If pdLosingFocus Then
        getDesiredZOrder = HWND_TOP
        Exit Function
    End If
    
    'If the program is not losing focus, this step is more complicated. What window is made top-most is determined
    ' by what is docked and what is not.  Roughly, speaking:
    
    'If images are docked, and toolbars are docked, z-order doesn't matter
    'If images are docked, and toolbars are floating, make toolbars top-most
    'If images are floating, and toolbars are floating, make toolbars top-most
    'If images are floating, and toolbars are docked, make images top-most
    
    'By default, windows are given a default "bring to top, but NOT top-most" order
    getDesiredZOrder = HWND_TOP
    
    Dim curWindowType As pdWindowType
    curWindowType = pdWindows(winIndex).typeOfWindow
    
    'If toolbars are floating, make them top-most
    If toolbarsFloating Then
        If curWindowType = TOOLBAR_WINDOW Then
            getDesiredZOrder = HWND_TOPMOST
        End If
    End If
    
End Function

'When the mouse moves over the main canvas, use this function to make floating toolbars transparent; similarly, use it to
' restore toolbar opacity when the mouse leaves
Public Sub notifyMouseMoveOverCanvas(Optional ByVal mouseIsOver As Boolean = True)

    'Translucency only matters if toolbars are floating; otherwise, we can ignore these requests entirely
    If toolbarsFloating Then
    
        Dim i As Long
    
        If mouseIsOver Then
        
            For i = 0 To numOfPDWindows - 1
                If (pdWindows(i).typeOfWindow = TOOLBAR_WINDOW) And (Not pdWindows(i).hasBeenMadeTransparent) Then
                    pdWindows(i).hasBeenMadeTransparent = True
                    setWindowTranslucency pdWindows(i).hWnd, INACTIVE_TOOLBOX_OPACITY
                End If
            Next i
        
        'We don't actually need to handle a FALSE case for mouse over, as translucency is restored by mouse detection
        ' on the toolbars themselves.
        Else
        
        End If
        
    End If

End Sub

'When the user has resized the image tabstrip, we need to redraw all image windows to match (but toolbars are okay).
Public Sub notifyImageTabStripResized()
    
    If Not mainFormTriggeredChange Then
    
        'Ask the main form's canvas to redraw itself
        FormMain.refreshAllCanvases
        
    End If
    
End Sub

'Whenever a size change occurs that requires multiple windows to be redrawn, this function can be called.  It will automatically handle redraw duties.
Private Sub requestWindowRedraw(ByVal winIndex As Long, ByRef curWinRect As winRect)

    Dim childRect As winRect
    GetWindowRect pdWindows(winIndex).hWnd, childRect
                    
    Select Case pdWindows(winIndex).typeOfWindow
    
        'Main window
        Case MAIN_WINDOW
            forceFullWindowRefresh winIndex
    
        'Toolbars
        Case TOOLBAR_WINDOW
            If toolbarsFloating Then
                MoveWindow pdWindows(winIndex).hWnd, curWinRect.x1 + pdWindows(winIndex).offsetX, curWinRect.y1 + pdWindows(winIndex).offsetY, childRect.x2 - childRect.x1, childRect.y2 - childRect.y1, 1 'IIf(pdWindows(winIndex).isVisible, 1, 0)
            Else
                
                'Because the toolbars are not floating, they will have alignment data stored.  Use that to calculate their
                ' position instead.
                setWindowPosByAlignment winIndex, True
                
                'Stretch the toolbars to fit the full width/height of the main form (as applicable)
                fixDockedToolbarSize winIndex
                
            End If
            
        'Image tabstrip
        Case IMAGE_TABSTRIP
            'Force the window to top-align within the client area of the parent
            getActualMainFormClientRect childRect, True, False
            
            With pdWindows(winIndex)
            
                If .alignmentOptions.alignLeft Then MoveWindow .hWnd, childRect.x1, childRect.y1, getWindowWidth(.hWnd), childRect.y2 - childRect.y1, 1 'IIf(.isVisible, 1, 0)
                If .alignmentOptions.alignTop Then MoveWindow .hWnd, childRect.x1, childRect.y1, childRect.x2 - childRect.x1, getWindowHeight(.hWnd), 1 'IIf(.isVisible, 1, 0)
                If .alignmentOptions.alignRight Then MoveWindow .hWnd, childRect.x2 - getWindowWidth(.hWnd), childRect.y1, getWindowWidth(.hWnd), childRect.y2 - childRect.y1, 1 'IIf(.isVisible, 1, 0)
                If .alignmentOptions.alignBottom Then MoveWindow .hWnd, childRect.x1, childRect.y2 - getWindowHeight(.hWnd), childRect.x2 - childRect.x1, getWindowHeight(.hWnd), 1 'IIf(.isVisible, 1, 0)
            
            End With
            
    End Select
                    
    'Regardless of the type of window being redrawn, it makes sense to resync the window's coordinates here (as a failsafe).
    requestWindowResync winIndex

End Sub

'Under certain circumstances (such as PD losing focus), it is important to remove the top-most status of any window types.  Use this
' function to do so.  If restoreOriginalZOrder is TRUE, the function will re-enable top-most status for any windows that require it.
Public Sub resetTopmostForAllWindows(Optional ByVal restoreOriginalZOrder As Boolean = False)

    Dim i As Long

    If restoreOriginalZOrder Then
    
        For i = 0 To numOfPDWindows - 1
            If pdWindows(i).isVisible Then
                SetWindowPos pdWindows(i).hWnd, getDesiredZOrder(i), 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOSENDCHANGING Or SWP_SHOWWINDOW
            End If
        Next i
        
        'If the program has registered any extra windows (e.g. modal dialogs), restore them last.  This will ensure that they still
        ' sit above any image or toolbar windows.
        If numOfExtraWindows > 0 Then
            For i = 0 To numOfExtraWindows - 1
                SetWindowPos extraWindowHwnds(i), HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOSENDCHANGING Or SWP_SHOWWINDOW
            Next i
        End If
        
    Else
        
        For i = 0 To numOfPDWindows - 1
            If pdWindows(i).isVisible Then
                SetWindowPos pdWindows(i).hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_SHOWWINDOW
            End If
        Next i
        
        'If the program has registered any extra windows (e.g. modal dialogs), reset them as well
        If numOfExtraWindows > 0 Then
            For i = 0 To numOfExtraWindows - 1
                SetWindowPos extraWindowHwnds(i), HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOSENDCHANGING
            Next i
        End If
        
    End If

End Sub

'As the name implies, move a given window to the last offset and size stored in its pdWindows entry.  This is helpful after a
' minimize/restore of the main window, for example, if either images or toolbars are floating.
Private Sub moveWindowToLastKnownLocation(ByVal windowIndex As Long, ByRef mainFormRect As winRect, Optional ByVal alsoRestoreSize As Boolean = False)

    Dim childRect As winRect
    GetWindowRect pdWindows(windowIndex).hWnd, childRect
    
    If alsoRestoreSize Then
        childRect = pdWindows(windowIndex).lastPosition
    End If
    
    'NOTE: Win 7 classic theme will need extra debugging, as I have just discovered.  Desktop redraws are handled COMPLETELY differently,
    ' so optimizations for Aero can cause all kinds of horrible distortions on classic theme!  Ugh....
    Select Case pdWindows(windowIndex).typeOfWindow
    
        Case TOOLBAR_WINDOW
            If toolbarsFloating Then MoveWindow pdWindows(windowIndex).hWnd, mainFormRect.x1 + pdWindows(windowIndex).offsetX, mainFormRect.y1 + pdWindows(windowIndex).offsetY, childRect.x2 - childRect.x1, childRect.y2 - childRect.y1, 1 'IIf(pdWindows(windowIndex).isVisible, 1, 0)
            
    End Select
    
    'Previously we would request a resync of position after this, but that is no longer necessary.
    'requestWindowResync i

End Sub

'If the main program knows an action is likely to screw up window appearance (such as forcible hide/restore during screen capture),
' this function can be used to refresh all windows.
Public Sub refreshAllWindows()

    'Refresh each tracked window
    Dim i As Long
    For i = 0 To numOfPDWindows - 1
        forceFullWindowRefresh i
    Next i

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
    
    Dim userSizeTooSmall As Boolean
    Dim resizeHandle As winResizeEdge
    
    Dim i As Long
    
    'Before doing anything else, check the user-supplied param.  If it is >= 64, we know this is a child (toolbar or image) window.
    ' (We know this because we add the user parameter when subclassing starts, and we add it using the system 64 + "index in array").
    If lParamUser >= 64 Then
    
        'Retrieve the window's index in our master tracking array
        Dim winIndex As Long
        winIndex = lParamUser - 64
    
        'Check child window messages
        Select Case uMsg
            
            'The window has moved
            Case WM_MOVE, WM_SIZE
            
                If Not insideMinimizeMode Then
            
                    'Debug.Print "Received move or size event for " & pdWindows(winIndex).formReference.Caption
                
                    'Update this window's position in our position database
                    If Not mainFormTriggeredChange Then requestWindowResync winIndex
                    
                End If
            
            'The window is currently being moved
            Case WM_MOVING, WM_SIZING
                
                'Debug.Print "Received moving or sizing event for " & pdWindows(winIndex).formReference.Caption
                
                'If mainFormTriggeredChange Then Exit Sub
                
                'lParam contains a pointer to this window's rect location.  Retrieve it now.
                Call CopyMemory(curWinRect, ByVal lParam, Len(curWinRect))
                
                'Note that in the case of WM_SIZING, we want to make sure the user is not stretching the window below an acceptable
                ' minimum size.  Check that now.
                If (uMsg = WM_SIZING) And ((pdWindows(winIndex).minimumWidth >= 0) Or (pdWindows(winIndex).minimumHeight >= 0)) Then
                
                    userSizeTooSmall = False
                    
                    'wParam contains a value that identifies the edge being used for the resize.
                    
                    resizeHandle = wParam
                    
                    'If the user has set the window size too small, force the window to a minimum size (640x480).
                    If (curWinRect.x2 - curWinRect.x1) < pdWindows(winIndex).minimumWidth Then
                        userSizeTooSmall = True
                        If (resizeHandle = WMSZ_BOTTOMLEFT) Or (resizeHandle = WMSZ_LEFT) Or (resizeHandle = WMSZ_TOPLEFT) Then
                            curWinRect.x1 = curWinRect.x2 - pdWindows(winIndex).minimumWidth
                        Else
                            curWinRect.x2 = curWinRect.x1 + pdWindows(winIndex).minimumWidth
                        End If
                    End If
                    
                    If (curWinRect.y2 - curWinRect.y1) < pdWindows(winIndex).minimumHeight Then
                        userSizeTooSmall = True
                        If (resizeHandle = WMSZ_TOPLEFT) Or (resizeHandle = WMSZ_TOP) Or (resizeHandle = WMSZ_TOPRIGHT) Then
                            curWinRect.y1 = curWinRect.y2 - pdWindows(winIndex).minimumHeight
                        Else
                            curWinRect.y2 = curWinRect.y1 + pdWindows(winIndex).minimumHeight
                        End If
                    End If
                    
                    'Copy the new rect into place
                    If userSizeTooSmall Then CopyMemory ByVal lParam, curWinRect, Len(curWinRect)
                    
                End If
                
                'Calculate this window's rectangle as an offset of its parent.
                Dim parentRect As winRect
                If pdWindows(winIndex).typeOfWindow = TOOLBAR_WINDOW And (Not toolbarsFloating) Then
                    GetClientRect mainWindowHwnd, parentRect
                Else
                    GetWindowRect mainWindowHwnd, parentRect
                End If
                                
                'Store the new offset values
                pdWindows(winIndex).offsetX = curWinRect.x1 - parentRect.x1
                pdWindows(winIndex).offsetY = curWinRect.y1 - parentRect.y1
                
            'The mouse is moving inside this child window.
            Case WM_MOUSEMOVE, WM_NCMOUSEMOVE
            
                Select Case pdWindows(winIndex).typeOfWindow
                
                    'If this is a toolbar window, and it was previously made transparent, restore its opacity now.
                    Case TOOLBAR_WINDOW
                        If pdWindows(winIndex).hasBeenMadeTransparent And toolbarsFloating Then
                            pdWindows(winIndex).hasBeenMadeTransparent = False
                            setWindowTranslucency lng_hWnd, 255
                        End If
                    
                    'If this is an image window, and we have non-transparent floating toolbar windows, change that now.
                    Case IMAGE_TABSTRIP
                        For i = 0 To numOfPDWindows - 1
                            If (pdWindows(i).typeOfWindow = TOOLBAR_WINDOW) And (Not pdWindows(i).hasBeenMadeTransparent) And toolbarsFloating Then
                                pdWindows(i).hasBeenMadeTransparent = True
                                setWindowTranslucency pdWindows(i).hWnd, INACTIVE_TOOLBOX_OPACITY
                            End If
                        Next i
                
                End Select
                
        End Select
    
    'The user param is not >= 64, meaning it references the primary window (FormMain).  Move all child windows to match the
    ' parent form's new position.
    Else
    
        Dim childRect As winRect
        
        'Check parent window messages
        Select Case uMsg
        
            'The parent window is currently being moved
            Case WM_MOVING, WM_MOVE
            
                'lParam contains a pointer to the window's rect location.  Retrieve it now.
                Call CopyMemory(curWinRect, ByVal lParam, Len(curWinRect))
                
                mainFormTriggeredChange = True
                
                'Move each child window, while retaining the current offset.  Note that docked (child) windows do not have to
                ' be manually moved - Windows takes care of this for us.
                For i = 0 To numOfPDWindows - 1
                    moveWindowToLastKnownLocation i, curWinRect
                Next i
                
                mainFormTriggeredChange = False
                
                'Remember the main window's location
                updateStoredWindowLocation mainWindowIndex
                
                'Notify the color management engine that the main window may have changed monitors, so it may need to
                ' be redrawn with a new color profile.
                If g_OpenImageCount > 0 Then checkParentMonitor
                
            'The parent window is currently being resized
            Case WM_SIZING, WM_SIZE
                
                mainFormTriggeredChange = True
                
                'Note that in the case of WM_SIZING, we want to make sure the user is not stretching the window below an acceptable
                ' minimum size.  Check that now.
                If uMsg = WM_SIZING Then
                
                    userSizeTooSmall = False
                    
                    'wParam contains a value that identifies the edge being used for the resize.
                    resizeHandle = wParam
                    
                    'lParam contains a pointer to the window's rect location.  Retrieve it now.
                    CopyMemory curWinRect, ByVal lParam, Len(curWinRect)
                    
                    'If the user has set the window size too small, force the window to a minimum size (640x480).
                    If (curWinRect.x2 - curWinRect.x1) < 640 Then
                        userSizeTooSmall = True
                        If (resizeHandle = WMSZ_BOTTOMLEFT) Or (resizeHandle = WMSZ_LEFT) Or (resizeHandle = WMSZ_TOPLEFT) Then
                            curWinRect.x1 = curWinRect.x2 - 640
                        Else
                            curWinRect.x2 = curWinRect.x1 + 640
                        End If
                    End If
                    
                    If (curWinRect.y2 - curWinRect.y1) < 480 Then
                        userSizeTooSmall = True
                        If (resizeHandle = WMSZ_TOPLEFT) Or (resizeHandle = WMSZ_TOP) Or (resizeHandle = WMSZ_TOPRIGHT) Then
                            curWinRect.y1 = curWinRect.y2 - 480
                        Else
                            curWinRect.y2 = curWinRect.y1 + 480
                        End If
                    End If
                    
                    If userSizeTooSmall Then
                    
                        'Copy the new rect into place
                        CopyMemory ByVal lParam, curWinRect, Len(curWinRect)
                        
                        'Because subsequent functions need access to the new rect data, we cheat a bit here - we use SetWindowPos to immediately
                        ' update the main window's rect, but we combine flags so that no new window messages are sent (to prevent infinite recursion of
                        ' window changes causing more window changes).
                        SetWindowPos lng_hWnd, 0, curWinRect.x1, curWinRect.y1, curWinRect.x2 - curWinRect.x1, curWinRect.y2 - curWinRect.y1, SWP_NOACTIVATE Or SWP_NOZORDER 'Or SWP_NOSENDCHANGING
                        
                    End If
                
                End If
                
                'Re-retrieve the current window rect for subsequent functions
                GetWindowRect lng_hWnd, curWinRect
                
                'Remember the main window's location
                updateStoredWindowLocation mainWindowIndex
                
                'Resize and/or move each child window as necessary
                For i = 0 To numOfPDWindows - 1
                    
                    requestWindowRedraw i, curWinRect
                    
                Next i
                
                mainFormTriggeredChange = False
                               
                                
            'The parent window has just been maximized, minimized, or restored
            Case WM_SYSCOMMAND
                
                'Per MSDN (http://msdn.microsoft.com/en-us/library/windows/desktop/ms646360%28v=vs.85%29.aspx):
                ' In WM_SYSCOMMAND messages, the four low-order bits of the wParam parameter are used internally by the system.
                ' To obtain the correct result when testing the value of wParam, an application must combine the value 0xFFF0 with
                ' the wParam value by using the bitwise AND operator.
                wParam = wParam And WM_SYSCOMMAND_WPARAM_MASK
                
                'WM_SYSCOMMAND functions are difficult to handle, especially on Win 7 where Aero Snap features interact with them
                ' in completely unpredictable ways.  This code block (which can be safely commented out) is helpful for seeing
                ' what messages the window manager receives from different actions.
                If Not g_IsProgramCompiled Then
                    If (wParam = SC_MAXIMIZE) Then Debug.Print "Received system maximize message " & bBefore
                    If (wParam = SC_RESTORE) Then Debug.Print "Received system restore window message " & bBefore
                    If (wParam = SC_SIZE) Then Debug.Print "Received system window size message " & bBefore
                    If (wParam = SC_MOVE) Then Debug.Print "Received system window move message " & bBefore
                End If
                
                If (wParam = SC_MAXIMIZE) Or (wParam = SC_RESTORE) Or (wParam = SC_SIZE) Or (wParam = SC_MOVE) Then
                    
                    If Not bBefore Then
                        
                        'Manually retrieve the new window position.
                        GetWindowRect lng_hWnd, curWinRect
                        
                        'If the window was previously minimized, restore any child windows now
                        For i = 0 To numOfPDWindows - 1
                            If pdWindows(i).forciblyMinimized Then
                                pdWindows(i).forciblyMinimized = False
                                pdWindows(i).formReference.WindowState = vbNormal
                                moveWindowToLastKnownLocation i, curWinRect, True
                            End If
                        Next i
                        
                        'Resize each child window, while retaining its current offset
                        For i = 0 To numOfPDWindows - 1
                            requestWindowRedraw i, curWinRect
                            
                            'Toolbar windows may not redraw properly after restoration, so forcibly update them now
                            If pdWindows(i).typeOfWindow = TOOLBAR_WINDOW Then forceFullWindowRefresh i
                            
                        Next i
                        
                        If (wParam = SC_MAXIMIZE) Or (wParam = SC_RESTORE) Then
                            Debug.Print "LEAVING MINIMIZED MODE"
                            insideMinimizeMode = False
                        End If
                        
                        'Because image window zOrder may get botched by this action, forcibly restore the active image to the foreground.
                        'restoreActiveImageToForeground "PD maximized or restored"
                                            
                    End If
                    
                'When minimizing, check all windows, and if one has not been minimized, minimize it now
                ElseIf wParam = SC_MINIMIZE Then
                
                    If bBefore Then
                
                        Debug.Print "ENTERING MINIMIZED MODE"
                        insideMinimizeMode = True
                    
                        For i = 0 To numOfPDWindows - 1
                            If Not (pdWindows(i).formReference Is Nothing) Then
                                If pdWindows(i).formReference.WindowState <> vbMinimized Then
                                    Debug.Print "forcibly minimizing window: " & pdWindows(i).formReference.Caption
                                    
                                    'Update the window's location before minimizing it
                                    requestWindowResync i
                                    pdWindows(i).forciblyMinimized = True
                                    pdWindows(i).formReference.WindowState = vbMinimized
                                    
                                End If
                            End If
                        Next i
                    
                    End If
                    
                End If
                    
            'Mouse has been moved.  If toolbars are floating, make them translucent.
            Case WM_MOUSEMOVE, WM_NCMOUSEMOVE
            
                If Not bBefore Then
                
                    For i = 0 To numOfPDWindows - 1
                        If (pdWindows(i).typeOfWindow = TOOLBAR_WINDOW) And (Not pdWindows(i).hasBeenMadeTransparent) And toolbarsFloating Then
                            pdWindows(i).hasBeenMadeTransparent = True
                            setWindowTranslucency pdWindows(i).hWnd, INACTIVE_TOOLBOX_OPACITY
                        End If
                    Next i
                    
                End If
                
            'The user is switching to or from PD and another application
            Case WM_ACTIVATEAPP
            
                If bBefore Then
            
                    'wParam contains a notification of whether we are gaining or losing focus.
                    Select Case wParam
                    
                        'PD is losing focus; make all top-most windows not topmost
                        Case WA_INACTIVE
                            Debug.Print "> PhotoDemon lost focus! <"
                            resetTopmostForAllWindows False
                            isPDDeactivated = True
                            
                        'PD is gaining focus
                        Case Else
                            Debug.Print "> PhotoDemon regained focus! <"
                            
                            If isPDDeactivated Then
                                resetTopmostForAllWindows True
                                isPDDeactivated = False
                            End If
                            
                    End Select
                    
                End If
                        
        End Select
    
    End If




' *************************************************************
' C A U T I O N   C A U T I O N   C A U T I O N   C A U T I O N
' -------------------------------------------------------------
' DO NOT ADD ANY OTHER CODE BELOW THE "END SUB" STATEMENT BELOW
'   add this warning banner to the last routine in your class
' *************************************************************
End Sub



