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
'Last updated: 10/October/13
'Last update: many minor fixes to improve performance and reliability
'
'PhotoDemon first began as a single-pane, single-image editor.  About a decade ago, I rewrote it as an MDI project, and in
' 2013, it became time to rewrite its window method once again - this time, to a true multi-window project, separate from
' MDI.  (MDI has a lot of interaction problems, not to mention a host of known bugs that will never be fixed because
' Microsoft is - rightfully - moving away from the MDI model.)
'
'One unfortunate result of this change is that PhotoDemon must now be responsible for much of its own window handling.  In
' particular, child windows must move with the parent window, and all document windows must be managed intelligently.
'
'Thus this class.  The main form, toolbar forms, and image forms are all registered with this handler when they are created.
' This class manages their placement, styles, opacity, and a host of other minor details.
'
'In order to perform all this management, this class subclasses quite a few window messages.  It should be IDE-safe, but
' consider yourself warned.
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
Private Const WM_MOVING As Long = &H216
Private Const WM_MOUSELEAVE As Long = &H2A3
Private Const WM_NCMOUSEMOVE As Long = &HA0
Private Const WM_SIZING As Long = &H214

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

'Constants for changing window style (including transparency)
Private Const LWA_ALPHA As Long = &H2&
Private Const WS_EX_LAYERED As Long = &H80000
Private Const GWL_EXSTYLE As Long = (-20)
Private Const GWL_STYLE As Long = (-16)

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

Public Enum enWindowStyles
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
    lastPosition As winRect
    lastWindowState As Long
    lastWindowBorder As Long
    offsetX As Long
    offsetY As Long
    optionalToolbarType As pdToolbarType
    originalWidth As Long
    originalHeight As Long
    subclassingReleased As Boolean
    typeOfWindow As pdWindowType
    windowName As String
End Type

'This class tracks a lot of different window types.  Each gets special treatment, so we track their type as follows:
Public Enum pdWindowType
    MAIN_WINDOW = 0
    TOOLBAR_WINDOW = 1
    IMAGE_WINDOW = 2
End Enum

#If False Then
    Private Const MAIN_WINDOW = 0, TOOLBAR_WINDOW = 1, IMAGE_WINDOW = 2
#End If

Public Enum pdToolbarType
    NOT_A_TOOLBAR = 0
    MAIN_TOOLBOX = 1
    SELECTION_TOOLBOX = 2
End Enum

#If False Then
    Private Const NOT_A_TOOLBAR = 0, MAIN_TOOLBOX = 1, SELECTION_TOOLBOX = 2
#End If

'When a window is being resized, it is helpful to know which edge is being resized.  Windows uses the following system, which we mimic:
Public Enum winResizeEdge
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

'API calls for window positioning and tracking
Private Declare Function ClientToScreen Lib "user32" (ByVal hndWindow As Long, ByRef lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hndWindow As Long, ByRef scrPoint As POINTAPI) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hndWindow As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hndWindow As Long, ByRef lpRect As winRect) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hndWindow As Long, ByRef lpRect As winRect) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
      
'API calls for changing window style
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal targetHwnd As Long, ByVal TransColor As Long, ByVal newAlpha As Byte, ByVal fFlags As Long) As Boolean
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal targetHwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal targetHwnd As Long, ByVal nIndex As Long) As Long
Private Declare Sub SetWindowPos Lib "user32" (ByVal targetHwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
'Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function BeginDeferWindowPos Lib "user32" (ByVal nNumWindows As Long) As Long
Private Declare Function DeferWindowPos Lib "user32" (ByVal hWinPosInfo As Long, ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function EndDeferWindowPos Lib "user32" (ByVal hWinPosInfo As Long) As Long
Private Declare Function InvalidateRect Lib "user32" (ByVal targetHwnd As Long, ByRef lpRect As Any, ByVal bErase As Long) As Long
Private Declare Function UpdateWindow Lib "user32" (ByVal targetHwnd As Long) As Long
Private Declare Function AdjustWindowRect Lib "user32" (ByRef lpRect As winRect, ByVal dwStyle As Long, ByVal bMenu As Long) As Long

'Child window tracking
Private pdWindows() As pdWindowData
Private numOfPDWindows As Long

'Primary window tracking
Private mainWindowHwnd As Long, mainWindowIndex As Long

'When PD loses focus, we automatically revert all top-level toolbar windows to normal style.  When PD regains focus, we restore them as top-most.
' Activation/deactivation messages are received multiple times, but we only want to do activation/deactivation stuff once, so our current state
' is tracked by this variable.
Private isPDDeactivated As Boolean

'We must subclass various window types in order to manage them properly
Private cSubclass As cSelfSubHookCallback

'XML handling (used to save/load window locations) is handled through a specialized class
Private xmlEngine As pdXML
Private windowDataPath As String

'Toolbars and child windows can be floated or docked (as a group - not individually!  e.g. all toolbars can be docked, or all can be floated)
' These variables track the floating/docked state.
Private toolbarsFloating As Boolean, imagesFloating As Boolean

'When the main form is being moved or resized, it will trigger the movement of all child forms to match.  We set this variable to TRUE
' so that child forms don't waste time and energy handling window move messages that have already been handled.
Private mainFormTriggeredChange As Boolean

'If external (or internal) functions move windows around on their own, they can manually request resyncing the window's offset.
Public Sub requestWindowResync(ByVal winIndex As Long)
    
    'Get the parent's window rect
    Dim parentRect As winRect
    GetWindowRect mainWindowHwnd, parentRect
    
    'Get the child's window rect
    Dim childRect As winRect
    GetWindowRect pdWindows(winIndex).hWnd, childRect
    
    'Store the new offset values
    pdWindows(winIndex).offsetX = childRect.x1 - parentRect.x1
    pdWindows(winIndex).offsetY = childRect.y1 - parentRect.y1
    
End Sub

'External functions can have us retrieve a window rect for them
Friend Sub getWindowRectByIndex(ByVal winIndex As Long, ByRef srcRect As winRect)
    GetWindowRect pdWindows(winIndex).hWnd, srcRect
End Sub

'Similar to the functions below, this sub can be used to retrieve a rect that contains (in *screen coordinates*) the position of the main
' form's "actual" client rect, by which I mean - the rect accounting for docked toolbars, as necessary.
Friend Sub getActualMainFormClientRect(ByRef modifiedRect As winRect)

    'Start by retrieving the main form's client rect, in screen coordinates.
    Dim clientRect As winRect, toolbarRect As winRect
    getClientCoordsRelativeToScreen mainWindowHwnd, clientRect
        
    'If toolbars are docked, subtract their width from the rect
    If Not toolbarsFloating Then
        
        Dim i As Long
        For i = 0 To numOfPDWindows - 1
            If pdWindows(i).typeOfWindow = TOOLBAR_WINDOW Then
                
                GetWindowRect pdWindows(i).hWnd, toolbarRect
            
                'Use the toolbar's alignment to determine if its width should be subtracted from the left or right of the client area
                If pdWindows(i).alignmentOptions.alignLeft Then
                    clientRect.x1 = clientRect.x1 + (toolbarRect.x2 - toolbarRect.x1)
                Else
                    clientRect.x2 = clientRect.x2 - (toolbarRect.x2 - toolbarRect.x1)
                End If
                
            End If
        Next i
        
    End If
    
    'Always subtract the progress bar's size from the client area
    clientRect.y2 = clientRect.y2 - FormMain.picProgBar.ScaleHeight
    
    'Provide a 1px border around the image
    clientRect.x1 = clientRect.x1 + 1
    clientRect.x2 = clientRect.x2 - 1
    clientRect.y1 = clientRect.y1 + 1
    clientRect.y2 = clientRect.y2 - 1
    
    modifiedRect = clientRect

End Sub

'External functions (such as those that auto-zoom an image to fit it on-screen) need to know the size of the client area on the main form.
' That client size will change based on the docking status of toolbars, so we must use the window manager to calculate it.
Public Function requestActualMainFormClientWidth() As Long

    Dim clientRect As winRect
    GetClientRect mainWindowHwnd, clientRect

    'When toolbars are floating, the main window's actual client size can be reported via GetClientRect
    If toolbarsFloating Then
        requestActualMainFormClientWidth = clientRect.x2 - clientRect.x1
    
    'When toolbars are docked, we must remove all toolbar widths from the client size.
    Else
    
        Dim newClientSize As Long
        newClientSize = clientRect.x2 - clientRect.x1
        
        Dim toolbarRect As winRect
        
        'Remove all toolbar widths from the client rect
        Dim i As Long
        For i = 0 To numOfPDWindows - 1
            If pdWindows(i).typeOfWindow = TOOLBAR_WINDOW Then
                GetWindowRect pdWindows(i).hWnd, toolbarRect
                newClientSize = newClientSize - (toolbarRect.x2 - toolbarRect.x1)
            End If
        Next i
        
        requestActualMainFormClientWidth = newClientSize
    
    End If

End Function

'At present, the docking status of toolbars doesn't affect the main form's client height, making this calculation quite easy.
Public Function requestActualMainFormClientHeight() As Long

    Dim clientRect As winRect
    GetClientRect mainWindowHwnd, clientRect
    
    requestActualMainFormClientHeight = clientRect.y2 - FormMain.picProgBar.ScaleHeight

End Function

'If a toolbar is docked, it will also be aligned.  We can use alignment data to position it instead of an x/y offset.
Private Sub setWindowPosByAlignment(ByVal winIndex As Long)

    Dim newPositionRect As winRect
        
    'Retrieve the window's current size
    Dim curWinRect As winRect
    GetWindowRect pdWindows(winIndex).hWnd, curWinRect
    
    'Retrieve the client area of the main form
    Dim mainClientRect As winRect
    getClientCoordsRelativeToScreen mainWindowHwnd, mainClientRect
        
    'At present, we assume that a window is either aligned top or bottom, right or left.  (e.g. it must be aligned to a corner)
    ' This makes positioning much simpler.
    
    With pdWindows(winIndex)
    
        'Horizontal alignment
        If .alignmentOptions.alignLeft Then newPositionRect.x1 = mainClientRect.x1
        If .alignmentOptions.alignRight Then newPositionRect.x1 = mainClientRect.x2 - (curWinRect.x2 - curWinRect.x1)
        
        'Vertical alignment
        If .alignmentOptions.alignTop Then newPositionRect.y1 = mainClientRect.y1
        If .alignmentOptions.alignBottom Then newPositionRect.y1 = mainClientRect.y2 - (curWinRect.y2 - curWinRect.y1)
    
    End With
    
    'Calculate the new position x2 and y2 to match the window's current width and height
    newPositionRect.x2 = newPositionRect.x1 + (curWinRect.x2 - curWinRect.x1)
    newPositionRect.y2 = newPositionRect.y1 + (curWinRect.y2 - curWinRect.y1)
    
    'Move the window into place
    MoveWindow pdWindows(winIndex).hWnd, newPositionRect.x1, newPositionRect.y1, newPositionRect.x2 - newPositionRect.x1, newPositionRect.y2 - newPositionRect.y1, 1
    
    'Resync the window's offset
    requestWindowResync winIndex

End Sub

'Shortcut function to determine if a window has one or more aligned statuses set.  If it does, those supercede the usual offset-tracking methods.
Private Function isWindowAligned(ByVal winIndex As Long) As Boolean
    With pdWindows(winIndex).alignmentOptions
        isWindowAligned = .alignBottom Or .alignLeft Or .alignRight Or .alignTop
    End With
End Function

'If toolbars are docked, their height (and/or width, if I ever add horizontal toolbars) needs to be kept in sync with the main form.
' After the main form is resized, this function can be called to resync all toolbar width/heights.
Private Sub fixDockedToolbarSize(ByVal winIndex As Long)

    'If toolbars are not docked, exit.
    If toolbarsFloating Then Exit Sub

    'Resize the window to fit inside the parent form, but not overlap the progress bar at the bottom
    Dim modifiedClientHeight As Long
    modifiedClientHeight = getClientHeight(mainWindowHwnd)
    modifiedClientHeight = modifiedClientHeight - FormMain.picProgBar.ScaleHeight
    
    'In rare circumstances, it's possible for the user to shrink the main form horizontally to the point that the main menu overflows
    ' into a second row.  To accomodate this case, also recalculate the proper top client coordinate.
    Dim curMainFormClient As winRect
    getClientCoordsRelativeToScreen mainWindowHwnd, curMainFormClient
    
    Dim childRect As winRect
    
    If pdWindows(winIndex).typeOfWindow = TOOLBAR_WINDOW Then
    
        'Get this toolbar window's current location and size
        GetWindowRect pdWindows(winIndex).hWnd, childRect
        
        'Change the height, then resize the window
        childRect.y1 = curMainFormClient.y1
        childRect.y2 = curMainFormClient.y1 + modifiedClientHeight
        SetWindowPos pdWindows(winIndex).hWnd, 0, childRect.x1, curMainFormClient.y1, childRect.x2 - childRect.x1, childRect.y2 - childRect.y1, SWP_NOACTIVATE Or SWP_NOOWNERZORDER Or SWP_NOZORDER
        
    End If

End Sub

'Given a window, return its horizontal (or vertical) client size, in pixels.
Public Function getClientWidth(ByVal targetHwnd As Long) As Long
    Dim tmpRect As winRect
    GetClientRect targetHwnd, tmpRect
    getClientWidth = tmpRect.x2
End Function

Public Function getClientHeight(ByVal targetHwnd As Long) As Long
    Dim tmpRect As winRect
    GetClientRect targetHwnd, tmpRect
    getClientHeight = tmpRect.y2
End Function

'Given a window, return its horizontal (or vertical) size, in pixels.
Private Function getWindowWidth(ByVal targetHwnd As Long) As Long
    Dim tmpRect As winRect
    GetWindowRect targetHwnd, tmpRect
    getWindowWidth = tmpRect.x2 - tmpRect.x1
End Function

Private Function getWindowHeight(ByVal targetHwnd As Long) As Long
    Dim tmpRect As winRect
    GetWindowRect targetHwnd, tmpRect
    getWindowHeight = tmpRect.y2 - tmpRect.y1
End Function

'Given a window, find the size differential between the window and client areas.
' These two functions *should* return accurate results regardless of OS visual settings.
Public Function getHorizontalChromeSize(ByVal targetHwnd As Long) As Long

    'Ask the OS to report window and client size
    Dim WindowRECT As winRect, clientRect As winRect
    GetWindowRect targetHwnd, WindowRECT
    GetClientRect targetHwnd, clientRect
    
    getHorizontalChromeSize = (WindowRECT.x2 - WindowRECT.x1) - clientRect.x2
    
End Function

Public Function getVerticalChromeSize(ByVal targetHwnd As Long) As Long

    'Ask the OS to report window and client size
    Dim WindowRECT As winRect, clientRect As winRect
    GetWindowRect targetHwnd, WindowRECT
    GetClientRect targetHwnd, clientRect
    
    getVerticalChromeSize = (WindowRECT.y2 - WindowRECT.y1) - clientRect.y2
    
End Function

'Force a full window refresh (invalidate + update)
Private Sub forceFullWindowRefresh(ByVal winIndex As Long)
    InvalidateRect pdWindows(winIndex).hWnd, ByVal 0&, 0
    UpdateWindow pdWindows(winIndex).hWnd
End Sub

'Retrieve the client coordinates of a window, but relative to the *screen* (e.g. with the top-left set to screen coords instead of the default (0,0)
Private Sub getClientCoordsRelativeToScreen(ByVal targetHwnd As Long, ByRef rectToFill As winRect)

    Dim curClientRect As winRect
    GetClientRect targetHwnd, curClientRect
    
    'Client rect is always based at (0,0); to convert it to screen coordinates, we use the ClientToScreen API function.
    Dim clientLeft As POINTAPI, clientRight As POINTAPI
    
    'Retrieve the top-left corner first.  Note that we don't have to fill clientLeft, because the top-left client coordinates are always (0,0)
    ClientToScreen targetHwnd, clientLeft
    
    'Fill the bottom-right coord object, then convert it to screen coords
    clientRight.x = curClientRect.x2
    clientRight.y = curClientRect.y2
    ClientToScreen targetHwnd, clientRight
    
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
    
    'Because we will access the main form's location and client location under a variety of circumstances, retrieve their values now
    GetWindowRect mainWindowHwnd, mainFormRect
    getClientCoordsRelativeToScreen mainWindowHwnd, mainFormClientRect
    
    Select Case pdWindows(winIndex).typeOfWindow
    
        'Toolbar windows
        Case TOOLBAR_WINDOW
                    
            'FLOATING STYLE
            If toolbarsFloating Then
            
                GetWindowRect pdWindows(winIndex).hWnd, childRect
                
                'Set the border style to fixed toolbar
                SetWindowLong pdWindows(winIndex).hWnd, GWL_STYLE, WS_BORDER Or WS_VISIBLE Or WS_CAPTION Or WS_SYSMENU 'Or WS_CHILD 'Or WS_CLIPCHILDREN
                SetWindowLong pdWindows(winIndex).hWnd, GWL_EXSTYLE, WS_EX_TOOLWINDOW
                
                'Using this new border style, calculate a desired window rectangle while forcing the client size to the window's
                ' original dimensions.  (Note that this does not matter the *first* time a floating window is rendered, because
                ' the window will already be in floating-style.)
                If Not isFirstLoad Then
                    childRect.x2 = childRect.x1 + pdWindows(winIndex).originalWidth
                    childRect.y2 = childRect.y1 + pdWindows(winIndex).originalHeight
                    AdjustWindowRect childRect, GetWindowLong(pdWindows(winIndex).hWnd, GWL_STYLE), 0
                End If
                
                'Use that new rect to resize the image
                SetWindowPos pdWindows(winIndex).hWnd, 0, childRect.x1, childRect.y1, childRect.x2 - childRect.x1, childRect.y2 - childRect.y1, SWP_FRAMECHANGED
                
                'Erase any alignment data, as floating windows cannot currently be aligned
                pdWindows(winIndex).alignmentOptions.alignTop = False
                pdWindows(winIndex).alignmentOptions.alignBottom = False
                pdWindows(winIndex).alignmentOptions.alignLeft = False
                pdWindows(winIndex).alignmentOptions.alignRight = False
                
                'Force a full repaint (to make sure children controls are also repainted)
                forceFullWindowRefresh winIndex
                
            'DOCKED STYLE
            Else
                
                'Before changing the border style, determine the client area.  We will use this to resize the window after removing its border.
                Dim clientDimensions As winRect
                GetClientRect pdWindows(winIndex).hWnd, clientDimensions
                
                'Remove the window border
                SetWindowLong pdWindows(winIndex).hWnd, GWL_STYLE, WS_VISIBLE Or WS_CHILD 'Or WS_CLIPCHILDREN
                SetWindowLong pdWindows(winIndex).hWnd, GWL_EXSTYLE, WS_EX_NOACTIVATE
                SetWindowPos pdWindows(winIndex).hWnd, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_FRAMECHANGED
                
                'Restore full window opacity
                SetWindowTranslucency pdWindows(winIndex).hWnd, 255
                
                'Get an updated window rect
                GetWindowRect pdWindows(winIndex).hWnd, childRect
            
                'Embed the toolbar in a specific section of the main form
                Select Case pdWindows(winIndex).optionalToolbarType
                
                    'Main toolbar gets embedded in the top-left
                    Case MAIN_TOOLBOX
                        MoveWindow pdWindows(winIndex).hWnd, mainFormClientRect.x1, mainFormClientRect.y1, pdWindows(winIndex).originalWidth, pdWindows(winIndex).originalHeight, 1
                        pdWindows(winIndex).alignmentOptions.alignTop = True
                        pdWindows(winIndex).alignmentOptions.alignLeft = True
                        
                    'Selection toolbar gets embedded in the top-right
                    Case SELECTION_TOOLBOX
                        MoveWindow pdWindows(winIndex).hWnd, mainFormClientRect.x2 - pdWindows(winIndex).originalWidth, mainFormClientRect.y1, pdWindows(winIndex).originalWidth, pdWindows(winIndex).originalHeight, 1
                        pdWindows(winIndex).alignmentOptions.alignTop = True
                        pdWindows(winIndex).alignmentOptions.alignRight = True
                        
                End Select
                
                'Force docked toolbars to the full height of the parent form's client area
                fixDockedToolbarSize winIndex
                
                'Force a full repaint (to make sure children controls are also repainted)
                forceFullWindowRefresh winIndex
                
            End If
                                    
        'Image windows
        Case IMAGE_WINDOW
        
            'FLOATING STYLE
            If imagesFloating Then
            
                GetWindowRect pdWindows(winIndex).hWnd, childRect
                
                'Set the border style to sizable dialog
                SetWindowLong pdWindows(winIndex).hWnd, GWL_STYLE, WS_THICKFRAME Or WS_VISIBLE Or WS_CAPTION Or WS_SYSMENU Or WS_CLIPCHILDREN
                
                'Using this new border style, calculate a desired window rectangle while forcing the client size to the window's
                ' original dimensions.  (Note that this does not matter the *first* time a floating window is rendered, because
                ' the window will already be in floating-style.)
                'If Not isFirstLoad Then
                '    childRect.x2 = childRect.x1 + pdWindows(winIndex).originalWidth
                '    childRect.y2 = childRect.y1 + pdWindows(winIndex).originalHeight
                '    AdjustWindowRect childRect, GetWindowLong(pdWindows(winIndex).hWnd, GWL_STYLE), 0
                'End If
                
                'Use that new rect to resize the image
                SetWindowPos pdWindows(winIndex).hWnd, 0, childRect.x1, childRect.y1, childRect.x2 - childRect.x1, childRect.y2 - childRect.y1, SWP_FRAMECHANGED
                
                'Erase any alignment data, as floating windows cannot currently be aligned
                pdWindows(winIndex).alignmentOptions.alignTop = False
                pdWindows(winIndex).alignmentOptions.alignBottom = False
                pdWindows(winIndex).alignmentOptions.alignLeft = False
                pdWindows(winIndex).alignmentOptions.alignRight = False
                
                'Force a full repaint (to make sure children controls are also repainted)
                forceFullWindowRefresh winIndex
                
            'DOCKED STYLE
            Else
            
                'Remove the window border
                SetWindowLong pdWindows(winIndex).hWnd, GWL_STYLE, WS_VISIBLE Or WS_CLIPCHILDREN
                SetWindowPos pdWindows(winIndex).hWnd, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_FRAMECHANGED
                
                'Force the window to take up the full client area of the parent
                getActualMainFormClientRect childRect
                MoveWindow pdWindows(winIndex).hWnd, childRect.x1, childRect.y1, childRect.x2 - childRect.x1, childRect.y2 - childRect.y1, 1
                
                'Force a full repaint (to make sure children controls are also repainted)
                forceFullWindowRefresh winIndex
            
            End If
            
    End Select
    
    'Finally, update this window's rect, and store its offset (so that the children windows will stay in their proper location)
    requestWindowResync winIndex

End Sub

'Get floating/docked state of toolbars or image windows
Public Function getFloatState(ByVal winType As pdWindowType) As Boolean

    Select Case winType
    
        Case TOOLBAR_WINDOW
            getFloatState = toolbarsFloating
        
        Case IMAGE_WINDOW
            getFloatState = imagesFloating
    
    End Select

End Function

'Set floating/docked state of toolbars or image windows
Public Sub setFloatState(ByVal winType As pdWindowType, ByVal isFloating As Boolean)

    Dim i As Long
    
    Select Case winType
    
        'Toolbars
        Case TOOLBAR_WINDOW
                    
            'Set the main tracking variable
            toolbarsFloating = isFloating
                    
            'If toolbars have been loaded, reset their position accordingly
            If (numOfPDWindows > 0) Then
            
                'Loop through all windows in the manager, and only adjust toolbar windows
                For i = 0 To numOfPDWindows - 1
                    If pdWindows(i).typeOfWindow = TOOLBAR_WINDOW Then updateWindowStyle i
                Next i
            
            End If
            
            'In the case of docked toolbars, we need to make sure their parent is properly assigned.
            If Not toolbarsFloating Then
                For i = 0 To numOfPDWindows - 1
                    If pdWindows(i).typeOfWindow = TOOLBAR_WINDOW Then
                        SetWindowLong pdWindows(i).hWnd, -8, mainWindowHwnd
                        SetWindowPos pdWindows(i).hWnd, -2, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOSENDCHANGING Or SWP_SHOWWINDOW
                    End If
                Next i
            End If
            
            'When toolbar floating status is changed, image windows also need to be repainted.
            For i = 0 To numOfPDWindows - 1
                If pdWindows(i).typeOfWindow = IMAGE_WINDOW Then updateWindowStyle i
            Next i
        
        'Image windows
        Case IMAGE_WINDOW
        
            'Set the main tracking variable
            imagesFloating = isFloating
        
            'If images have been loaded, reset their position accordingly
            If (numOfPDWindows > 0) Then
        
                'Loop through all windows in the manager, and only adjust image windows
                For i = 0 To numOfPDWindows - 1
                    If pdWindows(i).typeOfWindow = IMAGE_WINDOW Then updateWindowStyle i
                Next i
                
            End If
    
    End Select
    
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
            
            'Window and border state
            .lastWindowState = CLng(xmlEngine.getUniqueTag_String("windowState", 0, , "windowEntry", "id", .windowName))
            .lastWindowBorder = CLng(xmlEngine.getUniqueTag_String("windowStyle", 0, , "windowEntry", "id", .windowName))
            
            'Apply those values to the window, starting with border style, then state, then coordinates
            .formReference.BorderStyle = .lastWindowBorder
            .formReference.WindowState = .lastWindowState
            
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
    
        requestIdealPosition pdWindows(windowIndex).hWnd, preferredLocationIfNecessary
    
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
    
        'We only save locations for the main window and toolbar windows - not image windows!
        If (pdWindows(i).typeOfWindow = MAIN_WINDOW) Or (pdWindows(i).typeOfWindow = TOOLBAR_WINDOW) Then
    
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

'This class provides some helper functions to intelligently position child windows, typically when the program is run for the
' first time and custom positions have not yet been established.  Constants are used to specify where the window should be placed.
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
    
    'Right now, we only support two "desired positions": top-left and top-right corners.  I may add more after additional testing.
    Select Case desiredPosition
    
        'Do nothing
        Case 0
    
        'Top-left
        Case 1
            MoveWindow childHwnd, parentRect.x1 + parentChromeWidth \ 2, parentRect.y1 + parentChromeHeight, childRect.x2 - childRect.x1, childRect.y2 - childRect.y1, 1
            
        'Top-right
        Case 2
            MoveWindow childHwnd, parentRect.x2 - parentChromeWidth \ 2 - (childRect.x2 - childRect.x1), parentRect.y1 + parentChromeHeight, childRect.x2 - childRect.x1, childRect.y2 - childRect.y1, 1
    
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
        .lastWindowBorder = .formReference.BorderStyle
        .lastWindowState = .formReference.WindowState
    
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
    
    'Subclass the main window
    cSubclass.ssc_Subclass mainWindowHwnd, , , Me
    cSubclass.ssc_AddMsg mainWindowHwnd, MSG_AFTER, WM_SYSCOMMAND, WM_MOUSEMOVE, WM_NCMOUSEMOVE, WM_SIZING, WM_MOVING
    
    'In the future, it may be helpful to know when PD loses focus (e.g. to dump images from RAM to the hard drive).
    ' Thus I am leaving this code here "just in case".
    'cSubclass.ssc_AddMsg parentHwnd, MSG_AFTER, WM_ACTIVATEAPP, WM_ACTIVATE
    
    'Look for previous location data in the window location file.  If said data exists, load it and move the window to that location.
    restoreWindowLocation numOfPDWindows, 0
    
    'Resize the window array for future additions
    numOfPDWindows = numOfPDWindows + 1
    ReDim Preserve pdWindows(0 To numOfPDWindows) As pdWindowData
    
End Sub

'Each child window must be registered through this class, using its hWnd.  This class will then use that hWnd to do things
' like move child windows proportionally to their parent.
Public Sub registerChildForm(ByRef childForm As Form, ByVal whatWindowType As pdWindowType, Optional ByVal preferredLocationIfNecessary As Long = 1, Optional ByVal toolbarType As pdToolbarType = NOT_A_TOOLBAR, Optional imageWindowIndex As Long = -1)

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
        .offsetX = childRect.x1 - parentRect.x1
        .offsetY = childRect.y1 - parentRect.y1
        .optionalToolbarType = toolbarType
        .originalWidth = getClientWidth(pdWindows(numOfPDWindows).hWnd)
        .originalHeight = getClientHeight(pdWindows(numOfPDWindows).hWnd)
        .hasBeenMadeTransparent = False
        .typeOfWindow = whatWindowType
        .windowName = childForm.Name
    End With
        
    'Subclass this window, and add a user param of 64 + "index in array".  We can use this value to recognize this hWnd as belonging
    ' to a child window, and quickly look it up in our window tracking array.
    cSubclass.ssc_Subclass pdWindows(numOfPDWindows).hWnd, 64 + numOfPDWindows, , Me
    cSubclass.ssc_AddMsg pdWindows(numOfPDWindows).hWnd, MSG_AFTER, WM_MOVING, WM_MOUSEMOVE, WM_NCMOUSEMOVE, WM_SIZING
    
    'Move the window into place using either last-known location data, or the supplied "preferred location"
    restoreWindowLocation numOfPDWindows, preferredLocationIfNecessary
    
    'Paint this window with the proper visual style (based on whether it's docked or floating)
    updateWindowStyle numOfPDWindows, True
        
    'Update this window's offset (because it may have been moved by the above function)
    requestWindowResync numOfPDWindows
    
    'If this is an image window, notify the corresponding pdImages() entry of this window's location in the window manager.
    If pdWindows(numOfPDWindows).typeOfWindow = IMAGE_WINDOW Then pdImages(imageWindowIndex).indexInWindowManager = numOfPDWindows
    
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
            
            'Release this window's subclassing
            pdWindows(i).subclassingReleased = True
            cSubclass.ssc_UnSubclass pdWindows(i).hWnd
            pdWindows(i).hWnd = 0
            
            'Remove our reference to this form (if we don't, it will remain in memory)
            Set pdWindows(i).formReference = Nothing
            
        End If
    Next i

End Sub

'When all child windows have been unloaded, we need to reassign toolbars as children of the main form.  (Otherwise, they will get unloaded!)
Public Sub allImageWindowsUnloaded()

    Dim i As Long
    For i = 0 To numOfPDWindows - 1
        If pdWindows(i).typeOfWindow = TOOLBAR_WINDOW Then
            SetWindowLong pdWindows(i).hWnd, -8, mainWindowHwnd
            SetWindowPos pdWindows(i).hWnd, -2, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOSENDCHANGING Or SWP_SHOWWINDOW
        End If
    Next i
    
End Sub

'When an image window receives focus, we dynamically reassign the toolbars as children of that window, so that they stay on top of it.
Public Sub notifyChildReceivedFocus(ByRef childForm As Form)

    If toolbarsFloating Then

        'Make all toolbars a child of this form - this keeps them on top
        Dim i As Long
        For i = 0 To numOfPDWindows - 1
            If pdWindows(i).typeOfWindow = TOOLBAR_WINDOW Then
                SetWindowLong pdWindows(i).hWnd, -8, childForm.hWnd
                SetWindowPos pdWindows(i).hWnd, -2, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOSENDCHANGING Or SWP_SHOWWINDOW
            End If
        Next i
        
        'Redraw the child as well
        SetWindowPos childForm.hWnd, -2, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOSENDCHANGING Or SWP_SHOWWINDOW
        
    End If

End Sub

Private Sub Class_Initialize()

    'Reset all toolbar tracking variables
    numOfPDWindows = 0
    ReDim pdWindows(0) As pdWindowData
        
    'Prepare the subclasser
    Set cSubclass = New cSelfSubHookCallback
    
    'Prepare the XML handler, and retrieve window location data from file (if it exists)
    Set xmlEngine = New pdXML
    windowDataPath = g_UserPreferences.getPresetPath & "Program_WindowLocations.xml"
    loadAllWindowLocations
    
    mainFormTriggeredChange = False
    
End Sub

Private Sub Class_Terminate()

    'Release all subclassing
    cSubclass.ssc_Terminate
    
    'Write window location data out to file.
    saveAllWindowLocations

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
    
    Dim i As Long
    
    'Before doing anything else, check the user-supplied param.  If it is >= 64, we know this is a child (toolbar or image) window.
    ' (We know this because we add the user parameter when subclassing starts, and we add it using the system 64 + "index in array").
    If lParamUser >= 64 Then
    
        'Retrieve the window's index in our master tracking array
        Dim winIndex As Long
        winIndex = lParamUser - 64
    
        'Check child window messages
        Select Case uMsg
        
            'The window is currently being moved
            Case WM_MOVING, WM_SIZING
                
                If mainFormTriggeredChange Then Exit Sub
                
                'lParam contains a pointer to this window's rect location.  Retrieve it now.
                Call CopyMemory(curWinRect, ByVal lParam, Len(curWinRect))
                
                'Calculate this window's rectangle as an offset of its parent.
                Dim parentRect As winRect
                GetWindowRect mainWindowHwnd, parentRect
                
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
                            SetWindowTranslucency lng_hWnd, 255
                        End If
                    
                    'If this is an image window, and we have non-transparent floating toolbar windows, change that now.
                    Case IMAGE_WINDOW
                        For i = 0 To numOfPDWindows - 1
                            If (pdWindows(i).typeOfWindow = TOOLBAR_WINDOW) And (Not pdWindows(i).hasBeenMadeTransparent) And toolbarsFloating Then
                                pdWindows(i).hasBeenMadeTransparent = True
                                SetWindowTranslucency pdWindows(i).hWnd, INACTIVE_TOOLBOX_OPACITY
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
            Case WM_MOVING
            
                'lParam contains a pointer to the window's rect location.  Retrieve it now.
                Call CopyMemory(curWinRect, ByVal lParam, Len(curWinRect))
                
                mainFormTriggeredChange = True
                
                'Move each child window, while retaining the current offset
                For i = 0 To numOfPDWindows - 1
                    GetWindowRect pdWindows(i).hWnd, childRect
                    MoveWindow pdWindows(i).hWnd, curWinRect.x1 + pdWindows(i).offsetX, curWinRect.y1 + pdWindows(i).offsetY, childRect.x2 - childRect.x1, childRect.y2 - childRect.y1, 1
                Next i
                
                mainFormTriggeredChange = False
                
                'Remember the main window's location
                updateStoredWindowLocation mainWindowIndex
            
            'The parent window is currently being resized
            Case WM_SIZING
                
                'Note that in the case of sizing, wParam contains a value that identifies the edge being used for the resize.
                ' Determine this via use of the winResizeEdge Enum.
                
                mainFormTriggeredChange = True
                
                'lParam contains a pointer to the window's rect location.  Retrieve it now.
                CopyMemory curWinRect, ByVal lParam, Len(curWinRect)
                
                'If the user has set the window size too small, force the window to a minimum size (640x480).
                If (curWinRect.x2 - curWinRect.x1) < 640 Then
                    curWinRect.x2 = curWinRect.x1 + 640
                    CopyMemory ByVal lParam, curWinRect, Len(curWinRect)
                End If
                
                If (curWinRect.y2 - curWinRect.y1) < 480 Then
                    curWinRect.y2 = curWinRect.y1 + 480
                    CopyMemory ByVal lParam, curWinRect, Len(curWinRect)
                End If
                
                'Because subsequent functions need access to the new rect data, we cheat a bit here - we use SetWindowPos to immediately
                ' update the main window's rect, but we combine flags so that no new window messages are sent (to prevent infinite recursion of
                ' window changes causing more window changes).
                SetWindowPos lng_hWnd, 0, curWinRect.x1, curWinRect.y1, curWinRect.x2 - curWinRect.x1, curWinRect.y2 - curWinRect.y1, SWP_NOSENDCHANGING Or SWP_NOACTIVATE
                
                'Move each child window, while retaining the current offset
                For i = 0 To numOfPDWindows - 1
                    GetWindowRect pdWindows(i).hWnd, childRect
                    
                    Select Case pdWindows(i).typeOfWindow
                    
                        'Toolbars
                        Case TOOLBAR_WINDOW
                            If toolbarsFloating Then
                                MoveWindow pdWindows(i).hWnd, curWinRect.x1 + pdWindows(i).offsetX, curWinRect.y1 + pdWindows(i).offsetY, childRect.x2 - childRect.x1, childRect.y2 - childRect.y1, 1
                            Else
                                
                                'Because the toolbars are not floating, they will have alignment data stored.  Use that to calculate their
                                ' position instead.
                                setWindowPosByAlignment i
                                
                                'Stretch the toolbars to fit the full width/height of the main form (as applicable)
                                fixDockedToolbarSize i
                                
                            End If
                    
                        'Images
                        Case IMAGE_WINDOW
                            If imagesFloating Then
                                MoveWindow pdWindows(i).hWnd, curWinRect.x1 + pdWindows(i).offsetX, curWinRect.y1 + pdWindows(i).offsetY, childRect.x2 - childRect.x1, childRect.y2 - childRect.y1, 1
                            Else
                            
                                'Force all child windows to the new size of the parent
                                getActualMainFormClientRect childRect
                                MoveWindow pdWindows(i).hWnd, childRect.x1, childRect.y1, childRect.x2 - childRect.x1, childRect.y2 - childRect.y1, 1
                            
                            End If
                        
                    End Select
                    
                    mainFormTriggeredChange = False
                    
                Next i
                
                'Remember the main window's location
                updateStoredWindowLocation mainWindowIndex
                                
            'The parent window has just been maximized, minimized, or restored
            Case WM_SYSCOMMAND
            
                'Per MSDN (http://msdn.microsoft.com/en-us/library/windows/desktop/ms646360%28v=vs.85%29.aspx):
                ' In WM_SYSCOMMAND messages, the four low-order bits of the wParam parameter are used internally by the system.
                ' To obtain the correct result when testing the value of wParam, an application must combine the value 0xFFF0 with
                ' the wParam value by using the bitwise AND operator.
                wParam = wParam And WM_SYSCOMMAND_WPARAM_MASK
            
                If (wParam = SC_MAXIMIZE) Or (wParam = SC_RESTORE) Then
                
                    'If the window was previously minimized, restore any child windows now
                    For i = 0 To numOfPDWindows - 1
                        If pdWindows(i).forciblyMinimized Then
                            pdWindows(i).forciblyMinimized = False
                            pdWindows(i).formReference.WindowState = vbNormal
                        End If
                    Next i
                
                    'Manually retrieve the new window position.
                    GetWindowRect lng_hWnd, curWinRect
                    
                    'Move each child window, while retaining the current offset
                    For i = 0 To numOfPDWindows - 1
                        
                        GetWindowRect pdWindows(i).hWnd, childRect
                        
                        Select Case pdWindows(i).typeOfWindow
                    
                            'Toolbars
                            Case TOOLBAR_WINDOW
                                If toolbarsFloating Then
                                    MoveWindow pdWindows(i).hWnd, curWinRect.x1 + pdWindows(i).offsetX, curWinRect.y1 + pdWindows(i).offsetY, childRect.x2 - childRect.x1, childRect.y2 - childRect.y1, 1
                                Else
                                    
                                    'Because the toolbars are not floating, they will have alignment data stored.  Use that to calculate their
                                    ' position instead.
                                    setWindowPosByAlignment i
                                    
                                    'MoveWindow pdWindows(i).hWnd, curWinRect.x1 + pdWindows(i).offsetX, curWinRect.y1 + pdWindows(i).offsetY, childRect.x2 - childRect.x1, childRect.y2 - childRect.y1, 1
                                    fixDockedToolbarSize i
                                End If
                        
                            'Images
                            Case IMAGE_WINDOW
                                If imagesFloating Then
                                    MoveWindow pdWindows(i).hWnd, curWinRect.x1 + pdWindows(i).offsetX, curWinRect.y1 + pdWindows(i).offsetY, childRect.x2 - childRect.x1, childRect.y2 - childRect.y1, 1
                                Else
                                    'Force the window to take up the full client area of the parent
                                    getActualMainFormClientRect childRect
                                    MoveWindow pdWindows(i).hWnd, childRect.x1, childRect.y1, childRect.x2 - childRect.x1, childRect.y2 - childRect.y1, 1
                                End If
                        
                        End Select
                        
                    Next i
                    
                'When minimizing, check all windows, and if one has not been minimized, minimize it now
                ElseIf wParam = SC_MINIMIZE Then
                
                    For i = 0 To numOfPDWindows - 1
                        If Not (pdWindows(i).formReference Is Nothing) Then
                            If pdWindows(i).formReference.WindowState <> vbMinimized Then
                                pdWindows(i).forciblyMinimized = True
                                pdWindows(i).formReference.WindowState = vbMinimized
                            End If
                        End If
                    Next i
                
                End If
                
            'Mouse has been moved.  If toolbars are floating, make them translucent.
            Case WM_MOUSEMOVE, WM_NCMOUSEMOVE
                For i = 0 To numOfPDWindows - 1
                    If (pdWindows(i).typeOfWindow = TOOLBAR_WINDOW) And (Not pdWindows(i).hasBeenMadeTransparent) And toolbarsFloating Then
                        pdWindows(i).hasBeenMadeTransparent = True
                        SetWindowTranslucency pdWindows(i).hWnd, INACTIVE_TOOLBOX_OPACITY
                    End If
                Next i
                
            'The user is switching to or from PD and another application
            Case WM_ACTIVATEAPP
            
                'wParam contains a notification of whether we are gaining or losing focus.
                Select Case wParam
                
                    'PD is losing focus
                    Case WA_INACTIVE
                        
                    'PD is gaining focus
                    Case Else
                        
                End Select
                        
        End Select
    
    End If




' *************************************************************
' C A U T I O N   C A U T I O N   C A U T I O N   C A U T I O N
' -------------------------------------------------------------
' DO NOT ADD ANY OTHER CODE BELOW THE "END SUB" STATEMENT BELOW
'   add this warning banner to the last routine in your class
' *************************************************************
End Sub


