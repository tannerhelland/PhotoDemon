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
'Copyright 2013-2016 by Tanner Helland
'Created: 02/October/13
'Last updated: 05/March/06
'Last update: purge all image tabstrip code, as it's now handled by pdCanvas and pdImageStrip
'
'PhotoDemon first began as a single-pane, single-image editor.  About a decade ago, I rewrote it as an MDI project, and in
' 2013, it became time to rewrite its window method once again - this time, to a tabbed interface in line with the major
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
'In order to perform all this detailed window management, this class subclasses a lot of forms and a lot of messages.  That
' said, it should be wholly IDE-safe, but consider yourself warned.  (Also, some window order API calls do not function
' properly in the IDE per http://support.microsoft.com/kb/192254.  Compile for best results.)
'
'Also, special thanks to VB coder Merri for a simple trick that enables Unicode window captions
' (http://www.vbforums.com/showthread.php?527802-VB6-UniCaption).
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Subclassed messages related to window movement
Private Const WM_SYSCOMMAND As Long = &H112
Private Const WM_SYSCOMMAND_WPARAM_MASK As Long = &HFFF0&
Private Const WM_MOVE As Long = &H3
Private Const WM_MOVING As Long = &H216
Private Const WM_MOUSELEAVE As Long = &H2A3
Private Const WM_SIZING As Long = &H214
Private Const WM_SIZE As Long = &H5
Private Const WM_ERASEBKGND As Long = &H14
Private Const WM_SHOWWINDOW As Long = &H18
Private Const WM_WINDOWPOSCHANGED As Long = &H47
Private Const WM_NCACTIVATE As Long = &H86
Private Const WM_GETMINMAXINFO As Long = &H24
Private Const WM_CANCELMODE As Long = &H1F
Private Const WM_ENABLE As Long = &HA

'Subclassed messages for detecting when PD is losing focus to another app, or regaining it from another app
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

'Each window that the manager tracks stores some basic metadata in a variable of this type.
Private Type pdWindowData
    FormReference As Form
    hWnd As Long
    LastPosition As winRect
    lastWindowState As Long
    lastWindowBorder As Long
    minimumWidth As Long
    minimumHeight As Long
    offsetX As Long
    offsetY As Long
    isVisible As Boolean
    subclassingReleased As Boolean
    windowName As String
End Type

'When a window is being resized, it is helpful to know which edge is being resized.  Windows uses the following system, which we mimic:
Private Enum WinResizeEdge
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
Private Declare Function ShowWindow Lib "user32" (ByVal hndWindow As Long, ByVal nCmdShow As showWindowOptions) As Long

'APIs for getting/setting Unicode window captions
Private Declare Function DefWindowProcW Lib "user32" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Const WM_GETTEXT = &HD
Private Const WM_GETTEXTLENGTH = &HE
Private Const WM_SETTEXT = &HC

'APIs for changing window style
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal targetHwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal targetHwnd As Long, ByVal nIndex As Long) As Long
Private Declare Sub SetWindowPos Lib "user32" (ByVal targetHwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Declare Function GetParent Lib "user32" (ByVal targetHwnd As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function EndDeferWindowPos Lib "user32" (ByVal hWinPosInfo As Long) As Long
Private Declare Function InvalidateRect Lib "user32" (ByVal targetHwnd As Long, ByRef lpRect As Any, ByVal bErase As Long) As Long
Private Declare Function UpdateWindow Lib "user32" (ByVal targetHwnd As Long) As Long
Private Declare Function AnimateWindow Lib "user32" (ByVal targetHwnd As Long, ByVal dwTime As Long, ByVal dwFlags As Long) As Long
Private Declare Function GetFocus Lib "user32" () As Long
Private Declare Function SetFocus Lib "user32" (ByVal targetHwnd As Long) As Long
Private Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, ByVal lprcUpdateRectPtr As Long, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long

'RedrawWindow API constants
Private Const RDW_INVALIDATE As Long = &H1
Private Const RDW_INTERNALPAINT As Long = &H2
Private Const RDW_ERASE As Long = &H4
Private Const RDW_VALIDATE As Long = &H8
Private Const RDW_NOINTERNALPAINT As Long = &H10
Private Const RDW_NOERASE As Long = &H20
Private Const RDW_NOCHILDREN As Long = &H40
Private Const RDW_ALLCHILDREN As Long = &H80
Private Const RDW_UPDATENOW As Long = &H100
Private Const RDW_ERASENOW As Long = &H200
Private Const RDW_FRAME As Long = &H400
Private Const RDW_NOFRAME As Long = &H800

'Enforce a specific minimum width/height for the main window
Private Const PD_MAIN_WINDOW_MINIMUM_HEIGHT As Long = 640
Private Const PD_MAIN_WINDOW_MINIMUM_WIDTH As Long = 880

Private Const PD_MAIN_WINDOW_PREFERRED_HEIGHT As Long = 720
Private Const PD_MAIN_WINDOW_PREFERRED_WIDTH As Long = 980

'Child window tracking
Private pdWindows() As pdWindowData
Private numOfPDWindows As Long

'Primary window tracking; while the main window's data is stored just like the data for any other window, it is helpful to maintain dedicated
' links to its hWnd and position in the tracking array, to prevent us having to constantly search for it.
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

'When the main form is being moved or resized, it will trigger the movement of all child forms to match.  We set this variable to TRUE
' so that child forms don't waste time and energy handling window move messages that have already been handled.
Private mainFormTriggeredChange As Boolean

'When the main form is minimized, a flurry of window messages are sent to all child windows, including instructions to resize to (0, 0).
' To prevent this size from being synched as the "last-known window size", we track minimize mode, and disregard all location and
' size syncing while it is active.
Private insideMinimizeMode As Boolean

'The tool options panel is a special case.  It may have a child panel active attached to it, depending on the tool.  We track that panel's
' hWnd here.  (It is generated by the "activateToolPanel" sub.)
Private m_CurrentToolPanelHwnd As Long
Private m_ToolPanelDictionary As pdDictionary

'If m_AutoRefreshMode is TRUE, the window manager will forcibly re-align the main window's canvas area to match any
' changes to window settings.
Private m_AutoRefreshMode As Boolean

'As a convenience, this class can manage min/max behavior for individual windows.  Windows (the OS) makes this unpleasantly complicated,
' as you can't set a static value up-front; instead, you must respond to a dedicated window message, which is repeatedly sent whenever
' a window attempts a resize.  *sigh*
Private Type MinMaxTracker
    hWnd As Long
    internalID As Long
    minWidth As Long
    minHeight As Long
    maxWidth As Long
    maxHeight As Long
End Type

Private m_MinMaxEntries() As MinMaxTracker
Private m_numOfMinMaxEntries As Long

'This is the actual API struct used by WM_GETMINMAXINFO
Private Type POINTL
    x As Long
    y As Long
End Type

Private Type MINMAXINFO
    Reserved As POINTL
    MaxSize As POINTL
    maxPosition As POINTL
    MinTrackSize As POINTL
    MaxTrackSize As POINTL
End Type

Private m_tmpMinMax As MINMAXINFO

'***********************************************************************************************
'  GENERIC HELPER FUNCTIONS: VB-friendly wrappers to window-related APIs
'***********************************************************************************************

'The VB-specific "Form.Caption = <string>" statement isn't Unicode-compatible.  It also has the quirk of causing a docked child form
' to un-dock.  Thus, all caption requests must be passed through the window manager, who will then use the API to change caption text.
Public Sub SetWindowCaptionW(ByRef targetHwnd As Long, ByRef newCaption As String)
    DefWindowProcW targetHwnd, WM_SETTEXT, 0&, ByVal StrPtr(newCaption)
End Sub

Public Sub SetVisibilityByHWnd(ByVal srcHwnd As Long, ByVal visibilityState As Boolean, Optional ByVal activateWindowToo As Boolean = False)
    If visibilityState Then
        If activateWindowToo Then ShowWindow srcHwnd, SW_RESTORE Else ShowWindow srcHwnd, SW_SHOWNA
    Else
        ShowWindow srcHwnd, SW_HIDE
    End If
End Sub

Public Sub SetEnablementByHWnd(ByVal dstHwnd As Long, ByVal enabledState As Boolean)
    EnableWindow dstHwnd, IIf(enabledState, 1&, 0&)
End Sub

Friend Sub GetClientWinRect(ByVal srcHwnd As Long, ByRef dstRect As winRect)
    GetClientRect srcHwnd, dstRect
End Sub

Public Function GetClientWidth(ByVal targetHwnd As Long) As Long
    Dim tmpRect As winRect
    GetClientRect targetHwnd, tmpRect
    GetClientWidth = tmpRect.x2
End Function

Public Function GetClientHeight(ByVal targetHwnd As Long) As Long
    Dim tmpRect As winRect
    GetClientRect targetHwnd, tmpRect
    GetClientHeight = tmpRect.y2
End Function

Private Function GetWindowWidth(ByVal targetHwnd As Long) As Long
    Dim tmpRect As winRect
    GetWindowRect targetHwnd, tmpRect
    GetWindowWidth = tmpRect.x2 - tmpRect.x1
End Function

Private Function GetWindowHeight(ByVal targetHwnd As Long) As Long
    Dim tmpRect As winRect
    GetWindowRect targetHwnd, tmpRect
    GetWindowHeight = tmpRect.y2 - tmpRect.y1
End Function

Public Function GetFocusAPI() As Long
    GetFocusAPI = GetFocus()
End Function

Public Sub SetFocusAPI(ByVal hWndToReceiveFocus As Long)
    SetFocus hWndToReceiveFocus
End Sub

'***********************************************************************************************
'  (/End GENERIC HELPER FUNCTIONS)
'***********************************************************************************************

'Various window manager settings require us to re-align the primary canvas.  (e.g. showing or hiding a toolbar window
' changes the available canvas area)  Normally, the window manager will automatically request a refresh of the canvas
' when one of these actions occurs, but if the caller knows that multiple refresh-triggering-actions will be happening
' in a group (like when PD is first loaded, and all toolbars are loaded and positioned in turn), it can deactivate
' those auto-refreshes for a meaningful performance boost.
Public Sub SetAutoRefreshMode(ByVal newMode As Boolean)
    m_AutoRefreshMode = newMode
End Sub

Public Function GetAutoRefreshMode() As Boolean
    GetAutoRefreshMode = m_AutoRefreshMode
End Function

'Activate a new tool options panel
Public Sub ActivateToolPanel(ByVal formHWndToActivate As Long, ByVal optionWindowHwnd As Long)
    
    'Cache the hWnd of this form
    m_CurrentToolPanelHwnd = formHWndToActivate
    
    If (m_CurrentToolPanelHwnd <> 0) And (optionWindowHwnd <> 0) Then
        
        'Initialize the backup window setting dictionary, as necessary.  We use this to store a list of which tool panels have been
        ' activated this session.  This allows us to swap between previously-used panels more quickly than loading+unloading from scratch.
        If (m_ToolPanelDictionary Is Nothing) Then Set m_ToolPanelDictionary = New pdDictionary
        
        'First, make the form a child of the options panel form
        If GetParent(m_CurrentToolPanelHwnd) <> optionWindowHwnd Then
        
            'Cache the current window style settings.  We will restore these at exit time, to prevent issues with VB's unload process
            m_ToolPanelDictionary.AddEntry m_CurrentToolPanelHwnd, GetWindowLong(m_CurrentToolPanelHwnd, GWL_STYLE)
            
            'Convert the window to a child window and assign it to the options panel
            SetWindowLong m_CurrentToolPanelHwnd, GWL_STYLE, GetWindowLong(m_CurrentToolPanelHwnd, GWL_STYLE) Or CLng(WS_CHILD)
            SetWindowLong m_CurrentToolPanelHwnd, GWL_STYLE, GetWindowLong(m_CurrentToolPanelHwnd, GWL_STYLE) And CLng(Not WS_POPUP)
            SetParent m_CurrentToolPanelHwnd, optionWindowHwnd
            
        End If
        
        'Next, get the client dimensions of the options window
        Dim optionsPanelRect As winRect
        GetClientRect optionWindowHwnd, optionsPanelRect
        
        'Next, move the child form into position
        SetWindowPos m_CurrentToolPanelHwnd, HWND_TOP, 0, 1, optionsPanelRect.x2 - optionsPanelRect.x1, optionsPanelRect.y2 - optionsPanelRect.y1, SWP_NOACTIVATE Or SWP_NOOWNERZORDER Or SWP_NOZORDER
        
        'Finally, set visibility according to the parent window
        Me.SetVisibilityByHWnd m_CurrentToolPanelHwnd, toolbar_Options.Visible, False
        
    End If
    
End Sub

'After a tool panel has been loaded, it must be specially deactivated.  This is necessary because we dynamically mess with its window bits.
Public Sub DeactivateToolPanel(ByVal programIsGoingDown As Boolean, Optional ByVal useCustomHwnd As Long = 0)
    
    'By default, this function will automatically apply the new settings to the last tool options panel it created.
    ' However, the caller can specify a custom hWnd if they so desire.  (PD uses this when the program is closing,
    ' so the window manager can reset all necessary window bits, allowing VB to close the window normally.)
    Dim targetHwnd As Long
    If useCustomHwnd = 0 Then targetHwnd = m_CurrentToolPanelHwnd Else targetHwnd = useCustomHwnd
    
    'Release the parent reference for the previous activated tool panel
    If (targetHwnd <> 0) Then
        
        If programIsGoingDown Then
            
            If GetParent(targetHwnd) <> 0 Then
                
                'This is a dirty cheat, but basically, this function restores the initial WindowLong GWL_STYLE flags the window had
                ' when it was first passed to the window manager.  This lets us reset the window state to whatever VB's defaults are,
                ' which reduces the risk for error as we shut the program down.
                SetWindowLong targetHwnd, GWL_STYLE, m_ToolPanelDictionary.GetEntry_Long(targetHwnd, GetWindowLong(targetHwnd, GWL_STYLE))
                
            End If
            
        End If
        
        Me.SetVisibilityByHWnd targetHwnd, False
        
    End If
    
    m_CurrentToolPanelHwnd = 0
    
End Sub

'If external (or internal) functions move windows around on their own, they can manually request resyncing the window's offset.
Public Sub RequestWindowResync(ByVal winIndex As Long)
    
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
    UpdateStoredWindowLocation winIndex
    
End Sub

'Force a full window refresh (invalidate + update)
Private Sub ForceFullWindowRefresh(ByVal winIndex As Long)
    InvalidateRect pdWindows(winIndex).hWnd, ByVal 0&, 0
    UpdateWindow pdWindows(winIndex).hWnd
End Sub

Public Sub ForceWindowRepaint(ByVal hWnd As Long)
    InvalidateRect hWnd, ByVal 0&, 0&
    UpdateWindow hWnd
End Sub

'By default, tool windows do not show up in the taskbar or in the Alt+Tab rotation.  Thanks to excellent debug work by Chen Lin,
' this is a problem on older OSes that only display window icons in the Alt+Tab rotation.  We can work around this problem by
' manually setting the WS_EX_APPWINDOW bit after a window is created, but before the window is shown.
Public Sub ForceWindowAppearInAltTab(ByVal srcHwnd As Long, ByVal appearState As Boolean)
    
    Dim curExStyle As Long
    curExStyle = GetWindowLong(srcHwnd, GWL_EXSTYLE)
    
    If appearState Then
        curExStyle = curExStyle Or WS_EX_APPWINDOW
    Else
        curExStyle = curExStyle And (Not WS_EX_APPWINDOW)
    End If
    
    SetWindowLong srcHwnd, GWL_EXSTYLE, curExStyle
    
End Sub

'The first time PhotoDemon is run, this sub can be called to nicely center the window on the user's primary monitor.
Public Sub SetFirstRunMainWindowPosition()

    'Start by retrieving the primary monitor's dimensions.  Note that this check relies on access to PD's g_Displays class.
    Dim PrimaryDisplay As pdDisplay
    Set PrimaryDisplay = g_Displays.PrimaryDisplay
    
    Dim pWorkingRect As RECTL
    If Not (PrimaryDisplay Is Nothing) Then
        PrimaryDisplay.GetWorkingRect pWorkingRect
    Else
        With pWorkingRect
            .Left = 0
            .Top = 0
            .Right = Screen.Width / TwipsPerPixelXFix
            .Bottom = Screen.Height / TwipsPerPixelYFix
        End With
    End If
    
    Dim primaryMonitorRect As winRect
    With pWorkingRect
        primaryMonitorRect.x1 = .Left
        primaryMonitorRect.x2 = .Right
        primaryMonitorRect.y1 = .Top
        primaryMonitorRect.y2 = .Bottom
    End With
    
    'Using the primary monitor's dimensions, construct a new rect that fills most of (but not all) the user's screen
    Dim monitorWidth As Long, monitorHeight As Long
    monitorWidth = (primaryMonitorRect.x2 - primaryMonitorRect.x1)
    monitorHeight = (primaryMonitorRect.y2 - primaryMonitorRect.y1)
    
    Dim idealWidth As Long, idealHeight As Long
    idealWidth = monitorWidth * 0.85
    idealHeight = monitorHeight * 0.9
    
    'Make sure the newly calculated "ideal" dimensions aren't less than PD's default width/height
    If idealWidth < PD_MAIN_WINDOW_MINIMUM_WIDTH Then idealWidth = PD_MAIN_WINDOW_MINIMUM_WIDTH
    If idealHeight < PD_MAIN_WINDOW_MINIMUM_HEIGHT Then idealHeight = PD_MAIN_WINDOW_MINIMUM_HEIGHT
    
    'If PD's preferred minimum width/height is available, and it is larger than the currently calculated ideal width, use those instead.
    ' This is most relevant on 1024x768 monitors, where the "ideal" size of 80% of the monitor's available width and height is
    ' still ridiculously small (relative to PD's UI layout).
    If (PD_MAIN_WINDOW_PREFERRED_WIDTH < monitorWidth) And (idealWidth < PD_MAIN_WINDOW_PREFERRED_WIDTH) Then idealWidth = PD_MAIN_WINDOW_PREFERRED_WIDTH
    If (PD_MAIN_WINDOW_PREFERRED_HEIGHT < monitorHeight) And (idealHeight < PD_MAIN_WINDOW_PREFERRED_HEIGHT) Then idealHeight = PD_MAIN_WINDOW_PREFERRED_HEIGHT
    
    'Apply the new rect to the image
    Dim newWindowRect As winRect
    With newWindowRect
        .x1 = primaryMonitorRect.x1 + (monitorWidth - idealWidth) \ 2
        .x2 = .x1 + idealWidth
        .y1 = primaryMonitorRect.y1 + (monitorHeight - idealHeight) \ 2
        .y2 = .y1 + idealHeight
    
        MoveWindow mainWindowHwnd, .x1, .y1, .x2 - .x1, .y2 - .y1, 0
    End With

End Sub

'If a window had location data previously stored, this function will retrieve that data and move the window into place.
' If window data cannot be found, the form will be moved to the requested pseudo-location supplied by the second parameter.
Public Function RestoreWindowLocation(ByVal windowIndex As Long, ByVal preferredLocationIfNecessary As Long) As Boolean

    RestoreWindowLocation = False

    'Start by looking for this form's location data in the XML engine.
    
    'If an entry is found, restore the window to that location.
    If xmlEngine.DoesTagExist("windowEntry", "id", pdWindows(windowIndex).windowName) Then
    
        'Retrieve this window's location data from the XML file.
        With pdWindows(windowIndex)
        
            'Window rect values
            .LastPosition.x1 = CLng(xmlEngine.getUniqueTag_String("windowLeft", 0, , "windowEntry", "id", .windowName))
            .LastPosition.y1 = CLng(xmlEngine.getUniqueTag_String("windowTop", 0, , "windowEntry", "id", .windowName))
            .LastPosition.x2 = CLng(xmlEngine.getUniqueTag_String("windowRight", 0, , "windowEntry", "id", .windowName))
            .LastPosition.y2 = CLng(xmlEngine.getUniqueTag_String("windowBottom", 0, , "windowEntry", "id", .windowName))
            
            'Window and border state
            .lastWindowState = CLng(xmlEngine.getUniqueTag_String("windowState", 0, , "windowEntry", "id", .windowName))
            .lastWindowBorder = CLng(xmlEngine.getUniqueTag_String("windowStyle", 0, , "windowEntry", "id", .windowName))
            
            'Apply those values to the window, starting with border style, then state, then coordinates
            .FormReference.BorderStyle = .lastWindowBorder
            .FormReference.WindowState = .lastWindowState
            
            'Make sure the location values will result in an on-screen form.  If they will not (for example, if the user detached a
            ' secondary monitor on which PhotoDemon was being used), change the values to ensure this window appears on-screen.
            
            'Note that this check relies on access to PD's g_Displays class, which returns the full virtual desktop dimensions, and
            ' not just the primary monitor's (as VB's Screen object does).
            Dim winWidth As Long, winHeight As Long
            winWidth = .LastPosition.x2 - .LastPosition.x1
            winHeight = .LastPosition.y2 - .LastPosition.y1
            
            If (.LastPosition.x1 + winWidth) < g_Displays.GetDesktopLeft Then
                .LastPosition.x1 = g_Displays.GetDesktopLeft
                .LastPosition.x2 = .LastPosition.x1 + winWidth
            End If
            
            If .LastPosition.x1 > (g_Displays.GetDesktopLeft + g_Displays.GetDesktopWidth) Then
                .LastPosition.x1 = (g_Displays.GetDesktopWidth - winWidth)
                .LastPosition.x2 = .LastPosition.x1 + winWidth
            End If
            
            If .LastPosition.y1 < g_Displays.GetDesktopTop Then
                .LastPosition.y1 = g_Displays.GetDesktopTop
                .LastPosition.y2 = .LastPosition.y1 + winHeight
            End If
            
            If .LastPosition.y1 > g_Displays.GetDesktopHeight Then
                .LastPosition.y1 = (g_Displays.GetDesktopHeight - winHeight)
                .LastPosition.y2 = .LastPosition.y1 + winHeight
            End If
            
            'If the loaded positions are valid, restore them now
            If (.LastPosition.x2 > .LastPosition.x1) And (.LastPosition.y2 > .LastPosition.y1) Then
                MoveWindow .hWnd, .LastPosition.x1, .LastPosition.y1, .LastPosition.x2 - .LastPosition.x1, .LastPosition.y2 - .LastPosition.y1, 1
            End If
            
        End With
        
        RestoreWindowLocation = True
        
    End If

End Function

'Load previous window locations from file.
Public Function LoadAllWindowLocations() As Boolean
    
    'Attempt to load and validate the relevant preset file; if we can't, create a new, blank XML object
    If (Not xmlEngine.LoadXMLFile(windowDataPath)) Or Not (xmlEngine.isPDDataType("Window locations")) Then
        Debug.Print "No window location data found.  A new window location file has been created."
        ResetXMLData
    End If
    
    'We don't actually load window locations now.  Now that the XML data is safely inside our XML engine, we load window data from it
    ' on-demand as windows are added to the window manager.

End Function

'Write the current locations of all windows to the XML engine.  (These will be used to restore the window location on subsequent loads.)
Public Function SaveAllWindowLocations() As Boolean
    
    Dim i As Long
    For i = 0 To numOfPDWindows - 1
    
        'Check for this window preset in the file.  If it does not exist, add it now.
        If Not xmlEngine.DoesTagExist("windowEntry", "id", pdWindows(i).windowName) Then
        
            xmlEngine.writeTagWithAttribute "windowEntry", "id", pdWindows(i).windowName, "", True
            xmlEngine.closeTag "windowEntry"
            xmlEngine.writeBlankLine
        
        End If
        
        'Write this window's location data into the XML file.
        With pdWindows(i)
        
            'Window rect values
            xmlEngine.updateTag "windowLeft", .LastPosition.x1, "windowEntry", "id", .windowName
            xmlEngine.updateTag "windowTop", .LastPosition.y1, "windowEntry", "id", .windowName
            xmlEngine.updateTag "windowRight", .LastPosition.x2, "windowEntry", "id", .windowName
            xmlEngine.updateTag "windowBottom", .LastPosition.y2, "windowEntry", "id", .windowName
            
            'Window and border state
            xmlEngine.updateTag "windowState", .lastWindowState, "windowEntry", "id", .windowName
            xmlEngine.updateTag "windowStyle", .lastWindowBorder, "windowEntry", "id", .windowName
            
        End With
        
    Next i
    
    'Write the data out to file
    xmlEngine.writeXMLToFile windowDataPath

End Function

'Reset the XML engine.  Note that the XML object SHOULD ALREADY BE INSTANTIATED before calling this function.
Private Function ResetXMLData()
    xmlEngine.prepareNewXML "Window locations"
    xmlEngine.writeBlankLine
    xmlEngine.writeComment "Everything past this point is window location data for various PhotoDemon dialogs."
    xmlEngine.writeBlankLine
End Function

'The window manager keeps a persistent list of last-known locations for every window it accesses.  Window locations are updated when:
' - They're first loaded (and no previous position has been stored)
' - They're moved
' - They're closed
'Any of those actions just need to send a request to this sub, which will handle the actual location storage process.
Private Sub UpdateStoredWindowLocation(ByVal indexInWindowArray As Long)

    Dim newWindowRect As winRect
    
    With pdWindows(indexInWindowArray)
    
        GetWindowRect .hWnd, newWindowRect
        .LastPosition = newWindowRect
        If Not (.FormReference Is Nothing) Then
            .lastWindowBorder = .FormReference.BorderStyle
            .lastWindowState = .FormReference.WindowState
        End If
    
    End With

End Sub

'The primary PhotoDemon form must register its hWnd, so we can track its movement and move any children windows accordingly.
Public Sub RegisterMainForm(ByRef parentForm As Form)

    'Has this window already been registered?  If so, don't re-register it.
    Dim i As Long
    For i = 0 To numOfPDWindows
        If pdWindows(i).hWnd = parentForm.hWnd Then
            Debug.Print "WARNING!  Why are you attempting to register the main window twice??"
            Exit Sub
        End If
    Next i
    
    'Add this window to the collection, and cache its hWnd (because we reference the main window handle frequently)
    Set pdWindows(numOfPDWindows).FormReference = parentForm
    pdWindows(numOfPDWindows).hWnd = parentForm.hWnd
    mainWindowHwnd = parentForm.hWnd
    mainWindowIndex = numOfPDWindows
    pdWindows(numOfPDWindows).windowName = parentForm.Name
    
    'Apply any unique styles to the parent window
    Dim newWinStyle As Long
    newWinStyle = GetWindowLong(parentForm.hWnd, GWL_EXSTYLE) Or WS_EX_COMPOSITED
    
    'On Vista+, layered windows consume a bit more resources but they play very nicely with DWM
    If g_IsVistaOrLater Then newWinStyle = newWinStyle Or WS_EX_LAYERED
    SetWindowLong parentForm.hWnd, GWL_EXSTYLE, newWinStyle
    
    'Subclass the main window
    cSubclass.ssc_Subclass mainWindowHwnd, , , Me
    cSubclass.ssc_AddMsg mainWindowHwnd, MSG_AFTER, WM_SIZE
    cSubclass.ssc_AddMsg mainWindowHwnd, MSG_BEFORE, WM_SIZING, WM_MOVING
    cSubclass.ssc_AddMsg mainWindowHwnd, MSG_BEFORE_AFTER, WM_SYSCOMMAND
    
    'Look for previous location data in the window location file.  If said data exists, load it and move the window to that location.
    RestoreWindowLocation numOfPDWindows, 0
    
    'Resize the window array for future additions
    numOfPDWindows = numOfPDWindows + 1
    ReDim Preserve pdWindows(0 To numOfPDWindows) As pdWindowData
        
End Sub

'When a window is unloaded, call this function so that we can stop subclassing in a safe and predictable way, and also track the
' last-known location of this window.
Public Sub UnregisterForm(ByRef srcForm As Form)

    'Find this window in the array.  (Note that if it cannot be found, nothing will happen.)
    Dim i As Long
    For i = 0 To numOfPDWindows
        If pdWindows(i).hWnd = srcForm.hWnd Then
            
            'Update the last-known window location before the form closes
            UpdateStoredWindowLocation i
            
            'Mark the window as invisible
            pdWindows(i).isVisible = False
            
            'Release this window's subclassing
            pdWindows(i).subclassingReleased = True
            cSubclass.ssc_UnSubclass pdWindows(i).hWnd
            pdWindows(i).hWnd = 0
            
            'Remove our reference to this form (if we don't, it will remain in memory)
            Set pdWindows(i).FormReference = Nothing
            
        End If
    Next i

End Sub

Private Sub Class_Initialize()

    'Reset all tracking variables
    numOfPDWindows = 0
    ReDim pdWindows(0) As pdWindowData
    
    m_numOfMinMaxEntries = 0
    ReDim m_MinMaxEntries(0 To 3) As MinMaxTracker
    
    m_AutoRefreshMode = True
    
    'Prepare the subclasser
    Set cSubclass = New cSelfSubHookCallback
    
    'Prepare the XML handler, and retrieve window location data from file (if it exists)
    Set xmlEngine = New pdXML
    windowDataPath = g_UserPreferences.GetPresetPath & "Program_WindowLocations.xml"
    LoadAllWindowLocations
    
    'To prevent circular redraws, we disable certain subclassing when we know we are responsible for triggering the window messages.  Initialize
    ' that tracking variable to FALSE.
    mainFormTriggeredChange = False
    
End Sub

Private Sub Class_Terminate()

    'Release all subclassing
    cSubclass.ssc_Terminate
        
    'Manually clear all form references
    Dim i As Long
    For i = 0 To UBound(pdWindows)
        Set pdWindows(i).FormReference = Nothing
    Next i

End Sub

'Use this function to set minimum and/or maximum sizes for any arbitrary window.  (Note that this requires us to subclass
' the window in question; plan accordingly!)
Friend Sub RequestMinMaxTracking(ByVal srcHwnd As Long, Optional ByVal internalID As Long, Optional ByVal minWidth As Long = 0, Optional ByVal minHeight As Long = 0, Optional ByVal maxWidth As Long = 0, Optional ByVal maxHeight As Long = 0)

    With m_MinMaxEntries(m_numOfMinMaxEntries)
        .hWnd = srcHwnd
        .internalID = internalID
        .minWidth = minWidth
        .minHeight = minHeight
        .maxWidth = maxWidth
        .maxHeight = maxHeight
    End With
    
    m_numOfMinMaxEntries = m_numOfMinMaxEntries + 1
    If (m_numOfMinMaxEntries > UBound(m_MinMaxEntries)) Then
        ReDim Preserve m_MinMaxEntries(0 To m_numOfMinMaxEntries * 2 - 1) As MinMaxTracker
    End If
    
    'Unfortunately, there's no simple way to set min/max values in advance.  Instead, we must subclass the window, and respond
    ' to any received WM_GETMINMAXINFO messages.
    cSubclass.ssc_Subclass srcHwnd, , 1, Me
    cSubclass.ssc_AddMsg srcHwnd, MSG_BEFORE, WM_GETMINMAXINFO

End Sub

'Whenever a size change occurs that requires multiple windows to be redrawn, this function can be called.  It will automatically handle redraw duties.
Private Sub RequestWindowRedraw(ByVal winIndex As Long, ByRef curWinRect As winRect)
    ForceFullWindowRefresh winIndex
    RequestWindowResync winIndex
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
    Dim userSizeTooSmall As Boolean
    Dim resizeHandle As WinResizeEdge
    Dim i As Long
    
    'The primary purpose of this function is managing the main form's behavior.  Handle its case first.
    If (lng_hWnd = mainWindowHwnd) Then
        
        'Check parent window messages
        Select Case uMsg
        
            'The parent window is currently being moved
            Case WM_MOVING, WM_MOVE
            
                'lParam contains a pointer to the window's rect location.  Retrieve it now.
                Call CopyMemory(curWinRect, ByVal lParam, Len(curWinRect))
                
                'Remember the main window's location
                UpdateStoredWindowLocation mainWindowIndex

                'Notify the color management engine that the main window may have changed monitors, so it may need to
                ' be redrawn with a new color profile.
                If g_OpenImageCount > 0 Then CheckParentMonitor
                
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
                    If (curWinRect.x2 - curWinRect.x1) < PD_MAIN_WINDOW_MINIMUM_WIDTH Then
                        userSizeTooSmall = True
                        If (resizeHandle = WMSZ_BOTTOMLEFT) Or (resizeHandle = WMSZ_LEFT) Or (resizeHandle = WMSZ_TOPLEFT) Then
                            curWinRect.x1 = curWinRect.x2 - PD_MAIN_WINDOW_MINIMUM_WIDTH
                        Else
                            curWinRect.x2 = curWinRect.x1 + PD_MAIN_WINDOW_MINIMUM_WIDTH
                        End If
                    End If

                    If (curWinRect.y2 - curWinRect.y1) < PD_MAIN_WINDOW_MINIMUM_HEIGHT Then
                        userSizeTooSmall = True
                        If (resizeHandle = WMSZ_TOPLEFT) Or (resizeHandle = WMSZ_TOP) Or (resizeHandle = WMSZ_TOPRIGHT) Then
                            curWinRect.y1 = curWinRect.y2 - PD_MAIN_WINDOW_MINIMUM_HEIGHT
                        Else
                            curWinRect.y2 = curWinRect.y1 + PD_MAIN_WINDOW_MINIMUM_HEIGHT
                        End If
                    End If

                    If userSizeTooSmall Then

                        'Copy the new rect into place
                        CopyMemory ByVal lParam, curWinRect, Len(curWinRect)

                        'Because subsequent functions need access to the new rect data, we cheat a bit here - we use SetWindowPos to immediately
                        ' update the main window's rect, but we combine flags so that no new window messages are sent (to prevent infinite recursion of
                        ' window changes causing more window changes).
                        SetWindowPos lng_hWnd, 0, curWinRect.x1, curWinRect.y1, curWinRect.x2 - curWinRect.x1, curWinRect.y2 - curWinRect.y1, SWP_NOACTIVATE Or SWP_NOZORDER Or SWP_NOSENDCHANGING

                    End If

                End If
                
                'Remember the main window's location
                UpdateStoredWindowLocation mainWindowIndex
                
                mainFormTriggeredChange = False
                               
                                
            'The parent window has just been maximized, minimized, or restored
            Case WM_SYSCOMMAND
                
                'Per MSDN (http://msdn.microsoft.com/en-us/library/windows/desktop/ms646360%28v=vs.85%29.aspx):
                ' In WM_SYSCOMMAND messages, the four low-order bits of the wParam parameter are used internally by the system.
                ' To obtain the correct result when testing the value of wParam, an application must combine the value 0xFFF0 with
                ' the wParam value by using the bitwise AND operator.
                wParam = wParam And WM_SYSCOMMAND_WPARAM_MASK
                
                'RANDOM NOTE: for reasons I don't fully understand, bBefore will not always evaluate properly if used as a standalone
                ' boolean value (e.g. If Not bBefore Then...).  It must be *explicitly* checked against a VB Boolean type in order
                ' to return correct values.
                
                'The OS will automatically handle window restoration for us, but we need to make sure to resync all window positions
                ' afterwards.  If we don't do this, PD's internal window position values will fall out of sync.
                If ((wParam = SC_MAXIMIZE) Or (wParam = SC_RESTORE) Or (wParam = SC_SIZE) Or (wParam = SC_MOVE)) And (bBefore = False) And (Not isPDDeactivated) Then

                    'Debug.Print "PD restoration code running..."

                    GetWindowRect mainWindowHwnd, curWinRect
                    For i = 0 To numOfPDWindows - 1
                        RequestWindowResync i
                        RequestWindowRedraw i, pdWindows(mainWindowIndex).LastPosition
                    Next i

                    If insideMinimizeMode Then insideMinimizeMode = False

                End If

                'When minimizing, check all windows, and if one has not been minimized, minimize it now
                If (wParam = SC_MINIMIZE) And (bBefore = True) Then

                    'Debug.Print "PD minimize code running..."

                    'As a failsafe, update each window's location before minimizing it
                    For i = 0 To numOfPDWindows - 1
                        RequestWindowResync i
                    Next i

                    insideMinimizeMode = True

                End If
                
            'The user is switching to or from PD and another application
            Case WM_NCACTIVATE  'WM_ACTIVATEAPP was originally used here, but it is an unreliable piece of garbage.  It is sent to all top-level
                                ' windows, even when they are not being activated.  Thank you to the nice MSDN commenter who suggested NCACTIVATE
                                ' as a more reliable alternative.
            
                If bBefore Then
            
                    'wParam contains a notification of whether we are gaining or losing focus.
                    Select Case wParam
                    
                        'PD is losing focus; make all top-most windows not topmost
                        Case WA_INACTIVE
                            Debug.Print "> PhotoDemon lost focus! <"
                            isPDDeactivated = True
                            
                        'PD is gaining focus
                        Case Else
                            Debug.Print "> PhotoDemon regained focus! <"
                            
                            'Reactivate window state
                            If isPDDeactivated Then isPDDeactivated = False
                            
                    End Select
                    
                Else
                
                    If (wParam <> WA_INACTIVE) And (Not isPDDeactivated) And (Not insideMinimizeMode) Then
                        
                        'Debug.Print "Running Show Desktop hack"
                        
                        'TODO: see if this is still necessary with v7.0's new lightweight toolbox manager
                        
                        'This rather nasty hack is required to work around the "Show Desktop" event.  For reasons known only
                        ' to MS engineers, the Show Desktop command does not send *ANY* WM_SYSCOMMAND messages to your window.
                        ' Thus, you have literally no way to track whether the option is used, except to track window focus.
                        ' Because Show Desktop causes visual corruption of child windows, we must force a redraw of all open
                        ' toolboxes; do this by manually sending a "restore" command to the main PD window.
                        SendMessage mainWindowHwnd, WM_SYSCOMMAND, SC_RESTORE, ByVal 0&
                        
                        'Manually refresh the main window.  (Note that curWinRect doesn't matter here.)
                        RequestWindowRedraw mainWindowIndex, curWinRect
                        
                    End If
                
                End If
                        
        End Select
    
    'If this is *not* the main window, start testing other options
    Else
        
        'Windows who have requested min/max tracking will show up with the corresponding message
        If (uMsg = WM_GETMINMAXINFO) Then
            
            'lParam contains a pointer to the window's rect location.  Retrieve it now.
            CopyMemory ByVal VarPtr(m_tmpMinMax), ByVal lParam, Len(m_tmpMinMax)
            
            For i = 0 To m_numOfMinMaxEntries - 1
                If (m_MinMaxEntries(i).hWnd = lng_hWnd) Then
                    
                    With m_MinMaxEntries(i)
                        
                        'Only request minimum (or maximum) tracking if the caller supplied at least *1* valid value for that dimension
                        If (.minWidth <> 0) Or (.minHeight <> 0) Then
                            If (.minWidth <> 0) Then m_tmpMinMax.MinTrackSize.x = .minWidth
                            If (.minHeight <> 0) Then m_tmpMinMax.MinTrackSize.y = .minHeight
                        End If

                        If (.maxWidth <> 0) Or (.maxHeight <> 0) Then
                            If (.maxWidth <> 0) Then m_tmpMinMax.MaxTrackSize.x = .maxWidth Else m_tmpMinMax.MaxTrackSize.x = Me.GetClientWidth(mainWindowHwnd)
                            If (.maxHeight <> 0) Then m_tmpMinMax.MaxTrackSize.y = .maxHeight Else m_tmpMinMax.MaxTrackSize.y = Me.GetClientHeight(mainWindowHwnd)
                        End If
                        
                    End With
                    
                    CopyMemory ByVal lParam, ByVal VarPtr(m_tmpMinMax), Len(m_tmpMinMax)
                    bHandled = True
                    lReturn = 0
                    Exit For
                    
                End If
            Next i
        
        End If
        
    End If


' *************************************************************
' C A U T I O N   C A U T I O N   C A U T I O N   C A U T I O N
' -------------------------------------------------------------
' DO NOT ADD ANY OTHER CODE BELOW THE "END SUB" STATEMENT BELOW
'   add this warning banner to the last routine in your class
' *************************************************************
End Sub

