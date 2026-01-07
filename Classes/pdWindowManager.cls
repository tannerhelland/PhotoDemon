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
'Copyright 2013-2026 by Tanner Helland
'Created: 02/October/13
'Last updated: 04/March/25
'Last update: improve saving/restoring window location for child windows
'
'PhotoDemon first began as a single-pane, single-image editor.  About a decade ago,
' I rewrote it as an MDI project, and in 2013 I rewrote it again - this time, to a
' tabbed interface in line with the UI trends of the era.
'
'This class came into existence during the 2013 rewrite, as a way to better manage
' complex window interactions.  It has been rewritten quite a few times over the years
' as new window management techniques have emerged, and at present, it's a thin shell
' of its original (very complex) form.
'
'In order to perform detailed window management, this class subclasses multiple forms
' and/or window messages.  I've made it as IDE-safe as I can, but breakpoints may cause
' problems.  Consider yourself warned.  (Also, some window order API calls do not
' function properly in the IDE per http://support.microsoft.com/kb/192254.
' Like everything else in PD, compile for best results.)
'
'Also, special thanks to VB coder Merri for a simple trick that enables Unicode window captions
' (http://www.vbforums.com/showthread.php?527802-VB6-UniCaption).
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Subclassed messages related to window movement
Private Const WM_MOVING As Long = &H216
Private Const WM_SIZING As Long = &H214
Private Const WM_GETMINMAXINFO As Long = &H24

'Non-subclassed messages that are used to retrieve Unicode window captions
Private Const WM_GETTEXT = &HD
Private Const WM_GETTEXTLENGTH = &HE
Private Const WM_SETTEXT = &HC

'Constants for changing extended window style
Private Const WS_EX_APPWINDOW As Long = &H40000
'Private Const WS_EX_LAYERED As Long = &H80000      'No longer used due to problems under Wine
Private Const WS_EX_COMPOSITED As Long = &H2000000
Private Const WS_CHILD As Long = &H40000000
Private Const WS_POPUP As Long = &H80000000
Private Const GWL_EXSTYLE As Long = (-20)
Private Const GWL_STYLE As Long = (-16)
Private Const SWP_NOACTIVATE As Long = &H10
Private Const SWP_NOMOVE As Long = &H2
Private Const SWP_NOOWNERZORDER As Long = &H200
Private Const SWP_NOSENDCHANGING As Long = &H400
Private Const SWP_NOZORDER As Long = &H4
Private Const HWND_TOP As Long = 0

'Main window tracking
Private m_MainWindowHWnd As Long

'These constants can be used as the second parameter of the ShowWindow API function
Private Enum ShowWindowOptions
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

Private Declare Function DefWindowProcW Lib "user32" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal bEnable As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, ByRef lpRect As winRect) As Long
Private Declare Function GetFocus Lib "user32" () As Long
Private Declare Function GetParent Lib "user32" (ByVal targetHWnd As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal targetHWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, ByVal lpRect As Long) As Long
Private Declare Function InvalidateRect Lib "user32" (ByVal targetHWnd As Long, ByRef lpRect As Any, ByVal bErase As Long) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function MapWindowPoints Lib "user32" (ByVal hWndFrom As Long, ByVal hWndTo As Long, ByVal ptrToPointList As Long, ByVal numPoints As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function SetActiveWindow Lib "user32" (ByVal targetHWnd As Long) As Long
Private Declare Function SetFocus Lib "user32" (ByVal targetHWnd As Long) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal targetHWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Sub SetWindowPos Lib "user32" (ByVal targetHWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As ShowWindowOptions) As Long
Private Declare Function UpdateWindow Lib "user32" (ByVal targetHWnd As Long) As Long

'Misc window APIs.  Note that these are modern APIs with strict version requirements;
' see individual invocations for details.
Private Declare Function DwmIsCompositionEnabled Lib "dwmapi" (ByRef dstEnabledBool As Long) As Long
Private Declare Function DwmSetWindowAttribute Lib "dwmapi" (ByVal hWnd As Long, ByVal dwAttribute As Long, ByVal pvAttribute As Long, ByVal cbAttribute As Long) As Long

'We manually enforce a specific minimum width/height for the main window
Private Const PD_MAIN_WINDOW_MINIMUM_HEIGHT As Long = 640
Private Const PD_MAIN_WINDOW_MINIMUM_WIDTH As Long = 880
Private Const PD_MAIN_WINDOW_PREFERRED_HEIGHT As Long = 720
Private Const PD_MAIN_WINDOW_PREFERRED_WIDTH As Long = 980

'This class *performs subclassing*.  Be cautious in the IDE.
Implements ISubclass

'XML handling (used to save/load window locations) is handled through a specialized class
Private m_XML As pdXML
Private m_WindowPresetPath As String

'To improve performance (and code organization), dense forms in PD are typically handle as a single
' mostly-blank "parent" form, while child subpanels within that form are all separate forms.
' At run-time, we dynamically assign those child panels to their parent, manipulating their window
' bits as we do so, ensuring that only panels we actually need get loaded and displayed.
'
'Child forms that have been turned into "panels" must be tracked so that we can manually restore
' their original window bits prior to unloading.  (Otherwise, VB gets confused and can become
' unstable.  This dictionary tracks original window bits and restores them whenever it receives
' a "deactivate" notification.
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

Private Type MinMaxInfo
    Reserved As POINTL
    MaxSize As POINTL
    maxPosition As POINTL
    MinTrackSize As POINTL
    MaxTrackSize As POINTL
End Type

Private m_tmpMinMax As MinMaxInfo

'***********************************************************************************************
'  GENERIC HELPER FUNCTIONS: VB-friendly wrappers to window-related APIs
'***********************************************************************************************
Friend Sub BringWindowToForeground(ByVal dstHWnd As Long)
    SetForegroundWindow dstHWnd
End Sub

Friend Sub SetOSDarkTheme(ByVal dstHWnd As Long)
    If OS.IsWin10OrLater() Then
        Const DWMWA_USE_IMMERSIVE_DARK_MODE = 20
        If (OS.GetWin10Build() >= 22000) Then
            Dim srcValue As Long: srcValue = 1
            DwmSetWindowAttribute dstHWnd, DWMWA_USE_IMMERSIVE_DARK_MODE, VarPtr(srcValue), 4&
        End If
    End If
End Sub

'The VB-specific "Form.Caption = <string>" statement isn't Unicode-compatible.  It also has the quirk of causing a docked child form
' to un-dock.  Thus, all caption requests must be passed through the window manager, who will then use the API to change caption text.
Friend Sub SetWindowCaptionW(ByVal targetHWnd As Long, ByRef newCaption As String)
    DefWindowProcW targetHWnd, WM_SETTEXT, 0&, ByVal StrPtr(newCaption)
End Sub

Friend Function GetWindowCaptionW(ByVal targetHWnd As Long) As String
    
    'Note that we can't use GetWindowTextW directly, because VB will intercept the message and return an
    ' ANSI-ified string.  Instead, we have to go directly to the default wndProc handler.
    Dim sizeString As Long
    sizeString = DefWindowProcW(targetHWnd, WM_GETTEXTLENGTH, 0&, 0&)
    
    If (sizeString > 0) Then
        
        GetWindowCaptionW = String$(sizeString, 0)
        
        '+1 is required for the terminating null-char; see https://msdn.microsoft.com/en-us/library/windows/desktop/ms632627(v=vs.85).aspx
        Dim apiReturn As Long
        apiReturn = DefWindowProcW(targetHWnd, WM_GETTEXT, sizeString + 1, StrPtr(GetWindowCaptionW))
        If (apiReturn <> sizeString) Then GetWindowCaptionW = Strings.TrimNull(GetWindowCaptionW)
        
    Else
        GetWindowCaptionW = vbNullString
    End If
    
End Function

Friend Function GetVisibilityByHWnd(ByVal srcHWnd As Long) As Boolean
    GetVisibilityByHWnd = (IsWindowVisible(srcHWnd) <> 0)
End Function

Friend Sub SetVisibilityByHWnd(ByVal srcHWnd As Long, ByVal visibilityState As Boolean, Optional ByVal activateWindowToo As Boolean = False)
    If visibilityState Then
        If activateWindowToo Then ShowWindow srcHWnd, SW_RESTORE Else ShowWindow srcHWnd, SW_SHOWNA
    Else
        ShowWindow srcHWnd, SW_HIDE
    End If
End Sub

Friend Sub SetEnablementByHWnd(ByVal dstHWnd As Long, ByVal enabledState As Boolean)
    EnableWindow dstHWnd, IIf(enabledState, 1&, 0&)
End Sub

Friend Sub SetSizeByHWnd(ByVal dstHWnd As Long, ByVal newWidth As Long, ByVal newHeight As Long, Optional ByVal notifyWindow As Boolean = True)
    Dim swpFlags As Long
    swpFlags = SWP_NOMOVE Or SWP_NOZORDER Or SWP_NOACTIVATE Or SWP_NOOWNERZORDER
    If (Not notifyWindow) Then swpFlags = swpFlags Or SWP_NOSENDCHANGING
    SetWindowPos dstHWnd, 0&, 0&, 0&, newWidth, newHeight, swpFlags
End Sub

'Thin wrapper to SetWindowPos, which means left/top coordinates are in *client coordinates*.
Friend Sub SetSizeAndPositionByHWnd(ByVal dstHWnd As Long, ByVal newLeft As Long, ByVal newTop As Long, ByVal newWidth As Long, ByVal newHeight As Long, Optional ByVal notifyWindow As Boolean = True)
    Dim swpFlags As Long
    swpFlags = SWP_NOZORDER Or SWP_NOACTIVATE Or SWP_NOOWNERZORDER
    If (Not notifyWindow) Then swpFlags = swpFlags Or SWP_NOSENDCHANGING
    SetWindowPos dstHWnd, 0&, newLeft, newTop, newWidth, newHeight, swpFlags
End Sub

Friend Function GetWindowLongWrapper(ByVal dstHWnd As Long, Optional ByVal useExtendedStyle As Boolean = False) As Long
    Dim nIndex As Long
    If useExtendedStyle Then nIndex = GWL_EXSTYLE Else nIndex = GWL_STYLE
    GetWindowLongWrapper = GetWindowLong(dstHWnd, nIndex)
End Function

'As it says, a simplified wrapper to SetWindowLong.  Returns the original window style before any changes are made.
Friend Function SetWindowLongWrapper(ByVal dstHWnd As Long, ByVal newFlag As Long, Optional ByVal removeFlagInstead As Boolean = False, Optional ByVal useExtendedStyle As Boolean = False, Optional ByVal overrideEntireLong As Boolean = False) As Long
    
    Dim nIndex As Long
    If useExtendedStyle Then nIndex = GWL_EXSTYLE Else nIndex = GWL_STYLE
    
    Dim curStyle As Long
    curStyle = GetWindowLong(dstHWnd, nIndex)
    
    If overrideEntireLong Then
        SetWindowLong dstHWnd, nIndex, newFlag
    Else
        If removeFlagInstead Then
            SetWindowLong dstHWnd, nIndex, curStyle And (Not newFlag)
        Else
            SetWindowLong dstHWnd, nIndex, curStyle Or newFlag
        End If
    End If
    
    SetWindowLongWrapper = curStyle
    
End Function

Friend Sub SetWindowPos_API(ByVal targetHWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
    SetWindowPos targetHWnd, hWndInsertAfter, x, y, cx, cy, wFlags
End Sub

Friend Function GetWindowHeight(ByVal srcHWnd As Long) As Long
    Dim tmpRect As winRect
    GetWindowRect srcHWnd, VarPtr(tmpRect)
    GetWindowHeight = (tmpRect.y2 - tmpRect.y1)
End Function

Friend Function GetWindowWidth(ByVal srcHWnd As Long) As Long
    Dim tmpRect As winRect
    GetWindowRect srcHWnd, VarPtr(tmpRect)
    GetWindowWidth = (tmpRect.x2 - tmpRect.x1)
End Function

Friend Sub GetWindowRect_API(ByVal srcHWnd As Long, ByRef dstRect As winRect)
    GetWindowRect srcHWnd, VarPtr(dstRect)
End Sub

Friend Sub GetWindowRect_API_Universal(ByVal srcHWnd As Long, ByVal ptrToAnyRectObject As Long)
    GetWindowRect srcHWnd, ptrToAnyRectObject
End Sub

Friend Sub GetClientWinRect(ByVal srcHWnd As Long, ByRef dstRect As winRect)
    GetClientRect srcHWnd, dstRect
End Sub

Friend Function GetClientWidth(ByVal targetHWnd As Long) As Long
    Dim tmpRect As winRect
    GetClientRect targetHWnd, tmpRect
    GetClientWidth = tmpRect.x2
End Function

Friend Function GetClientHeight(ByVal targetHWnd As Long) As Long
    Dim tmpRect As winRect
    GetClientRect targetHWnd, tmpRect
    GetClientHeight = tmpRect.y2
End Function

Friend Function GetFocusAPI() As Long
    GetFocusAPI = GetFocus()
End Function

Friend Sub SetFocusAPI(ByVal hWndToReceiveFocus As Long)
    SetFocus hWndToReceiveFocus
End Sub

Friend Function GetClientToScreen(ByVal srcHWnd As Long, ByRef srcPoint As PointAPI) As Boolean
    
    'MapWindowPoints is preferred over ClientToScreen because it works on BiDi systems.  However,
    ' MapWindowPoints can return 0 as a valid output.  You can work around this using SetLastError
    ' (although I don't know how well that would work in VB, given the way Err.LastDllError works),
    ' but we simply assume success and always return TRUE.
    MapWindowPoints srcHWnd, 0, VarPtr(srcPoint), 1&
    GetClientToScreen = True
    
End Function

Friend Function GetClientToScreen_Universal(ByVal srcHWnd As Long, ByVal ptrPtAPI As Long) As Boolean
    
    'MapWindowPoints is preferred over ClientToScreen because it works on BiDi systems.  However,
    ' MapWindowPoints can return 0 as a valid output.  You can work around this using SetLastError
    ' (although I don't know how well that would work in VB, given the way Err.LastDllError works),
    ' but we simply assume success and always return TRUE.
    MapWindowPoints srcHWnd, 0, ptrPtAPI, 1&
    GetClientToScreen_Universal = True
    
End Function

Friend Function GetScreenToClient(ByVal srcHWnd As Long, ByRef srcPoint As PointAPI) As Boolean
    
    'See the notes for GetClientToScreen, above
    MapWindowPoints 0, srcHWnd, VarPtr(srcPoint), 1&
    GetScreenToClient = True
    
End Function

Friend Sub ActivateWindowAPI(ByVal hWndToActivate As Long)
    SetActiveWindow hWndToActivate
End Sub

Friend Function IsDWMCompositionEnabled() As Boolean
    If OS.IsVistaOrLater Then
        If OS.IsWin8OrLater Then
            IsDWMCompositionEnabled = True
        Else
            Dim tmpLong As Long
            DwmIsCompositionEnabled tmpLong
            IsDWMCompositionEnabled = (tmpLong <> 0)
        End If
    Else
        IsDWMCompositionEnabled = False
    End If
End Function

'***********************************************************************************************
'  (/End GENERIC HELPER FUNCTIONS)
'***********************************************************************************************

'Various window manager settings require us to re-align the primary canvas.  (e.g. showing or hiding a toolbar window
' changes the available canvas area)  Normally, the window manager will automatically request a refresh of the canvas
' when one of these actions occurs, but if the caller knows that multiple refresh-triggering-actions will be happening
' in a group (like when PD is first loaded, and all toolbars are loaded and positioned in turn), it can deactivate
' those auto-refreshes for a meaningful performance boost.
Friend Sub SetAutoRefreshMode(ByVal newMode As Boolean)
    m_AutoRefreshMode = newMode
End Sub

Friend Function GetAutoRefreshMode() As Boolean
    GetAutoRefreshMode = m_AutoRefreshMode
End Function

'Activate a new child options panel for some arbitrary parent window.  The child window will be sized to
' auto-fit the parent, unless you also pass dimensions and/or position.
Friend Sub ActivateToolPanel(ByVal formHWndToActivate As Long, ByVal parentWindowHWnd As Long, Optional ByVal hwndToDeactivate As Long = 0&, Optional ByVal pxLeft As Long = 0&, Optional ByVal pxTop As Long = 1&, Optional ByVal pxWidth As Long = 0&, Optional ByVal pxHeight As Long = 0&)
    
    'Initialize the backup window setting dictionary, as necessary.  We use this to store
    ' a list of which tool panels have been activated this session.  This allows us to swap
    ' between previously-used panels more quickly than loading+unloading from scratch.
    If (m_ToolPanelDictionary Is Nothing) Then Set m_ToolPanelDictionary = New pdDictionary
    
    'In PD, only one tool panel can be active at a time, and we always enforce a strict
    ' "free the previous tool panel before activating this one" behavior.
    If (hwndToDeactivate <> 0) Then Me.DeactivateToolPanel hwndToDeactivate
    
    'As a failsafe, make sure all passed hWnds are valid
    If (formHWndToActivate <> 0) And (parentWindowHWnd <> 0) Then
        
        'As a failsafe, ensure the window hasn't already been made a child of the tool panel owner.
        If (GetParent(formHWndToActivate) <> parentWindowHWnd) Then
            
            'Cache the current window style settings.  We will restore these before unloading
            ' the window, to prevent issues with VB's unload process.
            If (Not m_ToolPanelDictionary.DoesKeyExist(formHWndToActivate)) Then m_ToolPanelDictionary.AddEntry formHWndToActivate, GetWindowLong(formHWndToActivate, GWL_STYLE)
            
            'Convert the window to a child window and assign it to the options panel.
            ' Note that the window bit reassignments come directly from MSDN:
            ' https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-setparent
            SetWindowLong formHWndToActivate, GWL_STYLE, GetWindowLong(formHWndToActivate, GWL_STYLE) Or WS_CHILD
            SetWindowLong formHWndToActivate, GWL_STYLE, GetWindowLong(formHWndToActivate, GWL_STYLE) And (Not WS_POPUP)
            SetParent formHWndToActivate, parentWindowHWnd
            
        End If
            
        'Next, get the client dimensions of the options window
        Dim parentPanelRect As winRect
        GetClientRect parentWindowHWnd, parentPanelRect
        
        'Next, move the child form into position
        If (pxWidth = 0&) Then pxWidth = (parentPanelRect.x2 - parentPanelRect.x1) - pxLeft
        If (pxHeight = 0&) Then pxHeight = (parentPanelRect.y2 - parentPanelRect.y1) - pxTop
        SetWindowPos formHWndToActivate, HWND_TOP, pxLeft, pxTop, pxWidth, pxHeight, SWP_NOACTIVATE Or SWP_NOOWNERZORDER Or SWP_NOZORDER
        
        'Finally, set visibility according to the parent window
        Me.SetVisibilityByHWnd formHWndToActivate, Me.GetVisibilityByHWnd(parentWindowHWnd), False
        
    End If
    
End Sub

'After a tool panel has been loaded, it must be specially deactivated.
' This is necessary because we dynamically mess with its window bits as part of "syncing"
' it's size and position to a specific on-screen region.
'
'Note that at shutdown time, the shutdown code passes the specific hWnd it wants us to use; this isn't
Friend Sub DeactivateToolPanel(ByVal hwndToDeactivate As Long)
    
    'Ignore this step entirely if no tool panel is currently loaded
    If (hwndToDeactivate <> 0) Then
        
        'Before doing anything else, hide the window in question - this prevents flicker
        ' as window settings are changed.
        Me.SetVisibilityByHWnd hwndToDeactivate, False
        
        'Remove the current parent (if any); this prevents VB from getting confused since we didn't
        ' assign the parent through internal VB6 commands
        If (GetParent(hwndToDeactivate) <> 0) Then SetParent hwndToDeactivate, 0
        
        'Restore the original window style bits (which we likely changed as part of making the
        ' toolpanel a child window)
        If m_ToolPanelDictionary.DoesKeyExist(hwndToDeactivate) Then
            SetWindowLong hwndToDeactivate, GWL_STYLE, m_ToolPanelDictionary.GetEntry_Long(hwndToDeactivate, GetWindowLong(hwndToDeactivate, GWL_STYLE))
            m_ToolPanelDictionary.DeleteEntry hwndToDeactivate
        End If
        
    End If
    
End Sub

'Force a full window refresh (invalidate + update)
Friend Sub ForceWindowRepaint(ByVal hWnd As Long)
    InvalidateRect hWnd, ByVal 0&, 0&
    UpdateWindow hWnd
End Sub

'By default, tool windows do not show up in the taskbar or in the Alt+Tab rotation.
' Thanks to excellent debug work by ChenLin, this is a problem on older OSes that
' only display window icons in the Alt+Tab rotation.  We can work around this problem by
' manually setting the WS_EX_APPWINDOW bit after a window is created, but before the window
' is shown.
Friend Sub ForceWindowAppearInAltTab(ByVal srcHWnd As Long, ByVal appearState As Boolean)
    
    Dim curExStyle As Long
    curExStyle = GetWindowLong(srcHWnd, GWL_EXSTYLE)
    
    If appearState Then
        curExStyle = curExStyle Or WS_EX_APPWINDOW
    Else
        curExStyle = curExStyle And (Not WS_EX_APPWINDOW)
    End If
    
    SetWindowLong srcHWnd, GWL_EXSTYLE, curExStyle
    
End Sub

'The first time PhotoDemon is run, this sub can be called to nicely center the window on the user's primary monitor.
Friend Sub SetFirstRunMainWindowPosition()

    'Start by retrieving the primary monitor's dimensions.  Note that this check relies on access to PD's g_Displays class.
    Dim tmpPrimaryDisplay As pdDisplay
    Set tmpPrimaryDisplay = g_Displays.PrimaryDisplay
    
    Dim pWorkingRect As RectL
    If (Not tmpPrimaryDisplay Is Nothing) Then
        tmpPrimaryDisplay.GetWorkingRect pWorkingRect
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
    If (idealWidth < PD_MAIN_WINDOW_MINIMUM_WIDTH) Then idealWidth = PD_MAIN_WINDOW_MINIMUM_WIDTH
    If (idealHeight < PD_MAIN_WINDOW_MINIMUM_HEIGHT) Then idealHeight = PD_MAIN_WINDOW_MINIMUM_HEIGHT
    
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
    
        MoveWindow m_MainWindowHWnd, .x1, .y1, .x2 - .x1, .y2 - .y1, 0
    End With

End Sub

'If a window has a previous position stored, this will return TRUE
Friend Function IsPreviousPositionStored(ByRef frmReference As Form) As Boolean

    'Start by looking for this form's location data in the XML engine.
    Dim windowName As String
    windowName = m_XML.GetXMLSafeTagName(frmReference.Name)
    
    'If an entry is found, restore the window to that location.
    Dim tagPos As Long
    IsPreviousPositionStored = m_XML.DoesTagExist("WindowEntry", "id", windowName, tagPos)
    
End Function

'If a window had location data previously stored, this function will retrieve that data and move the window into place.
'Returns: TRUE if a previous location was stored; FALSE otherwise.
Friend Function RestoreWindowLocation(ByRef frmReference As Form) As Boolean

    RestoreWindowLocation = False

    'Start by looking for this form's location data in the XML engine.
    Dim windowName As String
    windowName = m_XML.GetXMLSafeTagName(frmReference.Name)
    
    'If an entry is found, restore the window to that location.
    Dim tagPos As Long
    If m_XML.DoesTagExist("WindowEntry", "id", windowName, tagPos) Then
    
        'Retrieve this window's location data from the XML file.
        
        'Window rect
        Dim tmpRect As winRect
        With tmpRect
            .x1 = m_XML.GetUniqueTag_Long("WindowLeft", 0, tagPos)
            .y1 = m_XML.GetUniqueTag_Long("WindowTop", 0, tagPos)
            .x2 = m_XML.GetUniqueTag_Long("WindowRight", 0, tagPos)
            .y2 = m_XML.GetUniqueTag_Long("WindowBottom", 0, tagPos)
        End With
        
        'Window state (so that maximized windows are handled correctly if the primary display has changed resolution)
        frmReference.WindowState = m_XML.GetUniqueTag_Long("WindowState", vbNormal, tagPos)
        
        'Window DPI (if DPI changes between sessions, we need to re-map all coordinates accordingly)
        Dim tmpDPIModifier As Single
        tmpDPIModifier = m_XML.GetUniqueTag_Double("WindowDPI", 1#, tagPos)
        
        'For now, we limit DPI coverage to 400% or less.  Future coverage may be expanded.
        If (tmpDPIModifier < 1#) Then tmpDPIModifier = 1#
        If (tmpDPIModifier > 4#) Then tmpDPIModifier = 4#
        
        'Calculate a difference between the old and new DPIs, and apply that modifier to our saved values, as necessary
        tmpDPIModifier = Interface.GetSystemDPIRatio() / tmpDPIModifier
        
        If (tmpDPIModifier <> 1!) Then
            With tmpRect
                .x1 = .x1 * tmpDPIModifier
                .y1 = .y1 * tmpDPIModifier
                .x2 = .x2 * tmpDPIModifier
                .y2 = .y2 * tmpDPIModifier
            End With
        End If
        
        'Make sure the location values will result in an on-screen form.
        ' If they will not (for example, if the user detached a secondary monitor on which PhotoDemon was being used),
        ' change the values to ensure this window appears on-screen.
        
        'Note that this check relies on access to PD's g_Displays class, which returns the full
        ' virtual desktop dimensions, and not just the primary monitor's (as VB's Screen object does).
        If (frmReference.WindowState <> vbMaximized) Then
        
            Dim winWidth As Long, winHeight As Long
            With tmpRect
                winWidth = .x2 - .x1
                winHeight = .y2 - .y1
            End With
            
            'If the loaded positions are valid, restore them now.
            If ((winWidth > 0) And (winHeight > 0)) Then
                
                'So far, everything has been done in screen coordinates.  However, everything in PD
                ' (except for FormMain) is a child window of FormMain, and MoveWindow uses client coordinates.
                ' As such, let's translate coordinates as necessary.
                Dim winIsTopLevel As Boolean
                winIsTopLevel = Strings.StringsEqual(frmReference.Name, "FormMain", True)
                
                'If this *is* a top-level window, translate coordinates to be relative to the parent window.
                If (Not winIsTopLevel) Then
                    
                    Dim rectMain As winRect
                    GetWindowRect FormMain.hWnd, VarPtr(rectMain)
                    With tmpRect
                        winWidth = .x2 - .x1
                        winHeight = .y2 - .y1
                        .x1 = .x1 + rectMain.x1
                        .x2 = .x1 + winWidth
                        .y1 = .y1 + rectMain.y1
                        .y2 = .y1 + winHeight
                    End With
                    
                    'We now have the correct coordinates, *in screen coordinate space*.
                    
                    'Perform one final check for on-screen-ness.
                    EnsureWindowRectIsOnScreen tmpRect
                    
                    'Translate those to the target hWnd space, then move the window into place!
                    MapWindowPoints 0&, GetParent(frmReference.hWnd), VarPtr(tmpRect), 2
                    With tmpRect
                        MoveWindow frmReference.hWnd, .x1, .y1, .x2 - .x1, .y2 - .y1, 1
                    End With
                    
                'If this is FormMain, restore it immediately without modification
                Else
                    
                    '(Okay, one modification - ensure the window is on-screen!)
                    EnsureWindowRectIsOnScreen tmpRect
        
                    With tmpRect
                        MoveWindow frmReference.hWnd, .x1, .y1, .x2 - .x1, .y2 - .y1, 1
                    End With
                    
                End If
                
            End If
            
        End If
        
        RestoreWindowLocation = True
        
    End If

End Function

'Ensure a given window rectangle is on-screen.  The window will be forcibly moved in-bounds, as necessary.
Private Sub EnsureWindowRectIsOnScreen(ByRef srcWinRect As winRect)
    
    'Failsafe only
    If (g_Displays Is Nothing) Then Exit Sub
    
    Dim winWidth As Long, winHeight As Long
    With srcWinRect
        
        'Calculate window width/height in advance (window rects use corner coordinates)
        winWidth = .x2 - .x1
        winHeight = .y2 - .y1
        
        'First ensure left/right window boundaries are on-screen
        If (.x1 < g_Displays.GetDesktopLeft) Then
            .x1 = g_Displays.GetDesktopLeft
            .x2 = .x1 + winWidth
        ElseIf ((.x1 + winWidth) > (g_Displays.GetDesktopLeft + g_Displays.GetDesktopWidth)) Then
            .x1 = (g_Displays.GetDesktopLeft + g_Displays.GetDesktopWidth) - winWidth
            If (.x1 < g_Displays.GetDesktopLeft) Then .x1 = g_Displays.GetDesktopLeft
            .x2 = .x1 + winWidth
        End If
        
        'Having done that, shrink the window as necessary to ensure it's in-bounds
        If (.x2 > (g_Displays.GetDesktopLeft + g_Displays.GetDesktopWidth)) Then .x2 = (g_Displays.GetDesktopLeft + g_Displays.GetDesktopWidth)
        
        'Repeat the above steps in the y-direction
        If (.y1 < g_Displays.GetDesktopTop) Then
            .y1 = g_Displays.GetDesktopTop
            .y2 = .y1 + winHeight
        ElseIf ((.y1 + winHeight) > (g_Displays.GetDesktopTop + g_Displays.GetDesktopHeight)) Then
            .y1 = (g_Displays.GetDesktopTop + g_Displays.GetDesktopHeight) - winHeight
            If (.y1 < g_Displays.GetDesktopTop) Then .y1 = g_Displays.GetDesktopTop
            .y2 = .y1 + winWidth
        End If
        
        If (.y2 > (g_Displays.GetDesktopTop + g_Displays.GetDesktopHeight)) Then .y2 = (g_Displays.GetDesktopTop + g_Displays.GetDesktopHeight)
        
    End With
    
End Sub

'Load previous window locations from file.
Friend Function LoadAllWindowLocations() As Boolean
    
    'Attempt to load and validate the relevant preset file; if we can't, create a new, blank XML object
    Dim windowLoadSuccessful As Boolean
    If Files.FileExists(m_WindowPresetPath) Then
        windowLoadSuccessful = m_XML.LoadXMLFile(m_WindowPresetPath)
        If windowLoadSuccessful Then windowLoadSuccessful = m_XML.IsPDDataType("Window locations")
    End If
    
    If (Not windowLoadSuccessful) Then
        PDDebug.LogAction "No window location data found.  A new window location file has been created."
        ResetXMLData
    End If
    
    'We don't actually load window locations now.  Now that the XML data is safely inside our XML engine,
    ' we load window data from it on-demand as windows are registered with the window manager.

End Function

'Write the current locations of all windows to the XML engine.
' (These will be used to restore the window location on subsequent loads.)
Friend Function SaveWindowLocation(ByRef frmReference As Form, Optional ByVal writeLocationsToFile As Boolean = False) As Boolean
    
    'Generate a name to be used in the settings file
    Dim windowName As String
    windowName = m_XML.GetXMLSafeTagName(frmReference.Name)
    
    'Check for this window preset in the file.  If it does not exist, add its section now
    If (Not m_XML.DoesTagExist("WindowEntry", "id", windowName)) Then
        m_XML.WriteTagWithAttribute "WindowEntry", "id", windowName, vbNullString, True
        m_XML.CloseTag "WindowEntry"
        m_XML.WriteBlankLine
    End If
    
    'Now that the section is guaranteed to exist, generate a window rect for this window and save it accordingly
    Dim tmpRect As winRect
    GetWindowRect frmReference.hWnd, VarPtr(tmpRect)
    
    'GetWindowRect always returns screen coordinates.  This is fine for e.g. the main window (FormMain),
    ' but for any child windows, we want coordinates *relative to FormMain*.
    Dim winIsTopLevel As Boolean
    winIsTopLevel = Strings.StringsEqual(frmReference.Name, "FormMain", True)
    
    'If this is *not* a top-level window, translate coordinates to be relative to FormMain.
    If (Not winIsTopLevel) Then
        
        Dim rectMain As winRect, winWidth As Long, winHeight As Long
        GetWindowRect FormMain.hWnd, VarPtr(rectMain)
        With tmpRect
            winWidth = .x2 - .x1
            winHeight = .y2 - .y1
            .x1 = .x1 - rectMain.x1
            .x2 = .x1 + winWidth
            .y1 = .y1 - rectMain.y1
            .y2 = .y1 + winHeight
        End With
        
    End If
    
    With m_XML
        .UpdateTag "WindowLeft", tmpRect.x1, "WindowEntry", "id", windowName
        .UpdateTag "WindowTop", tmpRect.y1, "WindowEntry", "id", windowName
        .UpdateTag "WindowRight", tmpRect.x2, "WindowEntry", "id", windowName
        .UpdateTag "WindowBottom", tmpRect.y2, "WindowEntry", "id", windowName
        .UpdateTag "WindowState", frmReference.WindowState, "WindowEntry", "id", windowName
        .UpdateTag "WindowDPI", Interface.GetSystemDPIRatio, "WindowEntry", "id", windowName
    End With
    
    'Write the data out to file
    If writeLocationsToFile Then m_XML.WriteXMLToFile m_WindowPresetPath

End Function

'Reset the XML engine.  Note that the XML object SHOULD ALREADY BE INSTANTIATED before calling this function.
Private Sub ResetXMLData()
    m_XML.PrepareNewXML "Window locations"
    m_XML.WriteBlankLine
    m_XML.WriteComment "Everything past this point is window location data for various PhotoDemon dialogs."
    m_XML.WriteBlankLine
End Sub

'The primary PhotoDemon form must register its hWnd, so we can track its movement and move any children windows accordingly.
Friend Sub RegisterMainForm(ByRef parentForm As Form)
    
    'Add this window to the collection, and cache its hWnd (because we reference the main window handle frequently)
    m_MainWindowHWnd = parentForm.hWnd
    
    If PDMain.IsProgramRunning() Then
    
        'Apply any unique styles to the parent window.  (PD uses complex nested windows for toolbars,
        ' and setting the WS_EX_COMPOSITED bit reduces flickering on older versions of Windows.)
        Dim newWinStyle As Long
        newWinStyle = GetWindowLong(parentForm.hWnd, GWL_EXSTYLE) Or WS_EX_COMPOSITED
        SetWindowLong m_MainWindowHWnd, GWL_EXSTYLE, newWinStyle
        
        'Enforce minimum size handling.  (Note that this will subclass the main window, which allows us to handle
        ' other useful window messages simultaneously.)
        Me.RequestMinMaxTracking m_MainWindowHWnd, , PD_MAIN_WINDOW_MINIMUM_WIDTH, PD_MAIN_WINDOW_MINIMUM_HEIGHT
        
        'Look for previous location data in the window location file.  If said data exists, load it and move the window to that location.
        If (Not RestoreWindowLocation(parentForm)) Then Me.SetFirstRunMainWindowPosition
        
    End If
        
End Sub

'When a window is unloaded, call this function so that we can stop subclassing in a safe and predictable way, and also track the
' last-known location of this window.
Friend Sub UnregisterMainForm(ByRef srcForm As Form)
    
    'Terminate main form subclassing
    m_MainWindowHWnd = 0
    StopMaxMinSubclassing srcForm.hWnd
    
    'Store the current main window position in the "last window position" collection and copy the collection out to file
    SaveWindowLocation srcForm, True

End Sub

Private Sub Class_Initialize()

    'Reset all tracking variables
    m_numOfMinMaxEntries = 0
    ReDim m_MinMaxEntries(0 To 3) As MinMaxTracker
    
    m_AutoRefreshMode = True
    
    'Prepare the XML handler, and retrieve window location data from file (if it exists)
    Set m_XML = New pdXML
    m_XML.SetTextCompareMode vbBinaryCompare
    m_WindowPresetPath = UserPrefs.GetPresetPath & "Program_WindowLocations.xml"
    LoadAllWindowLocations
    
End Sub

Private Sub Class_Terminate()
    
    'If individual objects have requested max/min window size tracking, free their subclassers now
    Dim i As Long
    If (m_numOfMinMaxEntries > 0) Then
        For i = 0 To m_numOfMinMaxEntries - 1
            StopMaxMinSubclassing m_MinMaxEntries(i).hWnd
        Next i
    End If
    
End Sub

'Use this function to set minimum and/or maximum sizes for any arbitrary window.  (Note that this requires us to subclass
' the window in question; plan accordingly!)
Friend Sub RequestMinMaxTracking(ByVal srcHWnd As Long, Optional ByVal internalID As Long, Optional ByVal minWidth As Long = 0, Optional ByVal minHeight As Long = 0, Optional ByVal maxWidth As Long = 0, Optional ByVal maxHeight As Long = 0)

    If ((srcHWnd <> 0) And PDMain.IsProgramRunning()) Then
    
        With m_MinMaxEntries(m_numOfMinMaxEntries)
            .hWnd = srcHWnd
            .internalID = internalID
            .minWidth = minWidth
            .minHeight = minHeight
            .maxWidth = maxWidth
            .maxHeight = maxHeight
        End With
        
        m_numOfMinMaxEntries = m_numOfMinMaxEntries + 1
        If (m_numOfMinMaxEntries > UBound(m_MinMaxEntries)) Then ReDim Preserve m_MinMaxEntries(0 To m_numOfMinMaxEntries * 2 - 1) As MinMaxTracker
        
        'Unfortunately, there's no simple way to set min/max values in advance.  Instead, we must subclass the window, and respond
        ' to any received WM_GETMINMAXINFO messages.
        VBHacks.StartSubclassing srcHWnd, Me
        
    End If

End Sub

Private Sub StopMaxMinSubclassing(ByVal hWnd As Long)
    
    If ((hWnd <> 0) And (m_numOfMinMaxEntries <> 0)) Then
        
        Dim i As Long
        For i = 0 To m_numOfMinMaxEntries - 1
            If (m_MinMaxEntries(i).hWnd = hWnd) Then
                VBHacks.StopSubclassing hWnd, Me
                m_MinMaxEntries(i).hWnd = 0
                Exit For
            End If
        Next i
        
    End If
    
End Sub

Private Function HandleMinMaxMessage(ByVal hWnd As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal dwRefData As Long) As Boolean
    
    HandleMinMaxMessage = False
    
    'lParam contains a pointer to the window's rect location.  Retrieve it now.
    If (lParam <> 0) Then CopyMemoryStrict VarPtr(m_tmpMinMax), lParam, LenB(m_tmpMinMax)
    
    'Iterate through all the windows we're handling min/max data for, and when we find the correct one, supply its min/max
    ' values to the parent handler.
    If (m_numOfMinMaxEntries > 0) Then
    
        Dim i As Long
        For i = 0 To m_numOfMinMaxEntries - 1
            If (m_MinMaxEntries(i).hWnd = hWnd) Then
                
                With m_MinMaxEntries(i)
                    
                    'Only request minimum (or maximum) tracking if the caller supplied at least *1* valid value for that dimension
                    If (.minWidth <> 0) Or (.minHeight <> 0) Then
                        If (.minWidth <> 0) Then m_tmpMinMax.MinTrackSize.x = .minWidth
                        If (.minHeight <> 0) Then m_tmpMinMax.MinTrackSize.y = .minHeight
                    End If
    
                    If (.maxWidth <> 0) Or (.maxHeight <> 0) Then
                        If (.maxWidth <> 0) Then m_tmpMinMax.MaxTrackSize.x = .maxWidth Else m_tmpMinMax.MaxTrackSize.x = Me.GetClientWidth(m_MainWindowHWnd)
                        If (.maxHeight <> 0) Then m_tmpMinMax.MaxTrackSize.y = .maxHeight Else m_tmpMinMax.MaxTrackSize.y = Me.GetClientHeight(m_MainWindowHWnd)
                    End If
                    
                End With
                
                CopyMemoryStrict lParam, VarPtr(m_tmpMinMax), LenB(m_tmpMinMax)
                HandleMinMaxMessage = True
                Exit For
                
            End If
        Next i
        
    End If
        
End Function

Private Function ISubclass_WindowMsg(ByVal hWnd As Long, ByVal uiMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal dwRefData As Long) As Long
    
    'Min/max handling is handled for multiple potential hWnds
    If (uiMsg = WM_GETMINMAXINFO) Then
        If HandleMinMaxMessage(hWnd, wParam, lParam, dwRefData) Then
            ISubclass_WindowMsg = 0
        Else
            ISubclass_WindowMsg = VBHacks.DefaultSubclassProc(hWnd, uiMsg, wParam, lParam)
        End If
    
    'Always tear down subclassing manually when a subclassed window is being destroyed
    ElseIf (uiMsg = WM_NCDESTROY) Then
        StopMaxMinSubclassing hWnd
        ISubclass_WindowMsg = VBHacks.DefaultSubclassProc(hWnd, uiMsg, wParam, lParam)
    
    Else
        
        'If this is the main window hWnd, update color-management settings on move and/or size events (as the parent monitor
        ' may have changed)
        If (hWnd = m_MainWindowHWnd) Then
            If (uiMsg = WM_MOVING) Or (uiMsg = WM_SIZING) Then ColorManagement.CheckParentMonitor
        End If
        
        ISubclass_WindowMsg = VBHacks.DefaultSubclassProc(hWnd, uiMsg, wParam, lParam)
        
    End If

End Function

