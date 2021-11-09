VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdFlyout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Move/Size Tool Panel
'Copyright 2021-2021 by Tanner Helland
'Created: 01/November/21
'Last updated: 05/November/21
'Last update: initial build (and testing on the Move/Size panel) looks good!  Time to try another tool...
'
'Controller for flyout panels from PD's (many) toolboxes.
'
'Toolboxes were originally located at the bottom of the screen - a hasty decision made years ago,
' before PD had any on-canvas tools.  (Only selection tools existed back then.)  As new canvas tools
' were added, the old toolbox design grew more and more unwieldy, and as the need for complex tools
' increased, there simply wasn't any way to do it with the old design.
'
'So in 2021 I bit the bullet and redesigned all toolpanels completely.  Instead of a using a large
' fixed height for all panels, they were converted to a much shorter design with flyout panels that
' automatically engage if any of their constituent controls receive focus.  Tab design is also
' manually handled on these panels to ensure intuitive behavior.
'
'This new design required a new class to handle the details of flyout show/hide behavior, since it
' requires manually modifying a bunch of window bits (without using different forms for every
' subpanel - which would require hundreds of new forms!).
'
'Note that this class also provides some flyout helper functions (like redraw state trackers)
' which are *not* used by this class directly, but are used by toolbox instances to prevent
' things like recursive redraws when swapping between panels.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

Public Event FlyoutClosed(ByRef origTriggerObject As Control)

'Positioning the dynamically raised listview window is a bit hairy; we use APIs so we can
' position things correctly in the screen's coordinate space (even on high-DPI displays)
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal targetHWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal srcHwnd As Long, ByRef dstRectL As RectL) As Long
Private Declare Function InvalidateRect Lib "user32" (ByVal hWnd As Long, ByVal ptrToRect As Long, ByVal bErase As Long) As Long
Private Declare Function MapWindowPoints Lib "user32" (ByVal hWndFrom As Long, ByVal hWndTo As Long, ByVal ptrToPointList As Long, ByVal numPoints As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal targetHWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Sub SetWindowPos Lib "user32" (ByVal targetHWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)

Private Const WS_EX_TOOLWINDOW As Long = &H80&
Private Const WS_EX_WINDOWEDGE As Long = &H100&
Private Const WS_EX_TOPMOST As Long = &H8&
Private Const WS_EX_PALETTEWINDOW As Long = (WS_EX_WINDOWEDGE Or WS_EX_TOOLWINDOW Or WS_EX_TOPMOST)
Private Const SWP_SHOWWINDOW As Long = &H40
Private Const SWP_NOACTIVATE As Long = &H10
Private Const SWP_FRAMECHANGED As Long = &H20

'When a flyout is visible, this is set to TRUE.
Private m_flyoutVisible As Boolean

'Flyout and flyout parent hWnds need to be manually tracked
Private m_flyoutHWnd As Long, m_ParentHWnd As Long
Private m_flyoutRectCopy As RectL

'Window styles only need to be set once
Private m_WindowStyleHasBeenSet As Boolean

'Window styles are blindly cached before we modify them; this allows us to restore whatever was
' originally present when the flyout is released
Private m_OriginalWindowBits As Long, m_OriginalWindowBitsEx As Long

'Soft reference to the original trigger object.  We don't actually touch this, but we return it
' when the flyout closes (so the parent can reset the state of the object as necessary)
Private m_origTriggerObject As Control

'An optional ID value set by the caller; useful for querying the currently open flyout, if any
Private m_flyoutTrackerID As Long

'Another convenience function for owner panels.  They can set this flag prior to initiating a
' series of synchronizations, then query the state during synchronization to avoid infinite loops.
Private m_flyoutSyncState As Boolean

'If something forces us to release our subclass while in the midst of the subclass proc, we want to delay the request until
' the subclass exits.  If we don't do this, PD will crash.
Private m_InSubclassNow As Boolean, m_SubclassActive As Boolean
Private WithEvents m_SubclassReleaseTimer As pdTimer
Attribute m_SubclassReleaseTimer.VB_VarHelpID = -1

'When a flyout is raised, we subclass the parent control.  If it is moved or sized or clicked,
' we automatically unload the flyout.  (This workaround is necessary for modal dialogs,
' among other things.)
Private Const WM_ENTERSIZEMOVE As Long = &H231
Private Const WM_WINDOWPOSCHANGING As Long = &H46&
Implements ISubclass

Friend Sub ShowFlyout(ByRef ownerForm As Form, ByRef originatingObject As Control, ByRef flyoutPanel As pdContainer, Optional ByVal flyoutTrackerID As Long = 0, Optional ByVal xOffset As Long = 0&)
    
    On Error GoTo UnexpectedTrouble
    
    'Basic failsafe checks on the owner form; disabled forms can't raise flyouts, for example.
    If (Not ownerForm.Visible) Or (Not ownerForm.Enabled) Or (Not PDMain.IsProgramRunning()) Then Exit Sub
    
    'If this flyout is already visible, exit
    If (flyoutPanel.hWnd = m_flyoutHWnd) Then Exit Sub
    
    'Hide any existing flyoutpanels
    If m_flyoutVisible Then UserControls.HideOpenFlyouts 0&
    
    'Update the internal flyout ID; the caller can use this to query flyout state later
    m_flyoutTrackerID = flyoutTrackerID
    
    'We first want to retrieve the originating object's window coordinates *in the screen's coordinate space*.
    ' (We need this to know how to position the listbox element.)
    Dim origObjRect As RectL, ownerFormRect As RectL
    GetWindowRect originatingObject.hWnd, origObjRect
    GetWindowRect ownerForm.hWnd, ownerFormRect
    Set m_origTriggerObject = originatingObject
    
    'We now want to determine idealized coordinates for the flyout.  In the future, I think the caller
    ' should be able to suggest this (using up/down/right/left hints, for example).
    Dim flyoutRect As RectL
    
    'To construct this rect, we start by calculating the position using the window rect of the
    ' originating object.
    With flyoutRect
        .Left = origObjRect.Left + xOffset
        .Top = ownerFormRect.Bottom - 2
        .Right = .Left + flyoutPanel.GetWidth
        .Bottom = .Top + flyoutPanel.GetHeight
    End With
    
    'We now want to make sure the flyout won't lie off-screen.  Check each dimension in turn.
    ' TODO!
    
    'The flyout is now ready to go.  The first time we raise the flyout, we want to cache its current
    ' window longs as whatever VB has set.  (These need to be restored when we hide the flyout, or VB
    ' may crash.)
    m_flyoutHWnd = flyoutPanel.hWnd
    m_ParentHWnd = ownerForm.hWnd
    If (Not m_WindowStyleHasBeenSet) Then
        m_WindowStyleHasBeenSet = True
        m_OriginalWindowBits = g_WindowManager.GetWindowLongWrapper(m_flyoutHWnd)
        m_OriginalWindowBitsEx = g_WindowManager.GetWindowLongWrapper(m_flyoutHWnd, True)
    End If
    
    'Now we are ready to display the flyout.  Make it a child of the main parent window (not this one),
    ' and apply any other relevant window styles.  (Switching parents allows us to display the window
    ' outside the boundaries of this toolbox, but with the window still disappearing if events like PD
    ' losing focus occur.)
    SetParent m_flyoutHWnd, FormMain.hWnd
    Const WS_CHILD As Long = &H40000000
    Const WS_POPUP As Long = &H80000000
    Const GWL_STYLE As Long = (-16)
    
    'g_WindowManager.SetWindowLongWrapper m_FlyoutHWnd, WS_EX_PALETTEWINDOW, False, True
    SetWindowLong m_flyoutHWnd, GWL_STYLE, GetWindowLong(m_flyoutHWnd, GWL_STYLE) Or WS_CHILD
    SetWindowLong m_flyoutHWnd, GWL_STYLE, GetWindowLong(m_flyoutHWnd, GWL_STYLE) And (Not WS_POPUP)
    
    'We also need to cache the flyout rect's position; when the flyout is closed, we will
    ' manually invalidate windows beneath it (only on certain OS + theme combinations;
    ' Aero handles this correctly).
    With m_flyoutRectCopy
        .Left = flyoutRect.Left
        .Top = flyoutRect.Top
        .Right = flyoutRect.Right
        .Bottom = flyoutRect.Bottom
    End With
    
    'Translate the flyout position to main window coordinate space (child windows are always positioned
    ' relative to parent client area).
    MapWindowPoints GetDesktopWindow(), FormMain.hWnd, VarPtr(flyoutRect), 2
    
    'Now we can show the flyout; we also need to notify it of changed window style bits.
    With flyoutRect
        SetWindowPos m_flyoutHWnd, 0&, .Left, .Top, .Right - .Left, .Bottom - .Top, SWP_SHOWWINDOW Or SWP_FRAMECHANGED
    End With
    
    'One last thing: because this is a (fairly? mostly? extremely?) hackish way to produce a flyout panel,
    ' we need to cover the case where the user selects outside the flyout panel, but *not* on an object
    ' that can receive focus (e.g. an exposed section of an underlying form).  Focusable objects are taken
    ' care of automatically, because a LostFocus event will fire, but non-focusable clicks are problematic.
    ' To solve this, we subclass our parent control and watch for mouse events.  Also, since we're
    ' subclassing the control anyway, we'll also hide the flyout panel if the parent window is moved.
    If (m_ParentHWnd <> 0) And PDMain.IsProgramRunning() Then
        
        'Make sure we're not currently trying to release a previous subclass attempt
        Dim subclassActive As Boolean: subclassActive = False
        If Not (m_SubclassReleaseTimer Is Nothing) Then
            If m_SubclassReleaseTimer.IsActive Then
                m_SubclassReleaseTimer.StopTimer
                subclassActive = True
            End If
        End If
        
        If (Not subclassActive) And (Not m_SubclassActive) Then
            VBHacks.StartSubclassing m_ParentHWnd, Me
            m_SubclassActive = True
        End If
        
    End If
    
    'As an additional failsafe, we also notify the master UserControl tracker that a flyout is active.
    ' If any other PD control receives focus, that tracker will automatically unload our flyout as
    ' necessary (the conditions of which are complicated; see the function for details).
    UserControls.NotifyFlyoutChangeState m_ParentHWnd, m_flyoutHWnd, Me, True
    
    m_flyoutVisible = True
    
    Exit Sub
    
UnexpectedTrouble:
    PDDebug.LogAction "WARNING!  Panels.ShowFlyout failed because of Err # " & Err.Number & ", " & Err.Description
    
End Sub

'These functions have no internal purpose.  They exist purely as a convenience for callers,
' so each toolbox doesn't have to maintain its own local synchronization flag.
Friend Function GetFlyoutSyncState() As Boolean
    GetFlyoutSyncState = m_flyoutSyncState
End Function

Friend Sub SetFlyoutSyncState(ByVal newState As Boolean)
    m_flyoutSyncState = newState
End Sub

Friend Function GetFlyoutTrackerID() As Long
    GetFlyoutTrackerID = m_flyoutTrackerID
End Function

Friend Sub HideFlyout()

    If m_flyoutVisible And (m_flyoutHWnd <> 0) Then
        
        'Notify the master UserControl tracker that our flyout is now inactive.
        UserControls.NotifyFlyoutChangeState m_ParentHWnd, m_flyoutHWnd, Me, False
        
        m_flyoutVisible = False
        SetParent m_flyoutHWnd, m_ParentHWnd
        If (m_OriginalWindowBits <> 0) Then g_WindowManager.SetWindowLongWrapper m_flyoutHWnd, m_OriginalWindowBits, , , True
        If (m_OriginalWindowBitsEx <> 0) Then g_WindowManager.SetWindowLongWrapper m_flyoutHWnd, m_OriginalWindowBits, , True, True
        g_WindowManager.SetVisibilityByHWnd m_flyoutHWnd, False
        
        'Reset any internal trackers
        m_flyoutHWnd = 0
        m_flyoutTrackerID = -1
        
        'If Aero theming is not active, hiding the flyout may cause windows beneath the current one
        ' to render incorrectly.
        If (OS.IsVistaOrLater And (Not g_WindowManager.IsDWMCompositionEnabled)) Then
            InvalidateRect 0&, VarPtr(m_flyoutRectCopy), 0&
        End If
        
        'Note that termination may result in the client site not being available.  If this happens, we simply want
        ' to continue; the subclasser will handle clean-up automatically.
        SafelyRemoveSubclass
        
        RaiseEvent FlyoutClosed(m_origTriggerObject)
        
    End If
    
End Sub

Private Sub Class_Initialize()
    'Initialize the ID to a sub-zero value (PD internal IDs always start at 0+, so this is a safe flag
    ' for indicating "no flyout ID has been set")
    m_flyoutTrackerID = -1
End Sub

Private Sub Class_Terminate()
    'Failsafe only; this state should never occur
    HideFlyout
End Sub

Private Function ISubclass_WindowMsg(ByVal hWnd As Long, ByVal uiMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal dwRefData As Long) As Long

    m_InSubclassNow = True
    
    'If certain events occur in our parent window, and our list box is visible, release it
    If m_flyoutVisible Then
        If (uiMsg = WM_ENTERSIZEMOVE) Or (uiMsg = WM_WINDOWPOSCHANGING) Then
            HideFlyout
        ElseIf (uiMsg = WM_NCDESTROY) Then
            HideFlyout
            Set m_SubclassReleaseTimer = Nothing
            VBHacks.StopSubclassing hWnd, Me
            m_ParentHWnd = 0
        End If
        
    End If
    
    'Const WM_COMMAND As Long = &H111
    'If (uiMsg = WM_COMMAND) Then Debug.Print "WM_COMMAND", uiMsg
    
    'Never eat parent window messages; just peek at them
    ISubclass_WindowMsg = VBHacks.DefaultSubclassProc(hWnd, uiMsg, wParam, lParam)
    
    m_InSubclassNow = False
    
End Function

'If a subclass exists, uninstall it.  DO NOT CALL THIS FUNCTION if the class is currently inside the subclass proc.
Private Sub RemoveSubclass()
    On Error GoTo UnsubclassUnnecessary
    If ((m_ParentHWnd <> 0) And m_SubclassActive) Then
        VBHacks.StopSubclassing m_ParentHWnd, Me
        m_ParentHWnd = 0
        m_SubclassActive = False
    End If
UnsubclassUnnecessary:
End Sub

'Release the edit box's keyboard hook.  In some circumstances, we can't do this immediately, so we set a timer that will
' release the hook as soon as the system allows.
Private Sub SafelyRemoveSubclass()
    If m_InSubclassNow Then
        If (m_SubclassReleaseTimer Is Nothing) Then Set m_SubclassReleaseTimer = New pdTimer
        m_SubclassReleaseTimer.Interval = 16
        m_SubclassReleaseTimer.StartTimer
    Else
        RemoveSubclass
    End If
End Sub

Private Sub m_SubclassReleaseTimer_Timer()
    If (Not m_InSubclassNow) Then
        m_SubclassReleaseTimer.StopTimer
        RemoveSubclass
    End If
End Sub
