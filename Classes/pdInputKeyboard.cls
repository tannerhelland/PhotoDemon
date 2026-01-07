VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdInputKeyboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Keyboard Input Handler class
'Copyright 2014-2026 by Tanner Helland
'Created: 27/May/14 (though many individual parts existed earlier than this!)
'Last updated: 09/February/17
'Last update: convert to safer subclassing via comctl32
'
'As usual, VB has some unpleasant quirks when it comes to keyboard interactions.  For example, things like specifying
' a Default control can cause Enter keypresses to become unreportable.  Keys like Tab or Arrows may not be returned
' on certain controls.  Accelerators (via Alt in particular) can also not be captured reliably.
'
'To work around these issues, I use this pdInputKeyboard class.  This class uses dynamic keyboard hooking to reliably
' capture and report any keypresses that occur while a given window has focus.  Hooking is initiated when an hWnd
' receives focus, and it is released when focus is lost.  This all happens transparently, so the only needs to worry
' about the following details:
'
'1) This class raises two events: KeyDown and KeyUp.  KeyUp corresponds to the physical release of a key.  KeyDown
'    *does not* correspond to the physical press of a key.  It is also raised if a key is pressed and held.  The user
'    cannot presently distinguish between physical presses and hold-to-repeat presses.
'
'2) To ensure that accelerators work properly, the KeyDown and KeyUp events supply a ByRef "markEventHandled" Boolean.
'    *It is very important to make use of this parameter!*  If a keypress is NOT used by your function, you need to mark
'    this variable as FALSE.  That will instruct this class to pass the keypress along to the next proc in line.  By
'    default, the "markEventHandled" value is set to TRUE, on the assumption that the caller is making use of all
'    key events, per #4...
'
'3) Rather than hook all keypresses, this class forces the caller to manually specify which keys it wants to intercept.
'    This allows unused keypresses to pass through with minimal effort, but it is not ideal if you intend to make use of
'    all (or nearly all) keys on the keyboard.  If many key events are wanted, you can remove the "check to see if key
'    should be tracked" behavior in the hook proc, but be forewarned that this may have ugly consequences, particularly
'    for accessibility purposes, where certain keys may be mapped to global functions that are not to be overridden.
'
'    tl;dr, Be careful about intercepting more keypress data than you actually require!
'
'4) Similar to VB's standard key events, Shift/Ctrl/Alt modifiers are reported as part of each event, as is the
'    *VIRTUAL KEYCODE* of the pressed key.  Everything in this class works by virtual keycodes, which are listed here:
'    http://msdn.microsoft.com/en-us/library/windows/desktop/dd375731%28v=vs.85%29.aspx
'
'    You will need to declare any relevant virtual keycodes within your own application.  For brevity's sake, I have not
'    provided them in this class.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************


Option Explicit

'The following event is raised whenever a tracked key has a change in state.
' If the user sets markEventHandled to FALSE, this class will pass the keycode on to its default handler.  Otherwise, it is assumed
' that the caller processed the key, and thus no further processing is required.
Event KeyDownCustom(ByVal Shift As ShiftConstants, ByVal vkCode As Long, ByRef markEventHandled As Boolean)
Event KeyUpCustom(ByVal Shift As ShiftConstants, ByVal vkCode As Long, ByRef markEventHandled As Boolean)

'This class tracks several different window messages.  Most objects won't need this level of tracking, but they're there if needed.
Private Const WM_ACTIVATE As Long = &H6
Private Const WM_SETFOCUS As Long = &H7
Private Const WM_KILLFOCUS As Long = &H8

'In the future, other virtual key codes can be retrieved here:
' http://msdn.microsoft.com/en-us/library/windows/desktop/dd375731%28v=vs.85%29.aspx

'Subclasser for all input window messages
Implements ISubclass

'hWnd to track.  Because this class works by dynamically hooking the object that has focus, a separate class instance
' is always required for each hWnd (unlike pdMouseInput, which can handle multiple hWnds at once).
Private m_targetHwnd As Long
Private m_SubclassActive As Boolean

'Key codes to track.  To minimize the impact of this class on keyboard events, the caller is required to *explicitly state*
' which keys it wants tracked.  Any keys that are not explicitly requested will be ignored by this class (and by "ignored"
' I mean, "left to their default handler").
'
'IMPORTANT NOTE: virtual keycodes are used, which may or may not correspond to VB key constants.  Plan accordingly!
Private m_TrackedKeys() As Long, m_NumOfTrackedKeys As Long

'Dynamic hooking requires us to track focus events with care.  When focus is lost, we must relinquish control of the keyboard.
' This value will be set to TRUE if the tracked object currently has focus.
Private m_HasFocus As Boolean

'Debugging keyboard hooks is a nightmare, so to make it easier, we track debug data on what controls have hooked the keyboard.
' This is done via a non-optional Name parameter supplied in the initial hook request.
Private m_HookedControlName As String

'When a keyboard hook is active, this will be set to TRUE, and the hook ID will be a handle to the actual Windows hook object.
Private m_HookID As Long
Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long

'While inside the hook event, this will be set to TRUE.  Because we raise events directly from the hook, we sometimes need to postpone
' crucial actions (like releasing the hook) until the hook proc has exited.
Private m_InHookNow As Boolean

'If something forces us to release our keyboard hook while in the midst of a hook, we want to delay the request until
' the hook exits.  If we don't do this, PD will crash.
Private WithEvents m_HookReleaseTimer As pdTimer
Attribute m_HookReleaseTimer.VB_VarHelpID = -1

'Parents can query last-received keycode and scancode.  These are required to map between keys and chars,
' as in the Tools > Keyboard shortcuts menu.
Private m_lastScanCode As Long, m_lastKeyCode As Long

Private Sub Class_Initialize()
    
    'Reset all input tracking variables
    m_targetHwnd = 0
    ReDim m_TrackedKeys(0) As Long
    m_NumOfTrackedKeys = 0
    
    m_HasFocus = False
    
End Sub

Private Sub Class_Terminate()
    Me.Shutdown
End Sub

'This function is automatically called at Class_Terminate time, but you can also call it manually if you want
' to ensure a particular unload time.
Friend Sub Shutdown()
    
    If (m_targetHwnd <> 0) Then
        If m_SubclassActive Then VBHacks.StopSubclassing m_targetHwnd, Me
        m_SubclassActive = False
        m_targetHwnd = 0
    End If
    
    If (m_HookID <> 0) Then
        VBHacks.NotifyPDIKHookNotNeeded ObjPtr(Me)
        UnhookWindowsHookEx m_HookID
        m_HookID = 0
    End If
    
End Sub

'Request keyboard tracking for a given hWnd.  This function will return TRUE if keyboard tracking was successfully initiated.
Friend Function CreateKeyboardTracker(ByVal nameOfTarget As String, ByVal targetHWnd As Long, ParamArray listOfKeys() As Variant) As Boolean
    
    'Failsafe check for bad user behavior
    If (m_targetHwnd <> 0) Then Debug.Print "WARNING!  Keyboard tracker requested, but we're already tracking an object!"
    m_HookedControlName = nameOfTarget
    
    If (targetHWnd <> 0) Then
        
        'Make a module-level note of the hWnd; this value will dictate all hooking and subclassing behavior
        m_targetHwnd = targetHWnd
        
        'Attach the subclasser to this hWnd
        m_SubclassActive = VBHacks.StartSubclassing(m_targetHwnd, Me)
        If m_SubclassActive Then
            
            CreateKeyboardTracker = m_SubclassActive
            
            'Key tracking is a nasty bit of business.  Let me just state that up front before
            ' breaking down how PD handles it.
        
            'Let's start by talking about something seemingly simple, like arrow keys.  By default,
            ' Windows typically handles arrow keypresses for applications.  Depending on the context,
            ' arrow keys may be used to navigate between command buttons, or navigate within a text
            ' box or edit control, or interact with controls in some other standard way.
            
            'This default behavior can be overridden by intercepting the WM_GETDLGCODE message and
            ' returning a value that specifies what default behaviors you want to override.
            ' Sounds easy, right?
            
            'In theory, it should be.  But in practice, it is anything but.  Some element of VB
            ' (who knows what) eats the WM_GETDLGCODE message before it ever arrives at our wndProc.
            ' There's possibly a way to override this behavior, or somehow intercept it, but after a
            ' great deal of testing I have not found a reliable way to do it for any possible hWnd.
            ' (This could be due to a whole bunch of factors, including the various window bits VB
            ' chooses to use, which I didn't want to custom-handle as the goal here is to support
            ' custom tracking for *any* type of hWnd.)
            
            'So, subclassing is out.  That leaves hooking, which is a more troublesome solution,
            ' especially in PD's main screen where a whole bunch of windows are simultaneously present.
            
            'Because hooking is an energy-intensive project, we don't want to hook anything more than
            ' is absolutely necessary.  So hooking is handled dynamically as any tracked window(s)
            ' gain or lose focus.  To accomplish this, key hooking requests also generate some
            ' subclassing bits as well, so we can track focus of individual windows.
            
            'We now want to copy the list of keycodes we have been passed.  If our hook encounters a
            ' keycode that was not specifically requested, we let the default window proc handle it.
            ' This allows things like the Tab key to behave normally, unless a control specifically
            ' requests otherwise.
            If (UBound(listOfKeys) >= LBound(listOfKeys)) Then
                
                'Resize our tracked keycode array
                m_NumOfTrackedKeys = UBound(listOfKeys) - LBound(listOfKeys) + 1
                ReDim m_TrackedKeys(0 To m_NumOfTrackedKeys - 1) As Long
                
                'Store the list of requested keycodes
                Dim i As Variant, curKeyIndex As Long
                curKeyIndex = 0
                
                For Each i In listOfKeys
                    m_TrackedKeys(curKeyIndex) = CLng(i)
                    curKeyIndex = curKeyIndex + 1
                Next i
            
            'It's okay to not supply keys here; the caller is allowed to postpone key requests until later
            Else
                ReDim m_TrackedKeys(0) As Long
            End If
            
        Else
            PDDebug.LogAction "WARNING: keyboard subclasser failed to subclass " & m_HookedControlName & ".  Fix it!"
        End If
    
    'The calling function didn't supply an hWnd.  Warn and exit.
    Else
        PDDebug.LogAction "WARNING: a function just requested keyboard tracking, but didn't supply an hWnd.  Fix it!"
    End If
    
End Function

'INSIDE A KEY EVENT (the only place this is valid), you can call this function to retrieve
' the unprocessed key code and scan code from the keyboardproc.
Friend Sub GetLastKeyAndScanCode(ByRef dstKeyCode As Long, ByRef dstScanCode As Long)
    dstKeyCode = m_lastKeyCode
    dstScanCode = m_lastScanCode
End Sub

'Request keyboard tracking for a given hWnd.  This function will return TRUE if keyboard tracking was successfully initiated.
Friend Sub RequestMoreKeys(ParamArray listOfKeys() As Variant)
    
    'See comments in the previous function (createKeyboardTracker) for details
    If (UBound(listOfKeys) >= LBound(listOfKeys)) Then
        
        Dim curKeyIndex As Long
        curKeyIndex = m_NumOfTrackedKeys
        
        m_NumOfTrackedKeys = m_NumOfTrackedKeys + (UBound(listOfKeys) - LBound(listOfKeys) + 1)
        
        ReDim Preserve m_TrackedKeys(0 To m_NumOfTrackedKeys - 1) As Long
        
        Dim i As Variant
        
        For Each i In listOfKeys
            If (CLng(i) <> 0) Then
                m_TrackedKeys(curKeyIndex) = CLng(i)
                curKeyIndex = curKeyIndex + 1
            End If
        Next i
    
    End If
    
End Sub

Private Function GetLoWord(ByVal lParam As Long) As Integer
    If lParam And &H8000& Then
        GetLoWord = &H8000 Or (lParam And &H7FFF&)
    Else
        GetLoWord = lParam And &HFFFF&
    End If
End Function

'If a hook is active, this timer will repeatedly try to kill it.  Do not enable it until you are certain the hook
' needs to be released.  (This is used as a failsafe if we cannot immediately release the hook when focus is lost,
' for example if we are currently inside an external event, as happens in the Layer toolbox, which hides the active
' text box inside the KeyPress event.)
Private Sub m_HookReleaseTimer_Timer()
    If (Not m_InHookNow) Then
        If (m_HookID <> 0) Then
            RemoveKeyboardHook
            m_HookReleaseTimer.StopTimer
        
        'Something weird happened and the hook was released; stop timing now!
        Else
            m_HookReleaseTimer.StopTimer
        End If
    End If
End Sub

'Install a keyboard hook in our window
Private Sub InstallKeyboardHook()
    
    'Check for an existing hook
    If (Not m_HasFocus) And (m_targetHwnd <> 0) And (m_HookID = 0) And PDMain.IsProgramRunning() Then
        
        'If a hook is already active, and we're trying to unhook it safely, we don't want to reinstall our hook
        Dim tryingToUnhook As Boolean: tryingToUnhook = False
        If (Not m_HookReleaseTimer Is Nothing) Then
            tryingToUnhook = m_HookReleaseTimer.IsActive
            If tryingToUnhook Then m_HookReleaseTimer.StopTimer
        End If
        
        'Note that this window is now active
        m_HasFocus = True
        
        'Install a new hook for this window
        If (Not tryingToUnhook) Then
            m_HookID = VBHacks.NotifyPDIKHookNeeded(Me)
            UserControls.PDControlReceivedFocus m_targetHwnd
        End If
        
    End If

End Sub

'If a hook exists, uninstall it.  DO NOT CALL THIS FUNCTION if the class is currently inside the hook proc.
Private Sub RemoveKeyboardHook()
    
    If (m_HookID <> 0) Then
        
        VBHacks.NotifyPDIKHookNotNeeded ObjPtr(Me)
        UnhookWindowsHookEx m_HookID
        m_HookID = 0
        
        'Debug.Print "Uninstalling keyboard hook for " & m_HookedControlName & ":" & m_targetHwnd
        UserControls.PDControlLostFocus m_targetHwnd
        
    End If
    
End Sub

'Release the edit box's keyboard hook.  In some circumstances, we can't do this immediately, so we set a timer that will
' release the hook as soon as the system allows.
Private Sub SafelyRemoveKeyboardHook()

    If m_InHookNow Then
        'Debug.Print "Starting safe unhook procedure for " & m_HookedControlName & ":" & m_targetHwnd
        If (m_HookReleaseTimer Is Nothing) Then Set m_HookReleaseTimer = New pdTimer
        m_HookReleaseTimer.Interval = 16
        m_HookReleaseTimer.StartTimer
    Else
        RemoveKeyboardHook
    End If
    
End Sub

Private Function HandleKeyProc(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long, ByRef msgEaten As Boolean) As Long
    
    'Ensure 1) the user wants key events, and 2) PD's main menu is not dropped (we don't receive a
    ' focus change notification for this, so we need to check it here)
    If (m_NumOfTrackedKeys > 0) Then
    If (Not Menus.IsMainMenuActive()) Then
        
        'Cache last key-code and scan-code (parents can query us for this data)
        m_lastKeyCode = wParam
        
        'From https://learn.microsoft.com/en-us/previous-versions/windows/desktop/legacy/ms644984(v=vs.85)
        ' (re lParam bitfields):
        ' 16-23: The scan code. The value depends on the OEM.
        m_lastScanCode = (lParam \ 65536) And &HFF&
        
        'Manually pull key modifier states (shift, control, alt/menu) in advance; these are standard for all key events
        Dim retShiftConstants As ShiftConstants
        If IsVirtualKeyDown(VK_SHIFT) Then retShiftConstants = retShiftConstants Or vbShiftMask
        If IsVirtualKeyDown(VK_CONTROL) Then retShiftConstants = retShiftConstants Or vbCtrlMask
        If IsVirtualKeyDown(VK_ALT) Then retShiftConstants = retShiftConstants Or vbAltMask
        
        'The keycode of the key being pressed is stored in the wParam.  See if it is a key we have been asked to track.
        Dim i As Long
        For i = 0 To m_NumOfTrackedKeys - 1
        
            If (wParam = m_TrackedKeys(i)) Then
            
                'This is a key we are supposed to be tracking.  Check its state by using isVirtualKeyDown.
                If IsVirtualKeyDown(wParam) Then
                    
                    'At present, we ignore messages raised by PeekMessage()
                    Const HC_NOREMOVE As Long = 3
                    If (nCode <> HC_NOREMOVE) Then
                    
                        msgEaten = True
                    
                        'This key is being pressed.  Raise the associated event.
                        RaiseEvent KeyDownCustom(retShiftConstants, wParam, msgEaten)
                        Exit For
                        
                    End If
                    
                'The key has just been released.
                Else
                
                    'This is a little complicated (see http://msdn.microsoft.com/en-us/library/windows/desktop/ms644984%28v=vs.85%29.aspx
                    ' for details).  Basically, lParam contains detailed key state tracking.  Bits 30 and 31 contain previous state and
                    ' transition state values, respectively.  By testing these values, we can insure that only one KeyUp event is raised;
                    ' specifically, it is only raised at the moment when the transition state is UP and the previous state is DOWN.
                    If ((lParam And 1) <> 0) And ((lParam And 3) = 1) Then
                
                        msgEaten = True
                        
                        'This key is being released.  Raise the associated event.
                        RaiseEvent KeyUpCustom(retShiftConstants, wParam, msgEaten)
                        Exit For
                        
                    End If
                
                End If
            
            End If
        
        Next i
    
    End If
    End If
        
End Function

Friend Function PDIKKeyboardProc(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
    m_InHookNow = True
    
    'Ignore this event completely if we've shut down our hook internally.  (Because you can't use AddressOf in a class,
    ' we have to use an ugly, hackish technique to handle these callbacks.)
    If (m_HookID <> 0) Then
        
        Dim msgEaten As Boolean: msgEaten = False
        
        'If we don't have focus, ignore this completely
        If m_HasFocus Then
            
            'MSDN states that negative codes must be passed to the next hook, without processing
            ' (see http://msdn.microsoft.com/en-us/library/ms644984.aspx)
            If (nCode >= 0) Then PDIKKeyboardProc = HandleKeyProc(nCode, wParam, lParam, msgEaten)
            
        End If
        
        'Per MSDN, return the value of CallNextHookEx, contingent on whether or not we handled the keypress internally.
        ' Note that if we do not manually handle a keypress, this behavior allows the default keyhandler to deal with
        ' the pressed keys (and raise its own WM_CHAR events, etc).
        If (Not msgEaten) Then
            PDIKKeyboardProc = CallNextHookEx(0&, nCode, wParam, lParam)
        Else
            PDIKKeyboardProc = 1&
        End If
        
    Else
        PDIKKeyboardProc = CallNextHookEx(0&, nCode, wParam, lParam)
    End If
    
    m_InHookNow = False
    
End Function

Private Sub HandleWMActivate(ByVal hWnd As Long, ByVal uiMsg As Long, ByVal wParam As Long, ByVal lParam As Long)

    'Figure out if the window is being activated or deactivated.
    Dim activationStatus As Long
    activationStatus = GetLoWord(wParam)
                
    'This hWnd is being deactivated.  If a hook has been installed, remove it now.
    Const WA_INACTIVE As Long = 0
    If (activationStatus = WA_INACTIVE) Then
        HandleWMKillFocus
            
    'This hWnd is being activated.  Install a hook now.
    Else
        HandleWMSetFocus
    End If
        
End Sub

Private Sub HandleWMSetFocus()
    If (Not m_HasFocus) Then
        InstallKeyboardHook
    Else
        PDDebug.LogAction m_HookedControlName & " just gained focus, but a keyboard hook is already installed??"
    End If
End Sub

Private Sub HandleWMKillFocus()

    'Release our hook.  In some circumstances, we can't do this immediately, so we set a timer that will
    ' release the hook as soon as the system allows.
    SafelyRemoveKeyboardHook
    
    'Immediately mark the control-wide focus state, even though the hook has not yet been removed
    If m_HasFocus Then
        m_HasFocus = False
    Else
        'Quick note (because I'm commenting this out in an unrelated commit while working on hotkeys);
        ' this path appears to trigger if you use a hotkey to trigger a control, but that control didn't
        ' receive focus first.  (To test: Open an image, hit Ctrl+R to load `Image > Resize`, then press
        ' ESCAPE and you'll see the warning below.)  This is a safe workflow that doesn't cause any issues,
        ' so I'm not sure the warning below is useful.  TODO: consider removing branch entirely after
        ' more testing.
        'PDDebug.LogAction m_HookedControlName & " just lost focus, but no keyboard hook was ever installed??"
    End If
    
End Sub

Private Function ISubclass_WindowMsg(ByVal hWnd As Long, ByVal uiMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal dwRefData As Long) As Long

    'Note that this function only peeks at passed messages; it never consumes the messages
    
    'A floating toolbox window is being activated or deactivated.
    If (uiMsg = WM_ACTIVATE) Then
        HandleWMActivate hWnd, uiMsg, wParam, lParam
    
    'Docked toolbox windows and standard controls raise WM_SETFOCUS instead of WM_ACTIVATE
    ElseIf (uiMsg = WM_SETFOCUS) Then
        HandleWMSetFocus
        
    ElseIf (uiMsg = WM_KILLFOCUS) Then
        HandleWMKillFocus
        
    ElseIf (uiMsg = WM_NCDESTROY) Then
        VBHacks.StopSubclassing hWnd, Me
        m_targetHwnd = 0
        m_SubclassActive = False
    End If
    
    ISubclass_WindowMsg = VBHacks.DefaultSubclassProc(hWnd, uiMsg, wParam, lParam)
    
End Function
