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
'Copyright 2014-2015 by Tanner Helland
'Created: 27/May/14 (though many individual parts existed earlier than this!)
'Last updated: 30/September/14
'Last update: separate keyboard tracking into its own class, as it requires a lot of extra work
'
'As usual, VB has some unpleasant quirks when it comes to keyboard interactions.  For example, things like specifying
' a Default control can cause Enter keypresses to become unreportable.  Keys like Tab or Arrows may not be returned
' on certain controls.  Accelerators (via Alt in particular) can also not be captured reliably.
'
'To work around these issues, I use this pdInputKeyboard class.  This class uses dynamic keyboard hooking to reliably
' capture and report any keypresses that occur while a given window has focus.  Hooking is initiated when an hWnd
' receives focus, and it is released when focus is lost.  This all happens transparently, so the user need not worry
' about anything except the following:
'
'1) Unlike pdInputMouse, this class only supports a single hWnd at a time.  This is by design, to get around some of
'    the uglier complications of dynamic hooking.  If you need to track multiple hWnds, you will need multiple
'    instances of this class.
'
'2) This class raises two events: KeyDown and KeyUp.  KeyUp corresponds to the physical release of a key.  KeyDown
'    *does not* correspond to the physical press of a key.  It is also raised if a key is pressed and held.  The user
'    cannot presently distinguish between physical presses and hold-to-repeat presses.
'
'3) To ensure that accelerators work properly, the KeyDown and KeyUp events supply a ByRef "markEventHandled" Boolean.
'    *It is very important to make use of this parameter!*  If a keypress is NOT used by your function, you need to mark
'    this variable as FALSE.  That will instruct this class to pass the keypress along to the next proc in line.  By
'    default, the "markEventHandled" value is set to TRUE, on the assumption that the caller is making use of all
'    key events, per #4...
'
'4) Rather than hook all keypresses, this class forces the caller to manually specify which keys it wants to intercept.
'    This allows unused keypresses to pass through with minimal effort, but it is not ideal if you intend to make use of
'    all (or nearly all) keys on the keyboard.  If many key events are wanted, you can remove the "check to see if key
'    should be tracked" behavior in the hook proc, but be forewarned that this may have ugly consequences, particularly
'    for accessibility purposes, where certain keys may be mapped to global functions that are not to be overridden.
'
'    tl;dr, Be careful about intercepting more keypress data than you actually require!
'
'5) Similar to VB's standard key events, Shift/Ctrl/Alt modifiers are reported as part of each event, as is the
'    *VIRTUAL KEYCODE* of the pressed key.  Everything in this class works by virtual keycodes, which are listed here:
'    http://msdn.microsoft.com/en-us/library/windows/desktop/dd375731%28v=vs.85%29.aspx
'
'    You will need to declare any relevant virtual keycodes within your own application.  For brevity's sake, I have not
'    provided them in this class.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************


Option Explicit

'The following event is raised whenever a tracked key has a change in state.
' If the user sets markEventHandled to FALSE, this class will pass the keycode on to its default handler.  Otherwise, it is assumed
' that the caller processed the key, and thus no further processing is required.
Event KeyDownCustom(ByVal Shift As ShiftConstants, ByVal vkCode As Long, ByRef markEventHandled As Boolean)
Event KeyUpCustom(ByVal Shift As ShiftConstants, ByVal vkCode As Long, ByRef markEventHandled As Boolean)

'This class tracks several different window messages.  Most objects won't need this level of tracking, but they're there if needed.
Private Const WM_KEYDOWN As Long = &H100
Private Const WM_GETDLGCODE As Long = &H87

Private Const WM_ACTIVATE As Long = &H6
Private Const WM_SETFOCUS As Long = &H7
Private Const WM_KILLFOCUS As Long = &H8

'The Alt mask won't be returned by mouse WM notifications, so we need to retrieve it manually
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

'Virtual key-codes currently supported by pdInput
Private Const VK_SHIFT As Long = &H10
Private Const VK_CONTROL As Long = &H11
Private Const VK_ALT As Long = &H12    'Note that VK_ALT is referred to as VK_MENU in MSDN documentation!

'In the future, other virtual key codes can be retrieved here:
' http://msdn.microsoft.com/en-us/library/windows/desktop/dd375731%28v=vs.85%29.aspx

'To see if a given key is marked as an "extended key" by the system (e.g. Return vs NumPad Enter key, which have the same
' virtual key code), perform a bitwise AND (&&) of this flag against the lParam of the WM_KEYDOWN event
Private Const KF_EXTENDED As Long = &H1000000

'Master subclasser for all input actions
Private cSubclass As cSelfSubHookCallback

'hWnd to track.  Because this class works by dynamically hooking the object that has focus, a separate class instance
' is always required for each hWnd (unlike pdMouseInput, which can handle multiple hWnds at once).
Private m_targetHwnd As Long

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

Private Sub Class_Initialize()
    
    'Reset all input tracking variables
    m_targetHwnd = 0
    ReDim m_TrackedKeys(0) As Long
    m_NumOfTrackedKeys = 0
    
    m_HasFocus = False
    
    'Initialize the subclasser
    Set cSubclass = New cSelfSubHookCallback
    
End Sub

'Request keyboard tracking for a given hWnd.  This function will return TRUE if keyboard tracking was successfully initiated.
Public Function CreateKeyboardTracker(ByVal nameOfTarget As String, ByVal targetHwnd As Long, ParamArray listOfKeys() As Variant) As Boolean
    
    m_HookedControlName = nameOfTarget
    
    If targetHwnd <> 0 Then
        
        'Make a module-level note of the hWnd; this value will dictate all hooking and subclassing behavior
        m_targetHwnd = targetHwnd
        
        'Attach the subclasser to this hWnd
        If Not cSubclass.ssc_Subclass(targetHwnd, , 1, Me) Then
            
            Debug.Print "WARNING: keyboard subclasser failed to subclass " & m_HookedControlName & ".  Fix it!"
            CreateKeyboardTracker = False
            Exit Function
        
        End If
        
        'Key tracking is a nasty bit of business.  Let me just state that up front before breaking down how PD handles it.
    
        'Let's start by talking about something seemingly simple, like arrow keys.  By default, Windows typically handles arrow
        ' keypresses for applications.  Depending on the context, arrow buttons may be used to navigate between command buttons,
        ' or navigate within a text box or edit control, or interact with controls in some other standard way.
        
        'This default behavior can be overridden by intercepting the WM_GETDLGCODE message and returning a value that specifies
        ' what default behaviors you want to override.  Sounds easy, right?
        
        'In theory, it should be.  But in practice, it is anything but.  Some element of VB (who knows what) eats the WM_GETDLGCODE
        ' message before it ever arrives at our wndProc.  It is possible that there's a way to override this behavior, or
        ' somehow intercept it, but after a great deal of testing I have not found a reliable way to do it for any possible hWnd.
        ' (This could be due to a whole bunch of factors, including the various window bits VB chooses to use, which I didn't want
        ' to custom-handle as the goal here is to support custom tracking for *any* type of hWnd.)
        
        'So, subclassing is out.  That leaves hooking, which is a more troublesome solution, especially in PD's main screen where
        ' a whole bunch of windows are simultaneously present.
        
        'Because hooking is an energy-intensive project, we don't want to hook anything more than is absolutely necessary.  So
        ' hooking is handled dynamically as any tracked window(s) gain or lose focus.  To accomplish this, key hooking requests
        ' also generate some subclassing bits as well, so we can track focus of our individual windows.
        
        'Subclass two messages: WM_ACTIVATE (for floating windows) and WM_SETFOCUS/KILLFOCUS (for docked windows)
        cSubclass.ssc_AddMsg targetHwnd, MSG_BEFORE, WM_ACTIVATE, WM_SETFOCUS, WM_KILLFOCUS
        
        'We now want to copy the list of keycodes we have been passed.  If our hook encounters a keycode that was not specifically
        ' requested, we're going to let the default window proc handle it.  This allows things like the Tab key to behave normally,
        ' unless a control specifically requests otherwise.
        If UBound(listOfKeys) >= LBound(listOfKeys) Then
            
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
        
        Else
            
            'It's okay to not supply keys here; the caller is allowed to postpone key requests until later
            ReDim m_TrackedKeys(0) As Long
            
        End If
    
    'The calling function didn't supply an hWnd.  Warn and exit.
    Else
        Debug.Print "WARNING: a function just requested keyboard tracking, but didn't supply an hWnd.  Fix it!"
        CreateKeyboardTracker = False
        Exit Function
    End If
    
    CreateKeyboardTracker = True
    
End Function

'Request keyboard tracking for a given hWnd.  This function will return TRUE if keyboard tracking was successfully initiated.
Public Sub RequestMoreKeys(ParamArray listOfKeys() As Variant)
    
    'See comments in the previous function (createKeyboardTracker) for details
    If UBound(listOfKeys) >= LBound(listOfKeys) Then
        
        Dim curKeyIndex As Long
        curKeyIndex = m_NumOfTrackedKeys
        
        m_NumOfTrackedKeys = m_NumOfTrackedKeys + (UBound(listOfKeys) - LBound(listOfKeys) + 1)
        
        ReDim Preserve m_TrackedKeys(0 To m_NumOfTrackedKeys - 1) As Long
        
        Dim i As Variant
        
        For Each i In listOfKeys
            If CLng(i) <> 0 Then
                m_TrackedKeys(curKeyIndex) = CLng(i)
                curKeyIndex = curKeyIndex + 1
            End If
        Next i
    
    End If
    
End Sub

Private Sub Class_Terminate()

    'Release all subclassing and hooking
    cSubclass.shk_TerminateHooks
    cSubclass.ssc_Terminate
    
End Sub

'Note that the vKey constant below is a virtual key mapping, not necessarily a standard VB key constant
Private Function IsVirtualKeyDown(ByVal vKey As Long) As Boolean
    IsVirtualKeyDown = GetAsyncKeyState(vKey) And &H8000
End Function

Private Function getHiWord(ByVal lParam As Long) As Integer
    If lParam And &H80000000 Then
        getHiWord = (lParam \ 65535) - 1
    Else
        getHiWord = lParam \ 65535
    End If
End Function

Private Function getLoWord(ByVal lParam As Long) As Integer
    If lParam And &H8000& Then
        getLoWord = &H8000 Or (lParam And &H7FFF&)
    Else
        getLoWord = lParam And &HFFFF&
    End If
End Function

'This routine MUST BE KEPT as the next-to-last routine for this form. Its ordinal position determines its ability to hook properly.
Private Sub myHookProc(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal lHookType As eHookType, ByRef lParamUser As Long)
'*************************************************************************************************
' http://msdn2.microsoft.com/en-us/library/ms644990.aspx
'* bBefore    - Indicates whether the callback is before or after the next hook in chain.
'* bHandled   - In a before next hook in chain callback, setting bHandled to True will prevent the
'*              message being passed to the next hook in chain and (if set to do so).
'* lReturn    - Return value. For Before messages, set per the MSDN documentation for the hook type
'* nCode      - A code the hook procedure uses to determine how to process the message
'* wParam     - Message related data, hook type specific
'* lParam     - Message related data, hook type specific
'* lHookType  - Type of hook calling this callback
'* lParamUser - User-defined callback parameter. Change vartype as needed (i.e., Object, UDT, etc)
'*************************************************************************************************
    
    
    If lHookType = WH_KEYBOARD Then
    
        bHandled = False
        
        'Manually pull key modifier states (shift, control, alt/menu) in advance; these are standard for all key events
        Dim retShiftConstants As ShiftConstants
        If IsVirtualKeyDown(VK_SHIFT) Then retShiftConstants = retShiftConstants Or vbShiftMask
        If IsVirtualKeyDown(VK_CONTROL) Then retShiftConstants = retShiftConstants Or vbCtrlMask
        If IsVirtualKeyDown(VK_ALT) Then retShiftConstants = retShiftConstants Or vbAltMask
        
        'The keycode of the key being pressed is stored in the wParam.  See if it is a key we have been asked to track.
        Dim i As Long
        For i = 0 To m_NumOfTrackedKeys - 1
        
            If wParam = m_TrackedKeys(i) Then
            
                'This is a key we are supposed to be tracking.  Check its state by using isVirtualKeyDown.
                If IsVirtualKeyDown(wParam) Then
                
                    bHandled = True
                
                    'This key is being pressed.  Raise the associated event.
                    RaiseEvent KeyDownCustom(retShiftConstants, wParam, bHandled)
                    Exit For
                
                'The key has just been released.
                Else
                
                    'This is a little complicated (see http://msdn.microsoft.com/en-us/library/windows/desktop/ms644984%28v=vs.85%29.aspx
                    ' for details).  Basically, lParam contains detailed key state tracking.  Bits 30 and 31 contain previous state and
                    ' transition state values, respectively.  By testing these values, we can insure that only one KeyUp event is raised;
                    ' specifically, it is only raised at the moment when the transition state is UP and the previous state is DOWN.
                    If ((lParam And 1) <> 0) And ((lParam And 3) = 1) Then
                
                        bHandled = True
                        
                        'This key is being released.  Raise the associated event.
                        RaiseEvent KeyUpCustom(retShiftConstants, wParam, bHandled)
                        Exit For
                        
                    End If
                
                End If
            
            End If
        
        Next i
        
        'If bHandled Then
        '    Debug.Print "pdInput is handling keypress: " & bHandled & ", " & nCode
        'Else
        '    Debug.Print "pdInput is NOT handling keypress: " & bHandled & ", " & nCode
        'End If
                
        'Per MSDN, return the value of CallNextHookEx
        If (Not bHandled) Then
            lReturn = CallNextHookEx(0, nCode, wParam, lParam)
        Else
            lReturn = 1
        End If
            
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
    
    'As a failsafe, make sure the window handle matches the hWnd we are tracking.  (This should never *not* be true,
    ' but better safe than sorry.)
    If (lng_hWnd = m_targetHwnd) Then
    
        'Parse window messages and raise any standard events; at present, this includes any messages that affect focus
        ' state of a window or control.
        Select Case uMsg
                
            'A floating toolbox window is being activated or deactivated.  This is important as this class uses
            ' activation status to install and remove keyboard hooks.
            Case WM_ACTIVATE
                
                'Figure out if the window is being activated or deactivated.
                Const WA_INACTIVE As Long = 0
                Const WA_ACTIVE As Long = 1
                Const WA_CLICKACTIVE As Long = 2
                
                Dim activationStatus As Long
                activationStatus = wParam And &HFFFF
                
                'This hWnd is being deactivated.  If a hook has been installed, remove it now.
                If activationStatus = WA_INACTIVE Then
                    
                    'Check for an existing hook
                    If m_HasFocus Then
                        
                        'A hook was found.  Uninstall it now.
                        Debug.Print "Uninstalling keyboard hook for " & m_HookedControlName & " due to WM_ACTIVATE"
                        cSubclass.shk_UnHook WH_KEYBOARD
                        
                        'Note that this window is now inactive
                        m_HasFocus = False
                        
                    Else
                        Debug.Print m_HookedControlName & " just lost focus, but no keyboard hook was ever installed??"
                    End If
                        
                'This hWnd is being activated.  Install a hook now.
                Else
                    
                    'Check for an existing hook
                    If Not m_HasFocus Then
                        
                        'No hook exists.  Hook the keyboard for this window.
                        Debug.Print "Installing keyboard hook for " & m_HookedControlName & " due to WM_ACTIVATE"
                        cSubclass.shk_SetHook WH_KEYBOARD, False, MSG_BEFORE, m_targetHwnd, 2, Me
                        
                        'Note that this window is now active
                        m_HasFocus = True
                        
                    Else
                        Debug.Print m_HookedControlName & " just gained focus, but a keyboard hook is already installed??"
                    End If
                    
                End If
                                
            'Docked toolbox windows and standard controls will raise WM_SETFOCUS instead of WM_ACTIVATE
            Case WM_SETFOCUS
                
                'Check for an existing hook
                If Not m_HasFocus Then
                
                    'No hook exists.  Hook the control now.
                    Debug.Print "Installing keyboard hook for " & m_HookedControlName & " due to WM_SETFOCUS"
                    cSubclass.shk_SetHook WH_KEYBOARD, False, MSG_BEFORE, m_targetHwnd, 2, Me
                    
                    'Note that this window is now active
                    m_HasFocus = True
                    
                Else
                    Debug.Print m_HookedControlName & " just gained focus, but a keyboard hook is already installed??"
                End If
                
            Case WM_KILLFOCUS
                
                'Check for an existing hook
                If m_HasFocus Then
                    
                    'A hook exists.  Uninstall it now.
                    Debug.Print "Uninstalling keyboard hook for " & m_HookedControlName & " due to WM_KILLFOCUS"
                    cSubclass.shk_UnHook WH_KEYBOARD
                        
                    'Note that this window is now considered inactive
                    m_HasFocus = False
                    
                Else
                    Debug.Print m_HookedControlName & " just lost focus, but no keyboard hook was ever installed??"
                End If
                
        End Select

    
    'Somehow, the hWnd we've captured is not the one being tracked!  Abandon ship.
    Else
    
        Debug.Print "WARNING!  Irrelevant hWnd intercepted by pdInput: " & lng_hWnd
        bHandled = False
    
    End If



' *************************************************************
' C A U T I O N   C A U T I O N   C A U T I O N   C A U T I O N
' -------------------------------------------------------------
' DO NOT ADD ANY OTHER CODE BELOW THE "END SUB" STATEMENT BELOW
'   add this warning banner to the last routine in your class
' *************************************************************
End Sub
