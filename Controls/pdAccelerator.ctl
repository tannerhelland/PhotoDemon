VERSION 5.00
Begin VB.UserControl pdAccelerator 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   CanGetFocus     =   0   'False
   ClientHeight    =   3600
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
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "pdAccelerator.ctx":0000
End
Attribute VB_Name = "pdAccelerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Accelerator ("Hotkey") handler
'Copyright 2013-2026 by Tanner Helland and contributors
'Created: 06/November/15 (split off from a heavily modified vbaIHookControl by Steve McMahon)
'Last updated: 06/October/21
'Last update: map "duplicate" virtual key IDs (e.g. keyboard + and numpad +) to the same internal key ID;
'             doing that here spares us from needing to track it in the hotkey collection
'
'In its early years, PD used a "hook control" by vbAccelerator.com to handle program hotkeys:
' http://www.vbaccelerator.com/home/VB/Code/Libraries/Hooks/Accelerator_Control/article.asp
'
'In 2013, I rewrote the control to solve some glaring stability issues.  Over time, I rewrote it more
' and more, tacking on PhotoDemon-specific features and attempting to fix problematic bugs,
' until ultimately the control grew into a horrible mishmash of spaghetti code: some old, some new,
' some completely unused, and too much of it that remained stubbornly unreliable.
'
'Because dynamic hooking has enormous potential for causing hard-to-replicate bugs, a ground-up
' rewrite seemed long overdue.  Thus this control was born.
'
'Thank you to Steve McMahon for his original implementation, which was my first introduction to
' hooking from VB6.  Steve's work is still a useful reference for beginners, and you can find the
' original here (hopefully... Steve's work has intermittently disappeared from the web in recent years):
' http://www.vbaccelerator.com/home/VB/Code/Libraries/Hooks/Accelerator_Control/
'
'Thank you also to Jason Brown (https://github.com/jpbro), who submitted many fixes and improvements to
' this module over the years.  Hotkey behavior in PhotoDemon is greatly improved thanks to him.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'This control only raises a single event: "HotkeyPressed".  This event is raised when a valid hotkey
' mapping is recognized (e.g. the hotkey has been validated by PD's central hotkey manager).
'
'At present, the only consumer of these events is FormMain, and FormMain basically just relays the
' raised hotkey ID back to PD's central hotkey manager for further action.  (This separation of duties
' is useful because if FormMain is disabled for some reason, hotkey handling gets properly postponed too.)
Public Event HotkeyPressed(ByVal hotkeyID As Long)

'PhotoDemon previously used on-demand virtual-key tracking for hotkeys, but this proved to be a bad
' idea because system hotkeys that switch focus between apps (e.g. Alt+Tab) throw off our tracking
' state(s).  A better solution is to manually track key up/down state for Ctrl/Alt/Shift presses and
' cache the results locally (see https://github.com/tannerhelland/PhotoDemon/issues/267 for details.)
Private m_CtrlDown As Boolean, m_AltDown As Boolean, m_ShiftDown As Boolean

'If the control's hook proc is active and primed, this handle will be non-zero.
' (Zero indicates an inactive or failed hook.)
Private m_HookID As Long
Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long

'When the control is actually inside the hook procedure, we set a module-level flag to TRUE.
' The hook *cannot be removed until this flag is FALSE*.  (PD will crash.)  To ensure correct
' unhooking behavior, we use a fire-once timer to initiate potentially long-running tasks
' *after* the hook exits.
Private m_InHookNow As Boolean, m_InFireTimerNow As Boolean

'To reduce the potential for double-fired keys, we track the last-fired accelerator ID and the time
' when we launched the associated action.  The current system keyboard delay must elapse before we
' fire that same accelerator a second time.
Private m_LastHotkeyIndex As Long, m_TimerAtAcceleratorPress As Currency

'This control may be de/activated many times in a given session.  (If an edit box gets focus,
' for example, PhotoDemon will disable hooking here so that shared hotkeys - like Ctrl+A - apply to
' the edit box instead of this control "stealing" them.)  To provide useful debug info, PD will note
' when the first attempt at hooking fails (during program startup) but *not* subsequent attempts.
' Buggy system-wide hotkey managers are the most common cause for failed hooking, and it's helpful to
' know this if users complain, but we don't need a billion entries in the debug log for the same thing.
Private m_SubsequentInitialization As Boolean

'In-memory timers are used for safely firing accelerators *outside* the hook event itself, as well as
' releasing active hooks after the hookproc safely exits.
Private WithEvents m_ReleaseTimer As pdTimer
Attribute m_ReleaseTimer.VB_VarHelpID = -1
Private WithEvents m_FireTimer As pdTimer
Attribute m_FireTimer.VB_VarHelpID = -1

'Thanks to a patch by jpbro (https://github.com/tannerhelland/PhotoDemon/pull/248), PD no longer drops
' accelerators that are triggered in quick succession.  Instead, it queues them and fires them in turn.
' Two collections are used for this - one that fires off all still-need-to-be-processed events, and a
' backup queue that accumulates events in the background (while the current queue is being worked through).
Private m_AcceleratorQueue As VBA.Collection        'Active queue of accelerators being processed (in FIFO order)
Private m_AcceleratorAccumulator As VBA.Collection  'Queue of accelerators allowed to accumulate while the active queue is processing

'Standard events required by all PD controls
Public Function GetControlType() As PD_ControlType
    GetControlType = pdct_Accelerator
End Function

Public Function GetControlName() As String
    GetControlName = UserControl.Extender.Name
End Function

'The Enabled property is a bit unique; see http://msdn.microsoft.com/en-us/library/aa261357%28v=vs.60%29.aspx
Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal newValue As Boolean)
    UserControl.Enabled = newValue
    PropertyChanged "Enabled"
End Property

'Accelerators may trigger long-running events.  Hook procs are supposed to be dealt with ASAP.
' To make these two conflicting requirements work, PD doesn't fire accelerator events from within
' the hook proc.  Instead, it queues them, allows the hook to exit, then safely triggers them
' 1/60 second later.
Private Sub m_FireTimer_Timer()
    
    'If we're still inside the hookproc, we'll wait another 16 ms before testing the keypress.
    If (Not m_InHookNow) Then
        
        'PhotoDemon maintains a list of conditions that disallow hotkeys from triggering.
        ' (For example, a modal dialog is active.)
        If (Not CanIRaiseAnAcceleratorEvent(True)) Then
            
            'We are not currently allowed to raise any events, so postpone this event until the
            ' next timer event.  (As an added failsafe, if the program is shutting down,
            ' we'll forcibly stop the timer so we don't raise any more hotkey events.)
            If g_ProgramShuttingDown Then m_FireTimer.StopTimer
            Exit Sub
            
        End If
        
        'To prevent issues with reentrancy, note that we're already inside the hotkey processing timer
        m_InFireTimerNow = True
        
        'Because the accelerator is about to be processed, we can stop the underlying timer object.
        ' (Obviously this function will continue running, but we don't need it to trigger again until
        ' a new hotkey is pressed.)
        m_FireTimer.StopTimer
        
        'Process accelerators from the active queue in FIFO order
        Dim i As Long, idxHotkey As Long
        For i = 1 To m_AcceleratorQueue.Count
            idxHotkey = m_AcceleratorQueue.Item(i)
            If (idxHotkey >= 0) Then RaiseEvent HotkeyPressed(idxHotkey)
        Next i
        
        'Swap the active queue for the accumuator queue and empty the old accumulator
        Set m_AcceleratorQueue = m_AcceleratorAccumulator
        Set m_AcceleratorAccumulator = New VBA.Collection
        
        'If the backup queue collected any hotkeys while we were processing the current batch,
        ' restart the timer.
        If (m_AcceleratorQueue.Count > 0) Then m_FireTimer.StartTimer
         
        m_InFireTimerNow = False
        
    End If
    
End Sub

'Bad things happen if we remove the hook while still inside the hook event.  To work around this,
' we use a safety timer that releases the hook on a delay.
Private Sub m_ReleaseTimer_Timer()
    If (m_HookID <> 0) Then
        SafelyReleaseHook
    Else
        m_ReleaseTimer.StopTimer
    End If
End Sub

'Hooks cannot be released while actually inside the hook proc.  Call this function to safely release a hook,
' even from within the hook proc.
Private Sub SafelyReleaseHook()
    
    'Failsafe check
    If (Not PDMain.IsProgramRunning()) Then Exit Sub
    
    'If this is being called from within the hook, activate the safety hook-release timer
    If m_InHookNow Then
        If (Not m_ReleaseTimer Is Nothing) Then
            If (Not m_ReleaseTimer.IsActive) Then m_ReleaseTimer.StartTimer
        End If
        
    'If we're not inside the hook, free the hook immediately.
    Else
        
        If (m_HookID <> 0) Then
            UnhookWindowsHookEx m_HookID
            VBHacks.NotifyAcceleratorHookNotNeeded ObjPtr(Me)
            m_HookID = 0
        End If
        
        'Also deactivate the failsafe timer (there's no harm in doing this if it's not running)
        If (Not m_ReleaseTimer Is Nothing) Then m_ReleaseTimer.StopTimer
        
    End If
    
End Sub

'Prior to shutdown, you can call this function to forcibly release any active hotkey-related handles.
Public Sub ReleaseResources()
    Me.DeactivateHook True
    If Not (m_ReleaseTimer Is Nothing) Then Set m_ReleaseTimer = Nothing
    If Not (m_FireTimer Is Nothing) Then Set m_FireTimer = Nothing
End Sub

Private Sub UserControl_Initialize()
    
    'To correctly handle multiple hotkeys entered in quick succession (and cover the case where a
    ' previous hotkey launched a long-running task), we store hotkeys in a collection, then fire
    ' them off in FIFO order.
    Set m_AcceleratorQueue = New VBA.Collection
    Set m_AcceleratorAccumulator = New VBA.Collection
    
    'PD (bravely? stupidly?) still activates hotkeys in the IDE.  You might consider disabling this
    ' if you want more stable behavior during pause-and-edit debugging.
    If PDMain.IsProgramRunning() Then
        
        'UI-related timers run at 60 fps
        Set m_ReleaseTimer = New pdTimer
        m_ReleaseTimer.Interval = 17
        
        Set m_FireTimer = New pdTimer
        m_FireTimer.Interval = 17
        
        'Hooks are no longer installed at initialization.  You must explicitly request initialization
        ' via the ActivateHook function.  (PhotoDemon does not do this until fairly deep into program startup.)
        
    End If
    
End Sub

'Generally, we prefer PD's main termination code to manually disable us, but we'll always safely release
' the hook at termination time.
Private Sub UserControl_Terminate()
    ReleaseResources
End Sub

'Hook activation/deactivation must be controlled manually by the caller.  This control does *not*
' automatically activate a keyboard hook.
Public Function ActivateHook() As Boolean
    
    If PDMain.IsProgramRunning() Then
        
        'If we're already hooked, don't attempt to hook again
        If (m_HookID = 0) Then
            
            m_HookID = VBHacks.NotifyAcceleratorHookNeeded(Me)
            
            'If this is our first attempt at hooking the keyboard (as part of program startup), note failure.
            ' (To prevent debug log spam, we don't do this on subsequent attempts.)
            If (Not m_SubsequentInitialization) Then
                If (m_HookID = 0) Then PDDebug.LogAction "WARNING!  pdAccelerator.ActivateHook failed.   Hotkeys disabled for this session."
                m_SubsequentInitialization = True
            End If
            
            ActivateHook = (m_HookID = 0)
            
        End If
        
    End If
    
End Function

'PD always prefers to release hooks safely.  At program termination, however, we'll forcibly release
' the keyboard hook to avoid leaks.  Do *NOT* forcibly release the hook at any other time.  Always call
' the SafelyReleaseHook function instead.
Public Sub DeactivateHook(Optional ByVal forciblyReleaseNow As Boolean = True)
    
    If (m_HookID <> 0) Then
        
        If forciblyReleaseNow Then
            VBHacks.NotifyAcceleratorHookNotNeeded ObjPtr(Me)
            UnhookWindowsHookEx m_HookID
            m_HookID = 0
        Else
            SafelyReleaseHook
        End If
        
    End If
    
End Sub

'When PD gains focus, call this function to update all key state tracking.  (This addresses the rare case
' where the user is *already* holding a key down when PD is activated - we can then use that down key as
' part of a subsequent hotkey combo.)
Private Sub RecaptureKeyStates()
    m_CtrlDown = IsVirtualKeyDown(VK_CONTROL)
    m_AltDown = IsVirtualKeyDown(VK_ALT)
    m_ShiftDown = IsVirtualKeyDown(VK_SHIFT)
End Sub

'With some keys (e.g. ALT), PD's main canvas sometimes has to "eat" a keypress to prevent the system
' from doing something unwanted with it (e.g. stealing focus and giving it to the menu bar).  When this
' happens, that window *must* notify us of the new key state, because we won't be able to track it
' (since the window "ate" it).
Public Sub NotifyAltKeystateChange(ByVal newState As Boolean)
    m_AltDown = newState
End Sub

'Dummy theme function; this control isn't visible, so theming isn't relevant - but having a bare sub spares
' pointless errors inside PD's central theme engine.
Public Sub UpdateAgainstCurrentTheme(Optional ByVal hostFormhWnd As Long = 0)

End Sub

'VB exposes a UserControl.EventsFrozen property to check for IDE breaks, but it isn't always reliable.
' We manually check a few other states to avoid hooking the keyboard when we shouldn't.
Private Function AreEventsFrozen() As Boolean
    
    On Error GoTo EventStateCheckError
    
    If UserControl.Enabled Then
        If PDMain.IsProgramRunning() Then
            AreEventsFrozen = UserControl.EventsFrozen
        Else
            AreEventsFrozen = True
        End If
    Else
        AreEventsFrozen = True
    End If
    
    Exit Function

'If an error occurs, assume events are frozen
EventStateCheckError:
    AreEventsFrozen = True
    
End Function

'Returns: TRUE if hotkeys are allowed to accumulate.
Private Function CanIAccumulateAnAccelerator() As Boolean
    CanIAccumulateAnAccelerator = (Not Interface.IsModalDialogActive())
End Function

'Want to globally disable hotkeys?  Call this function.  (And if you add a state to PhotoDemon where
' hotkeys shouldn't be available, add a check for that state here.)
Private Function CanIRaiseAnAcceleratorEvent(Optional ByVal ignoreActiveTimer As Boolean = False) As Boolean
   
    'By default, assume we can raise accelerator events
    CanIRaiseAnAcceleratorEvent = True
    
    'I'm not entirely sure how VB's message pumps work when WM_TIMER events hit disabled controls,
    ' so just to be safe, let's be paranoid and ensure this control hasn't been externally deactivated.
    If (Me.Enabled And (Hotkeys.GetNumOfHotkeys() > 0)) Then
        
        'Don't process accelerators when the main form is disabled (e.g. if a modal dialog is active,
        ' or if a previous action is still executing)
        If FormMain.Enabled Then
            
            'If the accelerator timer is already waiting to process an existing accelerator, exit.
            ' (We'll get a chance to try again on the next timer event.)
            If (m_FireTimer Is Nothing) Then
                CanIRaiseAnAcceleratorEvent = False
            Else
                
                'If the timer is active, let it finish its current task before we attempt to raise another accelerator
                If (Not ignoreActiveTimer) And m_FireTimer.IsActive Then CanIRaiseAnAcceleratorEvent = False
                If m_InFireTimerNow Then CanIRaiseAnAcceleratorEvent = False
                
            End If
        
        '/FormMain disabled
        Else
            CanIRaiseAnAcceleratorEvent = False
        End If
            
        'As one additional failsafe, see if PD is shutting down.  If it is, ignore hotkeys.
        If g_ProgramShuttingDown Then CanIRaiseAnAcceleratorEvent = False
    
    'If this control is disabled or no hotkeys have been loaded (a potential possibility in future builds,
    ' when the user can specify custom hotkey mappings), prevent further processing.
    Else
        CanIRaiseAnAcceleratorEvent = False
    End If
    
End Function

Private Function HandleActualKeypress(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Boolean
    
    'Returning TRUE means we handled the keypress and do *not* want to forward it down the chain
    HandleActualKeypress = False
    If (Not Me.Enabled) Then Exit Function
    
    'Translate modifier states (shift, control, alt/menu) to their masked VB equivalent
    RecaptureKeyStates
    
    Dim retShiftConstants As ShiftConstants
    If m_CtrlDown Then retShiftConstants = retShiftConstants Or vbCtrlMask
    If m_AltDown Then retShiftConstants = retShiftConstants Or vbAltMask
    If m_ShiftDown Then retShiftConstants = retShiftConstants Or vbShiftMask
    
    'Search our accelerator database for a match to the current keycode (only if a menu doesn't have focus)
    If (Hotkeys.GetNumOfHotkeys() > 0) Then
    If (Not Menus.IsMainMenuActive()) Then
        
        'Remap some virtual key IDs to more common equivalents.  For example, Windows will return different
        ' virtual keys (as it should) for the + key next to backspace vs the + key on your number pad.
        ' For the purposes of hotkeys, PD currently treats these as equivalent (though I reserve the right
        ' to revisit this in the future if people complain lol).  Anyway, remapping them here spares the
        ' hotkey manager from needing to double-add hotkeys that are duplicated by the number pad.
        If (wParam = VK_OEM_PLUS) Then wParam = vbKeyAdd
        If (wParam = VK_OEM_MINUS) Then wParam = vbKeySubtract
        'TODO: other numpad variants?
        
        'See if the keycode matches an entry in the hotkey collection
        Dim idxHotkey As Long
        idxHotkey = Hotkeys.GetHotkeyIndex(wParam, retShiftConstants)
        If (idxHotkey >= 0) Then
            
            'We have a match!
            
            'We have one last check to perform before firing this accelerator.  Users with accessibility
            ' constraints (including elderly users) may press-and-hold accelerators long enough to trigger
            ' repeat occurrences.  Accelerators should require full "release key and press again" behavior
            ' to avoid double-firing their associated events.  We handle this by looking for back-to-back
            ' presses of the same hotkey, and enforcing the system keyboard delay between presses.
            If (idxHotkey = m_LastHotkeyIndex) Then
                If (VBHacks.GetTimerDifferenceNow(m_TimerAtAcceleratorPress) < Interface.GetKeyboardRepeatRate()) Then Exit Function
            End If
            
            'Update the current time tracker
            VBHacks.GetHighResTime m_TimerAtAcceleratorPress
            
            'If we're already processing other hotkeys (or we're currently barred from raising hotkey events),
            ' store this hotkey in our running collection - we'll handle it later.
            If (m_InFireTimerNow Or (Not CanIRaiseAnAcceleratorEvent)) Then
                m_AcceleratorAccumulator.Add idxHotkey
            
            'Otherwise, add this hotkey to the live accelerator processing queue and attempt to handle
            ' it immediately.
            Else
                m_AcceleratorQueue.Add idxHotkey
                If (Not m_FireTimer Is Nothing) Then m_FireTimer.StartTimer
            End If
            
            'Return TRUE to notify the hook event that we've eaten this keystroke.  (This prevents
            ' other listeners in the chain from getting this key event.)
            HandleActualKeypress = True
            m_LastHotkeyIndex = idxHotkey
        
        '/no matching hotkey found
        End If
        
    '/no hotkeys in collection, and main menu (or system menu) has not stolen focus
    End If
    End If
    
End Function

Private Function UpdateCtrlAltShiftState(ByVal wParam As Long, ByVal lParam As Long) As Boolean
    
    UpdateCtrlAltShiftState = False
    
    If (wParam = VK_CONTROL) Then
        m_CtrlDown = (lParam >= 0)
        UpdateCtrlAltShiftState = True
    ElseIf (wParam = VK_ALT) Then
        m_AltDown = (lParam >= 0)
        UpdateCtrlAltShiftState = True
    ElseIf (wParam = VK_SHIFT) Then
        m_ShiftDown = (lParam >= 0)
        UpdateCtrlAltShiftState = True
    End If
    
End Function

Friend Function KeyboardHookProcAccelerator(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
    m_InHookNow = True
    On Error GoTo HookProcError
    
    Dim msgEaten As Boolean: msgEaten = False
    
    'Try to see if we're in an IDE break mode.  This isn't 100% reliable, but it's better than not checking at all.
    If (Not AreEventsFrozen) Then
        
        'MSDN states that negative codes must be passed to the next hook, without processing
        ' (see http://msdn.microsoft.com/en-us/library/ms644984.aspx).  Similarly, hooks passed with
        ' the code "3" mean that this is not an actual key event, but one triggered by a PeekMessage()
        ' call with PM_NOREMOVE specified.  We can ignore such peeks and only deal with actual key events.
        If (nCode = 0) Then
            
            'Key hook callbacks can be raised under a variety of conditions.  To ensure we only track actual
            ' "key down" or "key up" events, let's compare transition and previous states.  Because hotkeys
            ' are (by design) not triggered by hold-to-repeat behavior, we only want to deal with key events
            ' that are full transitions from "Unpressed" to "Pressed" or vice-versa.  (The byte masks here
            ' all come from MSDN - check the link above for details!)
            '
            '(Update 2024: some hotkeys (like brush size up/down) actually benefit from key repeat behavior.
            ' To enable this, I removed the old transition state check.  If for some reason we need to enable
            ' it in the future, use the boolean calculation below.)
            'Dim keyTransitionState As Boolean
            'keyTransitionState = ((lParam >= 0) And ((lParam And &H40000000) = 0)) Or ((lParam < 0) And ((lParam And &H40000000) <> 0))
            
            'We now want to check two things simultaneously.  First, we want to update Ctrl/Alt/Shift
            ' key state tracking.  (This is handled by a separate function.)  If something other than
            ' Ctrl/Alt/Shift was pressed, *and* this is a keydown event, let's look for hotkey matches.
            '
            '(How do we detect keydown vs keyup events?  The first bit (e.g. "bit 31" per MSDN) of lParam
            ' defines key state: 0 means the key is being pressed, 1 means the key is being released.
            ' Note the similarity to the transition check, above.)
            If (lParam >= 0) And (Not UpdateCtrlAltShiftState(wParam, lParam)) Then
            
                'Before proceeding with further checks, see if PD is even allowed to process accelerators
                ' in its current state (e.g. if a modal dialog is active, we don't want to raise events)
                If CanIAccumulateAnAccelerator Then
                    
                    'All checks have passed.  We'll handle the actual keycode evaluation matching in another function.
                    msgEaten = HandleActualKeypress(nCode, wParam, lParam)
                    
                End If
            
            '/Ctrl/Alt/Shift was pressed
            End If
            
        '/nCode negative
        End If
        
    'Events frozen
    End If
    
    'If we didn't handle this keypress, allow subsequent hooks to have their way with it
    If (Not msgEaten) Then
        KeyboardHookProcAccelerator = CallNextHookEx(0, nCode, wParam, lParam)
    Else
        KeyboardHookProcAccelerator = 1
    End If
    
    m_InHookNow = False
    Exit Function
    
'On errors, we simply want to bail, as there's little we can safely do to address an error from inside the hooking procedure
HookProcError:
    KeyboardHookProcAccelerator = CallNextHookEx(0, nCode, wParam, lParam)
    m_InHookNow = False
    
End Function
