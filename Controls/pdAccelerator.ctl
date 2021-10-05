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
'Copyright 2013-2021 by Tanner Helland and contributors
'Created: 06/November/15 (split off from a heavily modified vbaIHookControl by Steve McMahon)
'Last updated: 04/October/21
'Last update: migrate all hotkey management to a dedicated module in preparation for customizable
'             hotkeys.  From now on, this control will only be responsible for keyboard hooking/tracking.
'             Managing the underlying hotkey collection is now the responsibility of the Hotkeys module.
'
'For many years, PD used a "hook control" by vbAccelerator.com to handle program hotkeys:
' http://www.vbaccelerator.com/home/VB/Code/Libraries/Hooks/Accelerator_Control/article.asp
'
'In 2013 (https://github.com/tannerhelland/PhotoDemon/commit/373882e452201bb00584a52a791236e05bc97c1e)
' I rewrote the control to solve some glaring stability issues.  Over time, I rewrote it more and more
' (https://github.com/tannerhelland/PhotoDemon/commits/master/Controls/vbalHookControl.ctl),
' tacking on PhotoDemon-specific features and attempting to fix problematic bugs, until ultimately the
' control became a horrible mishmash of spaghetti code: some old, some new, some completely unused,
' and some that was still problematic and unreliable.
'
'Because dynamic hooking has enormous potential for causing hard-to-replicate bugs, a ground-up rewrite
' seemed long overdue.  Thus this control was born.
'
'Thank you to Steve McMahon for his original implementation, which was my first introduction to hooking
' from VB6.  It's still a useful reference for beginners, and you can find the original here
' (hopefully... Steve's work has intermittently disappeared from the web in recent years):
' http://www.vbaccelerator.com/home/VB/Code/Libraries/Hooks/Accelerator_Control/
'
'Thank you also to Jason Brown (https://github.com/jpbro), who has submitted many fixes and
' improvements to this module over the years.  Hotkey behavior in PhotoDemon is greatly improved
' thanks to him.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************


Option Explicit

'This control only raises a single event: "Accelerator", and it only does it when one (or more) keys in the
' combination are released.  (At present, the only object in PD that consumes these events is FormMain.)
Public Event HotkeyPressed(ByVal hotkeyID As Long)

'New solution!  Virtual-key tracking is a bad idea, because we want to know key state at the time the hotkey
' was pressed (not what it is right now).  Solving this is as easy as tracking key up/down state for Ctrl/Alt/Shift
' presses and storing the results locally - but note that this does require some extra checking for things
' like Alt+Tab keypresses.  (See https://github.com/tannerhelland/PhotoDemon/issues/267 for more details.)
Private m_CtrlDown As Boolean, m_AltDown As Boolean, m_ShiftDown As Boolean

'If the control's hook proc is active and primed, this will be set to TRUE.  (HookID is the actual Windows hook handle.)
Private m_HookingActive As Boolean, m_HookID As Long
Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long

'When the control is actually inside the hook procedure, this will be set to TRUE.  The hook *cannot be removed
' until this returns to FALSE*.  To ensure correct unhooking behavior, we use a timer failsafe.
Private m_InHookNow As Boolean
Private m_InFireTimerNow As Boolean

'When PD loses and then gains focus, we need to manually update our control key tracking.  This is done
' by manually checking key state (instead of waiting for a hook event, which we may have missed as
' PD wasn't active!).
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

'Keyboard accelerators are troublesome to handle because they interfere with PD's dynamic hooking solution for
' various canvas and control-related key events.  To work around this limitation, module-level variables are set
' by the accelerator hook control any time a potential accelerator is intercepted.  The hook then initiates an
' internal timer and immediately exits, which allows the keyboard hook proc to safely exit.  When the timer
' finishes enforcing a slight delay, we then perform the actual accelerator evaluation.
Private m_AcceleratorIndex As Long, m_TimerAtAcceleratorPress As Currency

'To reduce the potential for double-fired keys, we track the last-fired accelerator ID.  The current system
' keyboard delay must pass before we fire the same accelerator a second time.
Private m_LastHotkeyIndex As Long

'This control may be problematic on systems with system-wide custom key handlers (like some Intel systems, argh).
' As part of the debug process, we generate extra text on first activation - text that can be ignored on subsequent runs.
Private m_SubsequentInitialization As Boolean

'In-memory timers are used for firing accelerators and releasing hooks
Private WithEvents m_ReleaseTimer As pdTimer
Attribute m_ReleaseTimer.VB_VarHelpID = -1
Private WithEvents m_FireTimer As pdTimer
Attribute m_FireTimer.VB_VarHelpID = -1

'Thanks to a patch by jpbro (https://github.com/tannerhelland/PhotoDemon/pull/248), PD no longer drops accelerators
' that are triggered in quick succession.  Instead, it queues them and fires them in turn.
Private m_AcceleratorQueue As VBA.Collection        'Active queue of accelerators for which events are currently to be raised
Private m_AcceleratorAccumulator As VBA.Collection  'Queue of accelerators which are accumulating while the active queue is being processed

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

Private Sub m_FireTimer_Timer()
    
    Dim i As Long
    
    'If we're still inside the hookproc, wait another 16 ms before testing the keypress.
    If (Not m_InHookNow) Then
    
         If (Not CanIRaiseAnAcceleratorEvent(True)) Then
            
            'We are not currently allowed to raise any events, so short-circuit
            ' (If the program is shutting down, forcibly stop the timer so we don't raise hotkey events again)
            If g_ProgramShuttingDown Then m_FireTimer.StopTimer
            Exit Sub
            
         End If
        
         'Because the accelerator has now been processed, we can disable the timer; this will prevent it from firing again,
         ' but the current sub will still complete its actions.
         m_InFireTimerNow = True ' Notify other methods that we are busy in the timer
         m_FireTimer.StopTimer
         
         'Process accelerators in the active queue in FIFO order
         For i = 1 To m_AcceleratorQueue.Count
         
             m_AcceleratorIndex = m_AcceleratorQueue.Item(i)
         
             If (m_AcceleratorIndex <> -1) Then
                PDDebug.LogAction "raising accelerator-based event (#" & CStr(m_AcceleratorIndex) & ", " & HotKeyName(m_AcceleratorIndex) & ")"
                RaiseEvent HotkeyPressed(m_AcceleratorIndex)
                m_AcceleratorIndex = -1
             End If
             
         Next i
         
         'Swap the active queue for the accumuator queue and empty the old accumulator queue object
         Set m_AcceleratorQueue = m_AcceleratorAccumulator
         Set m_AcceleratorAccumulator = New VBA.Collection
         
         'If we have accumulated accelerators that are now active, restart the timer
         If (m_AcceleratorQueue.Count > 0) Then m_FireTimer.StartTimer
         
         m_InFireTimerNow = False   'Clear the "busy in timer" flag
        
    End If
    
End Sub

Private Sub m_ReleaseTimer_Timer()
    If m_HookingActive Then
        SafelyReleaseHook
    Else
        m_ReleaseTimer.StopTimer
    End If
End Sub

'Hooks cannot be released while actually inside the hookproc.  Call this function to safely release a hook, even from within a hookproc.
Private Sub SafelyReleaseHook()
    
    If (Not PDMain.IsProgramRunning()) Then Exit Sub
    
    'If we're still inside the hook, activate the failsafe timer release mechanism
    If m_InHookNow Then
        If (Not m_ReleaseTimer Is Nothing) Then
            If (Not m_ReleaseTimer.IsActive) Then m_ReleaseTimer.StartTimer
        End If
        
    'If we're not inside the hook, this is a perfect time to release.
    Else
        
        If m_HookingActive Then
            m_HookingActive = False
            If (m_HookID <> 0) Then UnhookWindowsHookEx m_HookID
            m_HookID = 0
            VBHacks.NotifyAcceleratorHookNotNeeded ObjPtr(Me)
        End If
        
        'Also deactivate the failsafe timer
        If (Not m_ReleaseTimer Is Nothing) Then m_ReleaseTimer.StopTimer
        
    End If
    
End Sub

'Prior to shutdown, you can call this function to forcibly release as many accelerator resources as we can.  In PD,
' we use this to free our menu references.
Public Sub ReleaseResources()
    If Not (m_ReleaseTimer Is Nothing) Then Set m_ReleaseTimer = Nothing
    If Not (m_FireTimer Is Nothing) Then Set m_FireTimer = Nothing
End Sub

Private Sub UserControl_Initialize()

    Set m_AcceleratorQueue = New VBA.Collection
    Set m_AcceleratorAccumulator = New VBA.Collection
    
    m_HookingActive = False
    m_AcceleratorIndex = -1
        
    'You may want to consider straight-up disabling hotkeys inside the IDE
    If PDMain.IsProgramRunning() Then
        
        'UI-related timers run at 60 fps
        Set m_ReleaseTimer = New pdTimer
        m_ReleaseTimer.Interval = 17
        
        Set m_FireTimer = New pdTimer
        m_FireTimer.Interval = 17
        
        'Hooks are no longer installed at initialization.  The program must explicitly request initialization
        ' (via the ActivateHook function).
        
    End If
    
End Sub

Private Sub UserControl_Terminate()
    
    'Generally, we prefer the caller to disable us manually, but as a last resort, check for termination at shutdown time.
    If (m_HookID <> 0) Then DeactivateHook True
    
    ReleaseResources
    
End Sub

'Hook activation/deactivation must be controlled manually by the caller
Public Function ActivateHook() As Boolean
    
    If PDMain.IsProgramRunning() Then
        
        'If we're already hooked, don't attempt to hook again
        If (Not m_HookingActive) Then
            
            m_HookID = VBHacks.NotifyAcceleratorHookNeeded(Me)
            m_HookingActive = (m_HookID <> 0)
            
            If (Not m_SubsequentInitialization) Then
                If (Not m_HookingActive) Then PDDebug.LogAction "WARNING!  pdAccelerator.ActivateHook failed.   Hotkeys disabled for this session."
            End If
            m_SubsequentInitialization = True
            
            ActivateHook = m_HookingActive
            
        End If
        
    End If
    
End Function

Public Sub DeactivateHook(Optional ByVal forciblyReleaseInstantly As Boolean = True)
    
    If m_HookingActive Then
        
        If forciblyReleaseInstantly Then
            m_HookingActive = False
            VBHacks.NotifyAcceleratorHookNotNeeded ObjPtr(Me)
            If (m_HookID <> 0) Then UnhookWindowsHookEx m_HookID
            m_HookID = 0
        Else
            SafelyReleaseHook
        End If
        
    End If
    
End Sub

'When PD gains focus, call this function to update all key state tracking.  (This addresses the rare case
' where the user is *already* holding a key down when PD is activated - we can then use that down key as
' part of a subsequent hotkey combo.)
Public Sub RecaptureKeyStates()
    m_CtrlDown = IsVirtualKeyDown(VK_CONTROL)
    m_AltDown = IsVirtualKeyDown(VK_ALT)
    m_ShiftDown = IsVirtualKeyDown(VK_SHIFT)
End Sub

'Note that the vKey constant below is a virtual key mapping, not necessarily a standard VB key constant
Private Function IsVirtualKeyDown(ByVal vKey As Long) As Boolean
    IsVirtualKeyDown = GetAsyncKeyState(vKey) And &H8000
End Function

'When PD loses focus, call this function to reset all key state tracking
Public Sub ResetKeyStates()
    m_CtrlDown = False
    m_AltDown = False
    m_ShiftDown = False
End Sub

'With some keys (e.g. ALT), PD's main canvas sometimes has to "eat" a keypress to prevent the system
' from doing something unwanted with it (e.g. stealing focus and giving it to the menu bar).  When this
' happens, that window *must* notify us of the new key state, because we won't be able to track it
' (since the window "ate" it).
Public Sub NotifyAltKeystateChange(ByVal newState As Boolean)
    m_AltDown = newState
End Sub

'Dummy theme function; this control isn't visible, so theming isn't relevant - but having a bare sub spares
' pointless errors inside the central theme engine.
Public Sub UpdateAgainstCurrentTheme(Optional ByVal hostFormhWnd As Long = 0)

End Sub

'VB exposes a UserControl.EventsFrozen property to check for IDE breaks, but in my testing it isn't reliable.
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

'Want to globally disable accelerators under certain circumstances?  Add code here to do it.
Private Function CanIRaiseAnAcceleratorEvent(Optional ByVal ignoreActiveTimer As Boolean = False) As Boolean
   
    'By default, assume we can raise accelerator events
    CanIRaiseAnAcceleratorEvent = True
    
    'I'm not entirely sure how VB's message pumps work when WM_TIMER events hit disabled controls, so just to be safe,
    ' let's be paranoid and ensure this control hasn't been externally deactivated.
    If (Me.Enabled And (Hotkeys.GetNumOfHotkeys() > 0)) Then
        
        'Don't process accelerators when the main form is disabled (e.g. if a modal form is present, or if a previous
        ' action is in the middle of execution)
        If (Not FormMain.Enabled) Then CanIRaiseAnAcceleratorEvent = False
        
        'If the accelerator timer is already waiting to process an existing accelerator, exit.  (We'll get a chance to
        ' try again on the next timer event.)
        If (m_FireTimer Is Nothing) Then
            CanIRaiseAnAcceleratorEvent = False
        Else
            
            'If the timer is active, let it finish its current task before we attempt to raise another accelerator
            If (Not ignoreActiveTimer) And m_FireTimer.IsActive Then CanIRaiseAnAcceleratorEvent = False
            If m_InFireTimerNow Then CanIRaiseAnAcceleratorEvent = False
            
        End If
        
        'If PD is shutting down, we obviously want to ignore accelerators entirely
        If g_ProgramShuttingDown Then CanIRaiseAnAcceleratorEvent = False
    
    'If this control is disabled or no hotkeys have been loaded (a potential possibility in future builds, when the
    ' user will have control over custom hotkeys), save some CPU cycles and prevent further processing.
    Else
        CanIRaiseAnAcceleratorEvent = False
    End If
    
End Function

Private Function HandleActualKeypress(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal p_AccumulateOnly As Boolean) As Boolean
    
    'Translate modifier states (shift, control, alt/menu) to their masked VB equivalent
    Dim retShiftConstants As ShiftConstants
    If m_CtrlDown Then retShiftConstants = retShiftConstants Or vbCtrlMask
    If m_AltDown Then retShiftConstants = retShiftConstants Or vbAltMask
    If m_ShiftDown Then retShiftConstants = retShiftConstants Or vbShiftMask
    
    'Search our accelerator database for a match to the current keycode
    If (Hotkeys.GetNumOfHotkeys() > 0) Then
        
        'See if the keycode matches an entry in the hotkey collection
        Dim idxHotkey As Long
        idxHotkey = Hotkeys.GetHotkeyIndex(wParam, retShiftConstants)
            
        If (idxHotkey >= 0) Then
            
            'We have a match!
            
            'We have one last check to perform before firing this accelerator.  Users with accessibility constraints
            ' (including elderly users) may press-and-hold accelerators long enough to trigger repeat occurrences.
            ' Accelerators should require full "release key and press again" behavior to avoid double-firing
            ' their associated events.
            If (idxHotkey = m_LastHotkeyIndex) Then
                If (VBHacks.GetTimerDifferenceNow(m_TimerAtAcceleratorPress) < Interface.GetKeyboardDelay()) Then Exit Function
            End If
            
            m_AcceleratorIndex = idxHotkey
            VBHacks.GetHighResTime m_TimerAtAcceleratorPress
    
            If p_AccumulateOnly Then
                'Add to accelerator accumulator, it will be processed later.
                m_AcceleratorAccumulator.Add m_AcceleratorIndex
                
            Else
                'Add to the live accelerator processing queue
                m_AcceleratorQueue.Add m_AcceleratorIndex
                If (Not m_FireTimer Is Nothing) Then m_FireTimer.StartTimer
            End If
            
            'Also, make sure to eat this keystroke
            HandleActualKeypress = True
            m_LastHotkeyIndex = idxHotkey
            
        End If
        
    End If  'Hotkey collection exists
    
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
        ' (see http://msdn.microsoft.com/en-us/library/ms644984.aspx).  Similarly, hooks passed with the code "3"
        ' mean that this is not an actual key event, but one triggered by a PeekMessage() call with PM_NOREMOVE specified.
        ' We can ignore such peeks and only deal with actual key events.
        If (nCode = 0) Then
            
            'Key hook callbacks can be raised under a variety of conditions.  To ensure we only track actual "key down"
            ' or "key up" events, let's compare transition and previous states.  Because hotkeys are (by design) not
            ' triggered by hold-to-repeat behavior, we only want to deal with key events that are full transitions from
            ' "Unpressed" to "Pressed" or vice-versa.  (The byte masks here all come from MSDN - check the link above
            ' for details!)
            If ((lParam >= 0) And ((lParam And &H40000000) = 0)) Or ((lParam < 0) And ((lParam And &H40000000) <> 0)) Then
                
                'We now want to check two things simultaneously.  First, we want to update Ctrl/Alt/Shift key state tracking.
                ' (This is handled by a separate function.)  If something other than Ctrl/Alt/Shift was pressed, *and* this is
                ' a keydown event, let's process the key for hotkey matches.
                
                '(How do we detect keydown vs keyup events?  The first bit (e.g. "bit 31" per MSDN) of lParam defines key state:
                ' 0 means the key is being pressed, 1 means the key is being released.  Note the similarity to the transition
                ' check, above.)
                If (lParam >= 0) And (Not UpdateCtrlAltShiftState(wParam, lParam)) Then
                
                    'Before proceeding with further checks, see if PD is even allowed to process accelerators in its
                    ' current state (e.g. if a modal dialog is active, we don't want to raise events)
                    If CanIAccumulateAnAccelerator Then
                    
                        'All checks have passed.  We'll handle the actual keycode evaluation matching in another function.
                        msgEaten = HandleActualKeypress(nCode, wParam, lParam, m_InFireTimerNow Or (Not CanIRaiseAnAcceleratorEvent))
                        
                    End If
                    
                End If  'Key is not in a transitionary state
                
            End If  'Key other than Ctrl/Alt/Shift was pressed
            
        End If  'nCode is not negative
        
    End If  'Events are not frozen
    
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

