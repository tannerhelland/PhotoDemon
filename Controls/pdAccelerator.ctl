VERSION 5.00
Begin VB.UserControl pdAccelerator 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ClipBehavior    =   0  'None
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
   Begin VB.Timer tmrRelease 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   0
      Top             =   480
   End
   Begin VB.Timer tmrAccelerator 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "pdAccelerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Accelerator ("Hotkey") handler
'Copyright 2013-2015 by Tanner Helland
'Created: 06/November/15 (formally split off from a heavily modified vbaIHookControl by Steve McMahon
'Last updated: 06/November/15
'Last update: rewrite the damn thing (mostly) from scratch
'
'For many years, PD used vbAccelerator's "hook control" to handle program hotkeys:
' http://www.vbaccelerator.com/home/VB/Code/Libraries/Hooks/Accelerator_Control/article.asp
'
'Starting in August 2013 (https://github.com/tannerhelland/PhotoDemon/commit/373882e452201bb00584a52a791236e05bc97c1e),
' I rewrote much of the control to solve some glaring stability issues.  Over time, I rewrote it more and more
' (https://github.com/tannerhelland/PhotoDemon/commits/master/Controls/vbalHookControl.ctl), tacking on PD-specific
' features and attempting to fix problematic bugs, until ultimately the control became a horrible mishmash of
' spaghetti code: some old, some new, some completely unused, and some that was still problematically unreliable.
'
'Because dynamic hooking has enormous potential for causing hard-to-replicate bugs, a ground-up rewrite seemed long
' overdue.  Hence this new control.
'
'Many, many thanks to Steve McMahon for his original implementation, which was my first introduction to hooking
' from VB6.  It's still a fine reference for beginners, and you can find the original here (good as of November '15):
' http://www.vbaccelerator.com/home/VB/Code/Libraries/Hooks/Accelerator_Control/article.asp
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************


Option Explicit

'This control only raises a single "Accelerator" event, and it only does it when one (or more) keys in the combination are released
Public Event Accelerator(ByVal acceleratorIndex As Long)

'TODO: use alternate, PD-specific mechanism for determining if FormMain has focus.
'Private Declare Function GetActiveWindow Lib "user32" () As Long

'Key state can be retrieved directly from the hook messages, but it's actually easier to dynamically query the API
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

'Each hotkey stores several additional (and sometimes optional) parameters.  This spares us from writing specialized
' handling code for each individual keypress.
Private Type pdHotkey
    vKeyCode As Long
    shiftState As ShiftConstants
    HotKeyName As String
    IsProcessorString As Boolean
    requiresOpenImage As Boolean
    showProcDialog As Boolean
    procUndo As PD_UNDO_TYPE
    relevantMenu As Menu
End Type

'The list of hotkeys is stored in a basic array.  This makes it easy to set/retrieve values using built-in VB functions,
' and because the list of keys is short, performance isn't in issue.
Private m_Hotkeys() As pdHotkey
Private m_NumOfHotkeys As Long
Private Const INITIAL_HOTKEY_LIST_SIZE As Long = 16&

'In some places, virtual key-codes are used to retrieve key states
Private Const VK_SHIFT As Long = &H10
Private Const VK_CONTROL As Long = &H11
Private Const VK_ALT As Long = &H12    'Note that VK_ALT is referred to as VK_MENU in MSDN documentation!

'If the control's hook proc is active and primed, this will be set to TRUE
Private m_HookingActive As Boolean

'When the control is actually inside the hook procedure, this will be set to TRUE.  The hook *cannot be removed
' until this returns to FALSE*.  To ensure correct unhooking behavior, we use a timer failsafe.
Private m_InHookNow As Boolean

'Keyboard accelerators are troublesome to handle because they interfere with PD's dynamic hooking solution for canvas hotkeys.  To work around this
' limitation, these module-level variables are set by the accelerator hook control any time a potential accelerator is intercepted.  The hook then
' initiates the tmrAccelerator timer and immediately exits, which allows the hookproc to safely exit.  After the timer enforces a slight delay,
' it then performs the actual accelerator evaluation.
Private m_AcceleratorIndex As Long, m_TimerAtAcceleratorPress As Double

'Dynamic hooking requires great care, particularly within the IDE.  PD makes all attempts to do it safely.
Private m_Subclass As cSelfSubHookCallback

'The Enabled property is a bit unique; see http://msdn.microsoft.com/en-us/library/aa261357%28v=vs.60%29.aspx
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_UserMemId = -514
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal newValue As Boolean)
    UserControl.Enabled = newValue
    PropertyChanged "Enabled"
End Property

Private Sub tmrAccelerator_Timer()
    
    'If we're still inside the hookproc, wait another 100 ms before testing the keypress
    If (Not m_InHookNow) Then
        
        'To prevent multiple events from firing too closely together, enforce a slight action delay before processing
        If Timer - m_TimerAtAcceleratorPress > 0.05 Then
    
            'Because the accelerator has now been processed, we can disable the timer; this will prevent it from firing again, but the
            ' current sub will still complete its actions.
            tmrAccelerator.Enabled = False
            
            'If the accelerator index is valid, raise a corresponding event, then reset the accelerator index
            If m_AcceleratorIndex <> -1 Then
                RaiseEvent Accelerator(m_AcceleratorIndex)
                m_AcceleratorIndex = -1
            End If
            
        End If
        
    End If
    
End Sub

Private Sub tmrRelease_Timer()
    
    If m_HookingActive Then
        SafelyReleaseHook
    Else
        tmrRelease.Enabled = False
    End If
    
End Sub

'Hooks cannot be released while actually inside the hookproc.  Call this function to safely release a hook, even from within a hookproc.
Private Sub SafelyReleaseHook()
    
    'If we're still inside the hook, activate the failsafe timer release mechanism
    If m_InHookNow Then
        If Not tmrRelease.Enabled Then tmrRelease.Enabled = True
        
    'If we're not inside the hook, this is a perfect time to release.
    Else
        
        If m_HookingActive Then
            m_HookingActive = False
            m_Subclass.shk_UnHook WH_KEYBOARD
            m_Subclass.shk_TerminateHooks
        End If
        
        'Also deactivate the failsafe timer
        tmrRelease.Enabled = False
        
    End If
    
End Sub

'Prior to shutdown, you can call this function to forcibly release as many accelerator resources as we can.  In PD,
' we use this to free our menu references.
Public Sub ReleaseResources()
    
    Dim i As Long
    For i = 0 To m_NumOfHotkeys - 1
        Set m_Hotkeys(i).relevantMenu = Nothing
    Next i
    
End Sub

Private Sub UserControl_Initialize()
    
    m_HookingActive = False
    m_AcceleratorIndex = -1
    
    m_NumOfHotkeys = 0
    ReDim m_Hotkeys(0 To INITIAL_HOTKEY_LIST_SIZE - 1) As pdHotkey
        
    'You may want to consider straight-up disabling hotkeys inside the IDE
    If g_IsProgramRunning Then
        
        Set m_Subclass = New cSelfSubHookCallback
        
        'Hooks are not installed at initialization.  The program must explicitly request initialization.
        
    End If
    
End Sub

Private Sub UserControl_Terminate()
    
    'Generally, we prefer the caller to disable us manually, but as a last resort, check for termination at shutdown time.
    If Not (m_Subclass Is Nothing) Then
        DeactivateHook True
        Set m_Subclass = Nothing
    End If
    
    ReleaseResources
    
End Sub

'Hook activation/deactivation must be controlled manually by the caller
Public Function ActivateHook() As Boolean
    
    If Not (m_Subclass Is Nothing) Then
        
        'If we're already hooked, don't attempt to hook again
        If (Not m_HookingActive) Then
        
            m_HookingActive = m_Subclass.shk_SetHook(WH_KEYBOARD, False, MSG_BEFORE, , 1, Me)
            
            #If DEBUGMODE = 1 Then
                If m_HookingActive Then
                    pdDebug.LogAction "pdAccelerator.ActivateHook successful.  Hotkeys enabled for this session."
                Else
                    pdDebug.LogAction "WARNING!  pdAccelerator.ActivateHook failed.   Hotkeys disabled for this session."
                End If
            #End If
            
            ActivateHook = m_HookingActive
            
        End If
        
    End If
    
End Function

Public Sub DeactivateHook(Optional ByVal forciblyReleaseInstantly As Boolean = True)
    
    If (Not (m_Subclass Is Nothing)) And m_HookingActive Then
        
        If forciblyReleaseInstantly Then
            m_HookingActive = False
            m_Subclass.shk_UnHook WH_KEYBOARD
            m_Subclass.shk_TerminateHooks
        Else
            SafelyReleaseHook
        End If
        
    End If
    
End Sub

'Add a new accelerator key combination to the collection.  A ton of PD-specific functionality is included in this function, so let me break it down.
' - "isProcessorString": if TRUE, hotKeyName is assumed to a be a string meant for PD's central processor.  It will be directly passed
'    to the processor there when that hotkey is used.
' - "correspondingMenu": a reference to the menu associated with this hotkey.  The reference is used to dynamically draw matching shortcut text
'    onto the menu.  It is not otherwise used.
' - "requiresOpenImage": specifies that this action *must be disallowed* unless one (or more) image(s) are loaded and active.
' - "showProcForm": controls the "showDialog" parameter of processor string directives.
' - "procUndo": controls the "createUndo" parameter of processor string directives.  Remember that UNDO_NOTHING means "do not create Undo data."
Public Function AddAccelerator(ByVal vKeyCode As KeyCodeConstants, ByVal Shift As ShiftConstants, Optional ByVal HotKeyName As String = vbNullString, Optional ByRef correspondingMenu As Menu = Nothing, Optional ByVal IsProcessorString As Boolean = False, Optional ByVal requiresOpenImage As Boolean = True, Optional ByVal showProcDialog As Boolean = True, Optional ByVal procUndo As PD_UNDO_TYPE = UNDO_NOTHING) As Long
    
    'Make sure this key combination doesn't already exist in the collection
    Dim failsafeCheck As Long
    failsafeCheck = GetAcceleratorIndex(vKeyCode, Shift)
    
    If failsafeCheck >= 0 Then
        AddAccelerator = failsafeCheck
        Exit Function
    End If
    
    'We now know that this key combination is unique.
    
    'Make sure the list is large enough to hold this new entry.
    If m_NumOfHotkeys > UBound(m_Hotkeys) Then ReDim Preserve m_Hotkeys(0 To UBound(m_Hotkeys) * 2 + 1) As pdHotkey
    
    'Add the new entry
    With m_Hotkeys(m_NumOfHotkeys)
        .vKeyCode = vKeyCode
        .shiftState = Shift
        .HotKeyName = HotKeyName
        Set .relevantMenu = correspondingMenu
        .IsProcessorString = IsProcessorString
        .requiresOpenImage = requiresOpenImage
        .showProcDialog = showProcDialog
        .procUndo = procUndo
    End With
    
    'Return this index, and increment the active hotkey count
    AddAccelerator = m_NumOfHotkeys
    m_NumOfHotkeys = m_NumOfHotkeys + 1
    
End Function

'If an accelerator exists in our current collection, this will return a value >= 0 corresponding to its position in the master array.
Private Function GetAcceleratorIndex(ByVal vKeyCode As KeyCodeConstants, ByVal Shift As ShiftConstants) As Long
    
    GetAcceleratorIndex = -1
    
    If m_NumOfHotkeys > 0 Then
        
        Dim i As Long
        For i = 0 To m_NumOfHotkeys - 1
            If (m_Hotkeys(i).vKeyCode = vKeyCode) And (m_Hotkeys(i).shiftState = Shift) Then
                GetAcceleratorIndex = i
                Exit For
            End If
        Next i
        
    End If

End Function

'Outside functions can retrieve certain accelerator properties.  Note that - by design - these properties should only be retrieved from inside
' an Accelerator event.
Public Function Count() As Long
    Count = m_NumOfHotkeys
End Function

Public Function IsProcessorString(ByVal hkIndex As Long) As Boolean
    If (hkIndex >= 0) And (hkIndex < m_NumOfHotkeys) Then
        IsProcessorString = m_Hotkeys(hkIndex).IsProcessorString
    End If
End Function

Public Function IsImageRequired(ByVal hkIndex As Long) As Boolean
    If (hkIndex >= 0) And (hkIndex < m_NumOfHotkeys) Then
        IsImageRequired = m_Hotkeys(hkIndex).requiresOpenImage
    End If
End Function

Public Function IsDialogDisplayed(ByVal hkIndex As Long) As Boolean
    If (hkIndex >= 0) And (hkIndex < m_NumOfHotkeys) Then
        IsDialogDisplayed = m_Hotkeys(hkIndex).showProcDialog
    End If
End Function

Public Function HasMenu(ByVal hkIndex As Long) As Boolean
    If (hkIndex >= 0) And (hkIndex < m_NumOfHotkeys) Then
        HasMenu = CBool(Not (m_Hotkeys(hkIndex).relevantMenu Is Nothing))
    End If
End Function

Public Function HotKeyName(ByVal hkIndex As Long) As String
    If (hkIndex >= 0) And (hkIndex < m_NumOfHotkeys) Then
        HotKeyName = m_Hotkeys(hkIndex).HotKeyName
    End If
End Function

Public Function MenuReference(ByVal hkIndex As Long) As Menu
    If (hkIndex >= 0) And (hkIndex < m_NumOfHotkeys) Then
        Set MenuReference = m_Hotkeys(hkIndex).relevantMenu
    End If
End Function

Public Function ProcUndoValue(ByVal hkIndex As Long) As PD_UNDO_TYPE
    ProcUndoValue = m_Hotkeys(hkIndex).procUndo
End Function

Public Function StringRepresentation(ByVal hkIndex As Long) As String

    If (hkIndex >= 0) And (hkIndex < m_NumOfHotkeys) Then
    
        Dim tmpString As String
        If m_Hotkeys(hkIndex).shiftState And vbCtrlMask Then tmpString = g_Language.TranslateMessage("Ctrl") & "+"
        If m_Hotkeys(hkIndex).shiftState And vbAltMask Then tmpString = tmpString & g_Language.TranslateMessage("Alt") & "+"
        If m_Hotkeys(hkIndex).shiftState And vbShiftMask Then tmpString = tmpString & g_Language.TranslateMessage("Shift") & "+"
        
        'Processing the string itself takes a bit of extra work, as some keyboard keys don't automatically map to a
        ' string equivalent.  (Also, translations need to be considered.)
        Select Case m_Hotkeys(hkIndex).vKeyCode
        
            Case vbKeyAdd
                tmpString = tmpString & "+"
            
            Case vbKeySubtract
                tmpString = tmpString & "-"
            
            Case vbKeyReturn
                tmpString = tmpString & g_Language.TranslateMessage("Enter")
            
            Case vbKeyPageUp
                tmpString = tmpString & g_Language.TranslateMessage("Page Up")
            
            Case vbKeyPageDown
                tmpString = tmpString & g_Language.TranslateMessage("Page Down")
                
            Case vbKeyF1 To vbKeyF16
                tmpString = tmpString & "F" & (CLng(m_Hotkeys(hkIndex).vKeyCode) - 111)
            
            'In the future I would like to enumerate virtual key bindings properly, using the data at this link:
            ' http://msdn.microsoft.com/en-us/library/windows/desktop/dd375731%28v=vs.85%29.aspx
            ' At the moment, however, they're implemented as magic numbers.
            Case 188
                tmpString = tmpString & ","
                
            Case 190
                tmpString = tmpString & "."
                
            Case 219
                tmpString = tmpString & "["
                
            Case 221
                tmpString = tmpString & "]"
                
            Case Else
                tmpString = tmpString & UCase$(ChrW$(m_Hotkeys(hkIndex).vKeyCode))
            
        End Select
        
        StringRepresentation = tmpString
    
    Else
        StringRepresentation = ""
    End If
    
End Function

'VB exposes a UserControl.EventsFrozen property to check for IDE breaks, but in my testing it isn't reliable.
Private Function AreEventsFrozen() As Boolean
    
    On Error GoTo EventStateCheckError
    
    If UserControl.Enabled Then
        If g_IsProgramRunning Then
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

'Note that the vKey constant taken by this function is a *virtual key mapping*.  This may or may not map to a
' standard VB key constant, so use care when calling it.
Private Function IsVirtualKeyDown(ByVal vKey As Long) As Boolean
    IsVirtualKeyDown = GetAsyncKeyState(vKey) And &H8000
End Function

'Want to globally disable accelerators under certain circumstances?  Add code here to do it.
Private Function CanIRaiseAnAcceleratorEvent() As Boolean
    
    'By default, assume we can raise accelerator events
    CanIRaiseAnAcceleratorEvent = True
    
    'Don't process accelerators when the main form is disabled (e.g. if a modal form is present, or if a previous
    ' action is in the middle of execution)
    If Not FormMain.Enabled Then CanIRaiseAnAcceleratorEvent = False
    
    'Don't process accelerators if the Language Editor is active
    If Not (FormLanguageEditor Is Nothing) Then
        If FormLanguageEditor.Visible Then CanIRaiseAnAcceleratorEvent = False
    End If

    'Accelerators can be fired multiple times by accident.  Don't allow the user to press accelerators
    ' faster than the system keyboard delay (250ms at minimum, 1s at maximum).
    If Abs(Timer - m_TimerAtAcceleratorPress < Interface.GetKeyboardDelay()) Then CanIRaiseAnAcceleratorEvent = False
    
    'If the accelerator timer is already waiting to process an existing accelerator, exit
    If tmrAccelerator.Enabled Then CanIRaiseAnAcceleratorEvent = False
    
    'If PD is shutting down, ignore accelerators
    If g_ProgramShuttingDown Then CanIRaiseAnAcceleratorEvent = False
    
    'By this point, the function is set to the proper pass/fail state
    
End Function

'This routine MUST BE KEPT as the final routine for this form. Its ordinal position determines its ability to hook properly.
' Hooking is required to track application-wide mouse presses
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
    
    On Error GoTo HookProcError
    
    m_InHookNow = True
    bHandled = False
    
    'Try to see if we're in an IDE break mode.  This isn't 100% reliable, but it's better than not checking at all.
    If (Not AreEventsFrozen) Then
        
        'MSDN states that negative codes must be passed to the next hook, without processing
        ' (see http://msdn.microsoft.com/en-us/library/ms644984.aspx)
        If nCode >= 0 Then
        
            'Bit 31 of lParam is 0 if a key is being pressed, and 1 if it is being released.  In the future, we could
            ' use this to raise separate KeyDown and KeyUp events, if desired.
            If lParam < 0 Then
                
                'Key up events are raised twice; once in a transitionary stage, and once again in a final stage.
                ' To prevent double-raising of KeyUp events, we check the transitionary state before proceeding
                If ((lParam And 1) <> 0) And ((lParam And 3) = 1) Then
                    
                    'Manually pull key modifier states (shift, control, alt/menu) in advance; these are standard for all key events
                    Dim retShiftConstants As ShiftConstants
                    If IsVirtualKeyDown(VK_SHIFT) Then retShiftConstants = retShiftConstants Or vbShiftMask
                    If IsVirtualKeyDown(VK_CONTROL) Then retShiftConstants = retShiftConstants Or vbCtrlMask
                    If IsVirtualKeyDown(VK_ALT) Then retShiftConstants = retShiftConstants Or vbAltMask
                    
                    'Search our accelerator database for a match to the current keycode
                    If m_NumOfHotkeys > 0 Then
                    
                        Dim i As Long
                        For i = 0 To m_NumOfHotkeys - 1
                            
                            'First, see if the keycode matches.
                            If m_Hotkeys(i).vKeyCode = wParam Then
                            
                                'Next, see if the Ctrl+Alt+Shift state matches
                                If m_Hotkeys(i).shiftState = retShiftConstants Then
                                
                                    'Boom!  This is a hit.  See if there is any reason to *not* raise an accelerator.
                                    ' (In PD, a lot of things can cause disabled accelerators - text boxes having focus,
                                    '  certain modal dialogs, no loaded images, etc.)
                                    If CanIRaiseAnAcceleratorEvent Then
                                    
                                        'No global PD state is preventing us from handling this accelerator.  Cache the index of the accelerator,
                                        ' note the current time, then initiate the accelerator evaluation timer.  It handles all further evaluation.
                                        m_AcceleratorIndex = i
                                        m_TimerAtAcceleratorPress = Timer
                                        tmrAccelerator.Enabled = True
                                        
                                        'Also, make sure to eat this keystroke
                                        bHandled = True
                                    
                                    End If
                                    
                                End If
                            End If
                            
                        Next i
                    
                    End If      'Hotkey collection exists
                End If      'Key is not in a transitionary state
            End If      'Key is being released, not pressed
        End If      'nCode >= 0
    End If      'Events are not frozen
    
    
    'If we didn't handle this keypress, allow subsequent hooks to have their way with it
    If (Not bHandled) Then
        lReturn = CallNextHookEx(0&, nCode, wParam, lParam)
    Else
        lReturn = 1
    End If
    
    m_InHookNow = False
    Exit Sub
    
'On errors, we simply want to bail, as there's little we can safely do to address an error from inside the hooking procedure
HookProcError:
    
    lReturn = CallNextHookEx(0&, nCode, wParam, lParam)
    m_InHookNow = False
    

' *************************************************************
' C A U T I O N   C A U T I O N   C A U T I O N   C A U T I O N
' -------------------------------------------------------------
' DO NOT ADD ANY OTHER CODE BELOW THE "END SUB" STATEMENT BELOW
'   add this warning banner to the last routine in your class
' *************************************************************
End Sub

