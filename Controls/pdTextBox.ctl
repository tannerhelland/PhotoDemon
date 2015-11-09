VERSION 5.00
Begin VB.UserControl pdTextBox 
   Appearance      =   0  'Flat
   BackColor       =   &H0080FF80&
   ClientHeight    =   975
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3015
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   65
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   201
   ToolboxBitmap   =   "pdTextBox.ctx":0000
   Begin VB.Timer tmrHookRelease 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   120
      Top             =   240
   End
End
Attribute VB_Name = "pdTextBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Unicode Text Box control
'Copyright 2014-2015 by Tanner Helland
'Created: 03/November/14
'Last updated: 21/November/14
'Last update: incorporate many fixes from Zhu JY's testing
'
'In a surprise to precisely no one, PhotoDemon has some unique needs when it comes to user controls - needs that
' the intrinsic VB controls can't handle.  These range from the obnoxious (lack of an "autosize" property for
' anything but labels) to the critical (no Unicode support).
'
'As such, I've created many of my own UCs for the program.  All are owner-drawn, with the goal of maintaining
' visual fidelity across the program, while also enabling key features like Unicode support.
'
'A few notes on this text box control, specifically:
'
' 1) Unlike other PD custom controls, this one is simply a wrapper to a system text box.
' 2) The idea with this control was not to expose all text box properties, but simply those most relevant to PD.
' 3) Focus is the real nightmare for this control, and as you will see, some complicated tricks are required to work
'    around VB's handling of tabstops in particular.
' 4) To allow use of arrow keys and other control keys, this control must hook the keyboard.  (If it does not, VB will
'    eat control keypresses, because it doesn't know about windows created via the API!)  A byproduct of this is that
'    accelerators flat-out WILL NOT WORK while this control has focus.  I haven't yet settled on a good way to handle
'    this; what I may end up doing is manually forwarding any key combinations that use Alt to the default window
'    handler, but I'm not sure this will help.  TODO!
' 5) Dynamic hooking can occasionally cause trouble in the IDE, particularly when used with break points.  It should
'    be fine once compiled.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************


Option Explicit

'By design, this textbox raises fewer events than a standard text box
Public Event Change()
Public Event KeyPress(ByVal vKey As Long, ByRef preventFurtherHandling As Boolean)
Public Event Resize()
Public Event GotFocusAPI()
Public Event LostFocusAPI()
Public Event TabPress(ByVal focusDirectionForward As Boolean)

'Window styles
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

Private Const ES_AUTOHSCROLL = &H80
Private Const ES_AUTOVSCROLL = &H40
Private Const ES_CENTER = &H1
Private Const ES_LEFT = &H0
Private Const ES_LOWERCASE = &H10
Private Const ES_MULTILINE = &H4
Private Const ES_NOHIDESEL = &H100
Private Const ES_NUMBER = &H2000
Private Const ES_OEMCONVERT = &H400
Private Const ES_PASSWORD = &H20
Private Const ES_READONLY = &H800
Private Const ES_RIGHT = &H2
Private Const ES_UPPERCASE = &H8
Private Const ES_WANTRETURN = &H1000

'Updating the font is done via WM_SETFONT
Private Const WM_SETFONT = &H30

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

'System window handling APIs
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExW" (ByVal dwExStyle As Long, ByVal lpClassName As Long, ByVal lpWindowName As Long, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByRef lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hndWindow As Long, ByVal nCmdShow As showWindowOptions) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hndWindow As Long) As Long
Private Declare Function SetFocus Lib "user32" (ByVal hndWindow As Long) As Long

'Getting/setting window data
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextW" (ByVal hWnd As Long, ByVal lpStringPointer As Long, ByVal nMaxCount As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthW" (ByVal hWnd As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long

'Handle to the system edit box wrapped by this control
Private m_EditBoxHwnd As Long

'pdFont handles the creation and maintenance of the font used to render the text box.  It is also used to determine control width for
' single-line text boxes, as the control is auto-sized to fit the current font.
Private curFont As pdFont

'Rather than use an StdFont container (which requires VB to create redundant font objects), we track font properties manually,
' via dedicated properties.
Private m_FontSize As Single

'The following property changes require creation/destruction of the text box.  PD will automatically backup the edit box's text
' prior to recreating it, but note that text cannot be non-destructively saved when toggling the multiline property if linefeed
' characters are in use!
Private m_Multiline As Boolean

'Custom subclassing is required for IME support
Private Type winMsg
    hWnd As Long
    sysMsg As Long
    wParam As Long
    lParam As Long
    msgTime As Long
    ptX As Long
    ptY As Long
End Type

'GetKeyboardState fills a [256] array with the state of all keyboard keys.  Rather than constantly redimming an array for holding those
' return values, we simply declare one array at a module level, and re-use it as necessary.
Private Declare Function GetKeyboardState Lib "user32" (ByRef pbKeyState As Byte) As Long
Private m_keyStateData(0 To 255) As Byte
Private m_OverrideDoubleCheck As Boolean

Private Declare Function ToUnicode Lib "user32" (ByVal uVirtKey As Long, ByVal uScanCode As Long, lpKeyState As Byte, ByVal pwszBuff As Long, ByVal cchBuff As Long, ByVal wFlags As Long) As Long

Private Const MAPVK_VK_TO_VSC As Long = &H0
Private Const MAPVK_VK_TO_CHAR As Long = &H2
Private Declare Function MapVirtualKey Lib "user32" Alias "MapVirtualKeyW" (ByVal uCode As Long, ByVal uMapType As Long) As Long

Private Const PM_REMOVE As Long = &H1
Private Const WM_KEYFIRST As Long = &H100
Private Const WM_KEYLAST As Long = &H108
Private Declare Function PeekMessage Lib "user32" Alias "PeekMessageW" (ByRef lpMsg As winMsg, ByVal hWnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long
Private Declare Function DispatchMessage Lib "user32" Alias "DispatchMessageW" (lpMsg As winMsg) As Long

'Alt-key state can be tracked a number of different ways, but I find GetAsyncKeyState to be the easiest.
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Private Const WM_KEYDOWN As Long = &H100
Private Const WM_SYSKEYDOWN As Long = &H104
Private Const WM_KEYUP As Long = &H101
Private Const WM_CHAR As Long = &H102
Private Const WM_UNICHAR As Long = &H109
Private Const UNICODE_NOCHAR As Long = &HFFFF&
Private Const WM_SETFOCUS As Long = &H7
Private Const WM_KILLFOCUS As Long = &H8
Private Const WM_SETTEXT As Long = &HC
Private Const WM_COMMAND As Long = &H111
Private Const WM_NEXTDLGCTL As Long = &H28
Private Const WM_MOUSEACTIVATE As Long = &H21
Private Const WM_CTLCOLOREDIT As Long = &H133

Private Const EN_UPDATE As Long = &H400
Private Const EM_GETSEL As Long = &HB0
Private Const EM_SETSEL As Long = &HB1

Private Const VK_SHIFT As Long = &H10
Private Const VK_CONTROL As Long = &H11
Private Const VK_ALT As Long = &H12    'Note that VK_ALT is referred to as VK_MENU in MSDN documentation!

'Obviously, we're going to be doing a lot of subclassing inside this control.
Private cSubclass As cSelfSubHookCallback

'Unlike the edit box, which may be recreated multiple times as properties change, we only need to subclass the parent window once.
' After it has been subclassed, this will be set to TRUE.
Private m_ParentHasBeenSubclassed As Boolean

'Now, for a rather lengthy explanation on these next variables, and why they're necessary.
'
'Per Wikipedia: "A dead key is a special kind of a modifier key on a typewriter or computer keyboard that is typically used to attach
' a specific diacritic to a base letter. The dead key does not generate a (complete) character by itself but modifies the character
' generated by the key struck immediately after... For example, if a keyboard has a dead key for the grave accent (`), the French
' character à can be generated by first pressing ` and then A, whereas è can be generated by first pressing ` and then E. Usually, the
' diacritic in an isolated form can be generated with the dead key followed by space, so a plain grave accent can be typed by pressing
' ` and then Space."
'
'Dead keys are a huge PITA in Windows.  As an example of how few people have cracked this problem, see:
' http://stackoverflow.com/questions/2604541/how-can-i-use-tounicode-without-breaking-dead-key-support
' http://stackoverflow.com/questions/1964614/toascii-tounicode-in-a-keyboard-hook-destroys-dead-keys
' http://stackoverflow.com/questions/3548932/keyboard-hook-changes-the-behavior-of-keys
' http://stackoverflow.com/questions/15488682/double-characters-shown-when-typing-special-characters-while-logging-keystrokes
'
'Because of this mess, we must do a fairly significant amount of custom processing to make sure dead keys are handled correctly.
'
'The problem, in a nutshell, is that Windows provides one function pair for processing a string of keypresses into a usable character:
' ToUnicode/ToUnicodeEx.  These functions have the unenviable job of taking a string of virtual keys (including all possible modifiers),
' comparing them against the thread's active keyboard layout, and returning one or more Unicode characters resulting from those keypresses.
' For some East Asian languages, a half-dozen keypresses can be chained together to form a single glyph, so this task is not a simple one,
' and it is DEFINITELY not something a sole programmer could ever hope to reverse-engineer.
'
'In an attempt to be helpful, Microsoft engineers decided that when ToUnicode successfully detects and returns a glyph, it will also
' purge the current key buffer of any keystrokes that were used to generate said glyph.  This is fine if you intend to immediately return
' the glyph without further processing, but then the engineers did something truly asinine - they also made ToUnicode the recommended way
' to check for dead keys!  ToUnicode returns -1 if it determines that the current keystroke pattern consists of only dead keys, and you can
' use that return value to detect times when you shouldn't raise a character press (because you're waiting for the rest of the dead key
' pattern to arrive).  The problem?  *As soon as you use ToUnicode to check the dead key state, the dead key press is permanently removed
' from the key buffer!*  AAARRRRGGGHHH @$%@#$^&#$
'
'After a painful amount of trial and error, I have devised the following system for working around this mess.  Whenever ToUnicode returns
' a -1 (indicating a dead keypress), PD makes a full copy of the dead key keyboard state, including scan codes and a full key map.  On a
' subsequent keypress, all that dead key information is artificially inserted into the key buffer, then ToUnicode is used to analyze
' this artificially constructed buffer.
'
'This is not pretty in any way, but the code is concise, and this is a resolution that thousands of other angry programmers
' were unable to locate - so I'm counting myself lucky and not obsessing over the inelegance of it.
'
'Anyway, these module-level variables are used to cache dead key values until they are actually needed.
Private m_DeadCharVal As Long
Private m_DeadCharScanCode As Long
Private m_DeadCharKeyStateData(0 To 255) As Byte

'Alt+number keycode entry has to be handled manually, if the keycode exceeds FFFF.
Private m_AltKeyMode As Boolean
Private assembledVirtualKeyString As String

'Like other custom PD UC's, tab behavior can be modified depending where the text box is sited
Private m_TabMode As PDUC_TAB_BEHAVIOR

'Dynamic hooking requires us to track focus events with care.  When focus is lost, we must relinquish control of the keyboard.
' This value will be set to TRUE if the API edit box currently has focus.
Private m_HasFocus As Boolean

'Because our API edit box is not tied into VB's default tab stop handling, we must jump through some hoops to forward focus correctly.
' Our hook proc will capture the Tab key that causes focus to enter the control, and mistakenly assume it is a Tab keypress from
' *within* the control.  To prevent this from happening, we enforce a slight time delay from when our hook procedure begins, to when
' we capture Tab keypresses.  This prevents faulty Tab-key handling.
Private m_TimeAtFocusEnter As Long

'Tracks whether the control (any component) has focus.  This is helpful as we must synchronize between VB's focus events and API
' focus events.  This value is deliberately kept separate from m_HasFocus, above, as we only use this value to raise our own
' Got/Lost focus events when the *entire control* loses focus (vs any one individual component).
Private m_ControlHasFocus As Boolean

'Persistent back buffer, which we manage internally
Private m_BackBuffer As pdDIB

'Flicker-free window painter
Private WithEvents cPainter As pdWindowPainter
Attribute cPainter.VB_VarHelpID = -1

'If the user resizes an edit box, the control's back buffer needs to be redrawn.  If we resize the edit box as part of an internal
' AutoSize calculation, however, we will already be in the midst of resizing the backbuffer - so we override the behavior of the
' UserControl_Resize event, using this variable.
Private m_InternalResizeState As Boolean

'The system handles drawing of the edit box.  This persistent brush handle is returned to the relevant window message,
' and WAPI uses it to draw the edit box background.
Private m_EditBoxBrush As Long

'While inside the hook event, this will be set to TRUE.  Because we raise events directly from the hook, we sometimes need to postpone
' crucial actions (like releasing the hook) until the hook proc has exited.
Private m_InHookNow As Boolean

'If the user attempts to set the Text property before the edit box is created (e.g. when the control is invisible), we will back up the
' text to this string.  When the edit box is created, this text will be automatically placed inside the control.
Private m_TextBackup As String

'Additional helpers for rendering themed and multiline tooltips
Private m_Tooltip As pdToolTip
Private m_ToolString As String

'hWnds aren't exposed by default
Public Property Get hWnd() As Long
Attribute hWnd.VB_UserMemId = -515
    hWnd = UserControl.hWnd
End Property

'Container hWnd must be exposed for external tooltip handling
Public Property Get containerHwnd() As Long
    containerHwnd = UserControl.containerHwnd
End Property

'The Enabled property is a bit unique; see http://msdn.microsoft.com/en-us/library/aa261357%28v=vs.60%29.aspx
' TODO: disable text box as well
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_UserMemId = -514
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal newValue As Boolean)
    
    If m_EditBoxHwnd <> 0 Then EnableWindow m_EditBoxHwnd, IIf(newValue, 1, 0)
    UserControl.Enabled = newValue
    
    'Redraw the window to match
    If g_IsProgramRunning Then redrawBackBuffer
    
    PropertyChanged "Enabled"
    
End Property

'Font properties; only a subset are used, as PD handles most font settings automatically
Public Property Get FontSize() As Single
    FontSize = m_FontSize
End Property

Public Property Let FontSize(ByVal newSize As Single)
    
    If newSize <> m_FontSize Then
        m_FontSize = newSize
        
        If Not (curFont Is Nothing) Then
            
            'Recreate the font object
            curFont.ReleaseFromDC
            curFont.SetFontSize m_FontSize
            curFont.CreateFontObject
            
            'Edit box sizes are ideally set by the system, at creation time, so we don't have a choice but to recreate the box now
            createEditBox
            
        End If
        
    End If
    
End Property

Public Property Get Multiline() As Boolean
    Multiline = m_Multiline
End Property

Public Property Let Multiline(ByVal newState As Boolean)
    
    If newState <> m_Multiline Then
        m_Multiline = newState
        
        'Changing the multiline property requires a full recreation of the edit box (e.g. it cannot be changed via window message alone).
        ' Also, note that the createEditBox function will automatically handle the backup/restoration of any text currently in the edit box.
        createEditBox
        
        PropertyChanged "Multiline"
        
    End If
    
End Property

'SelStart is used by some PD functions to control caret positioning after automatic text updates (as used in the text up/down)
Public Property Get SelStart() As Long
    
    Dim startPos As Long, endPos As Long, retVal As Long
    
    If m_EditBoxHwnd <> 0 Then
        retVal = SendMessage(m_EditBoxHwnd, EM_GETSEL, startPos, endPos)
        SelStart = startPos
    End If
    
End Property

Public Property Let SelStart(ByVal newPosition As Long)
    
    If m_EditBoxHwnd <> 0 Then
        SendMessage m_EditBoxHwnd, EM_SETSEL, newPosition, newPosition
    End If
    
End Property

'Tab handling for API windows is complicated.  This control (like most other PD controls) supports variable tab key behavior.
' For UC instances embedded inside other UCs, we will raise a "TabPress" event, which the UC can then handle in any way it chooses
' (for example, forwarding focus to another control in the UC, vs forwarding focus outside the UC).  For standalone instances,
' the default tab behavior can be specified.
Public Property Get TabBehavior() As PDUC_TAB_BEHAVIOR
    TabBehavior = m_TabMode
End Property

Public Property Let TabBehavior(ByVal newBehavior As PDUC_TAB_BEHAVIOR)
    m_TabMode = newBehavior
    PropertyChanged "TabBehavior"
End Property

'For performance reasons, the current control's text is not stored persistently.  It is retrieved, as needed, using an on-demand model.
Public Property Get Text() As String
Attribute Text.VB_ProcData.VB_Invoke_Property = ";Text"
Attribute Text.VB_UserMemId = 0
Attribute Text.VB_MemberFlags = "200"
    
    'Make sure the text box has been initialized
    If m_EditBoxHwnd <> 0 Then
    
        'Retrieve the length of the edit box's current string.  Note that this is not necessarily the *actual* length.  Instead, it is an
        ' interop-friendly measurement that represents the maximum possible size of the buffer, when accounting for mixed ANSI and Unicode
        ' strings (among other things).
        Dim bufLength As Long
        bufLength = GetWindowTextLength(m_EditBoxHwnd) + 1
        
        'Note that there is a disconnect here.  GetWindowTextLength returns the length, in characters, of of the string to be returned.
        ' It does NOT include a +1 for the null terminator (which is implicit in VB strings, but relevant when preparing a buffer).
        ' This is why we append a +1, above.
        
        'Prepare a string buffer at that size
        Text = Space$(bufLength)
        
        'Retrieve the window's text.  Note that the retrieval function will return that actual length of the buffer (not counting the null
        ' terminator).  On the off chance that the actual length differs from the buffer we were initially given, trim the string to match.
        Dim actualBufLength As Long
        actualBufLength = GetWindowText(m_EditBoxHwnd, StrPtr(Text), bufLength)
        
        If actualBufLength <> bufLength Then Text = Left$(Text, actualBufLength)
        
    Else
        Text = m_TextBackup
    End If
    
End Property

Public Property Let Text(ByRef newString As String)

    'Unfortunately, we cannot use SetWindowText here.  SetWindowText does not expand tab characters in a string,
    ' so our only option is to manually send a WM_SETTEXT message to the text box.
    If m_EditBoxHwnd <> 0 Then
        SendMessage m_EditBoxHwnd, WM_SETTEXT, 0&, ByVal StrPtr(newString)
    Else
        m_TextBackup = newString
    End If
    
    'We now fork our behavior according to IDE vs run-time.  PropertyChanged events are slow and unnecessary at run-time, while raising
    ' events is unnecessary in the IDE.
    If g_IsProgramRunning Then
    
        'Note that updating text this way will not raise an EN_UPDATE message for the parent.  As such, we must raise a Change event manually.
        RaiseEvent Change
        
    Else
    
        m_TextBackup = newString
        PropertyChanged "Text"
    
    End If

End Property

'External functions can call this to fully select the text box's contents
Public Sub selectAll()

    If m_EditBoxHwnd <> 0 Then
        SendMessage m_EditBoxHwnd, EM_SETSEL, ByVal 0&, ByVal -1&
    End If

End Sub

'After curFont has been created, this function can be used to return the "ideal" height of a string rendered via the current font.
Private Function getIdealStringHeight() As Long
    getIdealStringHeight = curFont.GetHeightOfString("abc123")
End Function

'The pdWindowPaint class raises this event when the control needs to be redrawn.  The passed coordinates contain the
' rect returned by GetUpdateRect (but with right/bottom measurements pre-converted to width/height).
Private Sub cPainter_PaintWindow(ByVal winLeft As Long, ByVal winTop As Long, ByVal winWidth As Long, ByVal winHeight As Long)
    
    'Flip the relevant chunk of the buffer to the screen
    BitBlt UserControl.hDC, winLeft, winTop, winWidth, winHeight, m_BackBuffer.getDIBDC, winLeft, winTop, vbSrcCopy
        
End Sub

Private Sub tmrHookRelease_Timer()

    'If a hook is active, this timer will repeatedly try to kill it.  Do not enable it until you are certain the hook needs to be released.
    ' (This is used as a failsafe if we cannot immediately release the hook when focus is lost, for example if we are currently inside an
    '  external event, as happens in the Layer toolbox, which hides the active text box inside the KeyPress event.)
    If (m_EditBoxHwnd <> 0) And (Not m_InHookNow) Then
        RemoveHookConditional
        tmrHookRelease.Enabled = False
    End If
    
End Sub

'When the control receives focus, forcibly forward focus to the API edit box
Private Sub UserControl_GotFocus()
    
    'Mark the control-wide focus state
    If Not m_ControlHasFocus Then
        m_ControlHasFocus = True
        RaiseEvent GotFocusAPI
    End If
    
    'The user control itself should never have focus.  Forward it to the API edit box.
    If m_EditBoxHwnd <> 0 Then
        SetForegroundWindow m_EditBoxHwnd
        SetFocus m_EditBoxHwnd
    End If
    
End Sub

'When the user control is hidden, we must hide the edit box window as well
Private Sub UserControl_Hide()
    If m_EditBoxHwnd <> 0 Then ShowWindow m_EditBoxHwnd, SW_HIDE
End Sub

Private Sub UserControl_Initialize()

    m_EditBoxHwnd = 0
    
    Set curFont = New pdFont
    m_FontSize = 10
    
    'Note that we are not currently responsible for any resize events
    m_InternalResizeState = False
    
    'At run-time, initialize a subclasser
    If g_IsProgramRunning Then Set cSubclass = New cSelfSubHookCallback
    
    'When not in design mode, initialize a tracker for mouse events
    If g_IsProgramRunning Then
        
        'Start a flicker-free window painter
        Set cPainter = New pdWindowPainter
        cPainter.StartPainter Me.hWnd
                
    'In design mode, initialize a base theming class, so our paint function doesn't fail
    Else
        
        Set g_Themer = New pdVisualThemes
        
    End If
    
    'Create an initial font object
    refreshFont
    
End Sub

Private Sub UserControl_InitProperties()
    Enabled = True
    FontSize = 10
    Multiline = False
    TabBehavior = TabDefaultBehavior
    Text = ""
End Sub

Private Sub UserControl_LostFocus()
    
    'Mark the control-wide focus state
    If m_ControlHasFocus Then
        m_ControlHasFocus = False
        RaiseEvent LostFocusAPI
    End If
    
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    With PropBag
        Enabled = .ReadProperty("Enabled", True)
        FontSize = .ReadProperty("FontSize", 10)
        Multiline = .ReadProperty("Multiline", False)
        TabBehavior = .ReadProperty("TabBehavior", TabDefaultBehavior)
        Text = .ReadProperty("Text", "")
    End With

End Sub

'When the user control is resized, the text box must be resized to match
Private Sub UserControl_Resize()

    'Ignore resize events generated internally (e.g. sizing a text box to the current font)
    If Not m_InternalResizeState Then
    
        'Reposition the edit text box
        If m_EditBoxHwnd <> 0 Then
            
            'Retrieve the edit box's window rect, which is generated relative to the underlying DC
            Dim tmpRect As winRect
            getEditBoxRect tmpRect
            
            With tmpRect
                MoveWindow m_EditBoxHwnd, .x1, .y1, .x2, .y2, 1
            End With
            
        End If
        
        'Redraw the control background
        UpdateControlSize
                
    End If
    
    RaiseEvent Resize

End Sub

'Show the control and the edit box.  (This is the first place the edit box is typically created, as well.)
Private Sub UserControl_Show()
    
    'Redraw the control
    'UpdateControlSize
    
    'If we have not yet created the edit box, do so now
    If m_EditBoxHwnd = 0 Then
        
        createEditBox
    
    'The edit box has already been created, so we just need to show it.  Note that we explicitly set flags to NOT activate
    ' the window, as we don't want it stealing focus.
    Else
        If m_EditBoxHwnd <> 0 Then ShowWindow m_EditBoxHwnd, SW_SHOWNA
    End If
    
    'When the control is first made visible, remove the control's tooltip property and reassign it to the checkbox
    ' using a custom solution (which allows for linebreaks and theming).  Note that this has the ugly side-effect of
    ' permanently erasing the extender's tooltip, so FOR THIS CONTROL, TOOLTIPS MUST BE SET AT RUN-TIME!
    '
    'TODO!  Add helper functions for setting the tooltip to the created hWnd, instead of the VB control
    m_ToolString = Extender.ToolTipText
    
    If Len(m_ToolString) <> 0 Then
    
        If (m_Tooltip Is Nothing) Then Set m_Tooltip = New pdToolTip
        m_Tooltip.setTooltip m_EditBoxHwnd, Me.hWnd, m_ToolString
        Extender.ToolTipText = ""
        
    End If

End Sub

Private Sub getEditBoxRect(ByRef targetRect As winRect)

    With targetRect
        .x1 = 2
        .y1 = 2
        .x2 = UserControl.ScaleWidth - 4
        .y2 = UserControl.ScaleHeight - 4
        '.x2 = apiWidth(Me.hWnd) - 4
        '.y2 = apiWidth(Me.hWnd) - 4
    End With

End Sub

'Create a brush for drawing the box background
Private Sub createEditBoxBrush()

    If m_EditBoxBrush <> 0 Then DeleteObject m_EditBoxBrush
    
    If g_IsProgramRunning Then
        m_EditBoxBrush = CreateSolidBrush(g_Themer.GetThemeColor(PDTC_BACKGROUND_DEFAULT))
    Else
        m_EditBoxBrush = CreateSolidBrush(RGB(0, 255, 0))
    End If

End Sub

'As the wrapped system edit box may need to be recreated when certain properties are changed, this function is used to
' automate the process of destroying an existing window (if any) and recreating it anew.
Private Function createEditBox() As Boolean

    'If the edit box already exists, copy its text, then kill it
    Dim curText As String
    If m_EditBoxHwnd <> 0 Then
        curText = Text
    Else
        curText = m_TextBackup
    End If
    destroyEditBox
    
    'Create a brush for drawing the box background
    createEditBoxBrush
    
    'Figure out which flags to use, based on the control's properties
    Dim flagsWinStyle As Long, flagsWinStyleExtended As Long, flagsEditControl As Long
    flagsWinStyle = WS_VISIBLE Or WS_CHILD
    flagsWinStyleExtended = 0
    flagsEditControl = 0
    
    If m_Multiline Then
        flagsWinStyle = flagsWinStyle Or WS_VSCROLL
        flagsEditControl = flagsEditControl Or ES_MULTILINE Or ES_WANTRETURN Or ES_AUTOVSCROLL 'Or ES_NOHIDESEL
    Else
        flagsEditControl = flagsEditControl Or ES_AUTOHSCROLL 'Or ES_NOHIDESEL
    End If
    
    'Multiline text boxes can have any height.  Single-line text boxes cannot; they are forced to an ideal height,
    ' using the current font as our guide.  We check for this here, prior to creating the edit box, as we can't easily
    ' access our font object once we assign it to the edit box.
    If Not m_Multiline Then
        If Not (curFont Is Nothing) Then
            
            'Determine a standard string height
            Dim idealHeight As Long
            idealHeight = getIdealStringHeight()
            
            'Resize the user control accordingly; the formula for height is the string height + 5px of borders.
            ' (5px = 2px on top, 3px on bottom.)  User control width is not changed.
            m_InternalResizeState = True
            
            'If the program is running (e.g. NOT design-time) resize the user control to match.  This improves compile-time performance, as there
            ' are a lot of instances in this control, and their size events will be fired during compilation.
            If g_IsProgramRunning Then
                UserControl.Height = PXToTwipsY(idealHeight + 5)
            End If
            
            m_InternalResizeState = False
            
        End If
    End If
    
    'Retrieve the edit box's window rect, which is generated relative to the underlying DC
    Dim tmpRect As winRect
    getEditBoxRect tmpRect
    
    With tmpRect
        m_EditBoxHwnd = CreateWindowEx(flagsWinStyleExtended, ByVal StrPtr("EDIT"), ByVal StrPtr(""), flagsWinStyle Or flagsEditControl, _
        .x1, .y1, .x2, .y2, UserControl.hWnd, 0, App.hInstance, ByVal 0&)
    End With
    
    'Enable the window per the current UserControl's extender setting
    EnableWindow m_EditBoxHwnd, IIf(Me.Enabled, 1, 0)
    
    'Assign a subclasser to enable IME support
    If g_IsProgramRunning Then
        If Not (cSubclass Is Nothing) Then
            
            'Subclass the edit box
            cSubclass.ssc_Subclass m_EditBoxHwnd, 0, 1, Me, True, True, True
            cSubclass.ssc_AddMsg m_EditBoxHwnd, MSG_BEFORE, WM_KEYDOWN, WM_SETFOCUS, WM_KILLFOCUS, WM_CHAR, WM_UNICHAR, WM_MOUSEACTIVATE
            
            'Subclass the user control as well.  This is necessary for handling update messages from the edit box
            If Not m_ParentHasBeenSubclassed Then
                cSubclass.ssc_Subclass UserControl.hWnd, 0, 1, Me, True, True, False
                cSubclass.ssc_AddMsg UserControl.hWnd, MSG_BEFORE, WM_CTLCOLOREDIT, WM_COMMAND
                m_ParentHasBeenSubclassed = True
            End If
            
        End If
    End If
    
    'Assign the default font to the edit box
    refreshFont True
    
    'If the edit box had text before we killed it, restore that text now
    If Len(curText) <> 0 Then Text = curText
    
    'Return TRUE if successful
    createEditBox = (m_EditBoxHwnd <> 0)

End Function

'If an edit box currently exists, this function will destroy it.
Private Function destroyEditBox() As Boolean

    If m_EditBoxHwnd <> 0 Then
        
        If Not cSubclass Is Nothing Then
            cSubclass.ssc_UnSubclass m_EditBoxHwnd
            cSubclass.shk_TerminateHooks
        End If
        
        DestroyWindow m_EditBoxHwnd
        
    End If
    
    destroyEditBox = True

End Function

Private Sub UserControl_Terminate()
    
    'Release the edit box background brush
    If m_EditBoxBrush <> 0 Then DeleteObject m_EditBoxBrush
    
    'Destroy the edit box, as necessary
    destroyEditBox
    
    'Release any extra subclasser(s)
    If Not cSubclass Is Nothing Then cSubclass.ssc_Terminate
    
End Sub

'When the font used for the edit box changes in some way, it can be recreated (refreshed) using this function.  Note that font
' creation is expensive, so it's worthwhile to avoid this step as much as possible.
Private Sub refreshFont(Optional ByVal forceRefresh As Boolean = False)
    
    Dim fontRefreshRequired As Boolean
    fontRefreshRequired = curFont.HasFontBeenCreated
    
    'Update each font parameter in turn.  If one (or more) requires a new font object, the font will be recreated as the final step.
    
    'Font face is always set automatically, to match the current program-wide font
    If (Len(g_InterfaceFont) <> 0) And (StrComp(curFont.GetFontFace, g_InterfaceFont, vbBinaryCompare) <> 0) Then
        fontRefreshRequired = True
        curFont.SetFontFace g_InterfaceFont
    End If
    
    'In the future, I may switch to GDI+ for font rendering, as it supports floating-point font sizes.  In the meantime, we check
    ' parity using an Int() conversion, as GDI only supports integer font sizes.
    If Int(m_FontSize) <> Int(curFont.GetFontSize) Then
        fontRefreshRequired = True
        curFont.SetFontSize m_FontSize
    End If
        
    'Request a new font, if one or more settings have changed
    If fontRefreshRequired Or forceRefresh Then
        
        curFont.CreateFontObject
        
        'Whenever the font is recreated, we need to reassign it to the text box.  This is done via the WM_SETFONT message.
        If m_EditBoxHwnd <> 0 Then SendMessage m_EditBoxHwnd, WM_SETFONT, curFont.GetFontHandle, IIf(UserControl.Extender.Visible, 1, 0)
            
        'Also, the back buffer needs to be rebuilt to reflect the new font metrics
        UpdateControlSize
            
    End If
    
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    'Store all associated properties
    With PropBag
        .WriteProperty "Enabled", Me.Enabled, True
        .WriteProperty "FontSize", m_FontSize, 10
        .WriteProperty "Multiline", m_Multiline, False
        .WriteProperty "TabBehavior", m_TabMode, TabDefaultBehavior
        .WriteProperty "Text", m_TextBackup, ""
    End With
    
End Sub

'External functions can call this to request a redraw.  This is helpful for live-updating theme settings, as in the Preferences dialog.
Public Sub UpdateAgainstCurrentTheme()
    
    If g_IsProgramRunning Then
        
        'Create a brush for drawing the box background
        createEditBoxBrush
        
        'Update the current font, as necessary
        refreshFont
        
        'Force an immediate repaint
        UpdateControlSize
                
    End If
    
End Sub

'When the control is resized, several things need to happen:
' 1) We need to forward the resize request to the API edit window
' 2) We need to resize the button's back buffer, then redraw it
Private Sub UpdateControlSize()

    'Reset our back buffer, and reassign the font to it
    If m_BackBuffer Is Nothing Then Set m_BackBuffer = New pdDIB
    m_BackBuffer.createBlank UserControl.ScaleWidth, UserControl.ScaleHeight, 24
        
    'Redraw the back buffer
    redrawBackBuffer

End Sub

'After the back buffer has been correctly sized and positioned, this function handles the actual painting.  Similarly, for state changes
' that don't require a resize (e.g. gain/lose focus), this function should be used.
Private Sub redrawBackBuffer()
    
    'Start by erasing the back buffer
    If g_IsProgramRunning Then
    
        'Fill color changes depending on enablement
        Dim editBoxBackgroundColor As Long
        
        If Me.Enabled Then
            editBoxBackgroundColor = g_Themer.GetThemeColor(PDTC_BACKGROUND_DEFAULT)
        Else
            editBoxBackgroundColor = g_Themer.GetThemeColor(PDTC_GRAY_HIGHLIGHT)
        End If
        
        GDI_Plus.GDIPlusFillDIBRect m_BackBuffer, 0, 0, m_BackBuffer.getDIBWidth, m_BackBuffer.getDIBHeight, editBoxBackgroundColor, 255
        
        'If the control is disabled, the BackColor property actually becomes relevant (because the edit box will allow the back color
        ' to "show through").  As such, set it now, and note that we can use VB's internal property, because it simply wraps the
        ' matching GDI function(s).
        UserControl.BackColor = g_Themer.GetThemeColor(PDTC_GRAY_HIGHLIGHT)
        
    Else
        m_BackBuffer.createBlank m_BackBuffer.getDIBWidth, m_BackBuffer.getDIBHeight, 24, RGB(255, 255, 255)
    End If
    
    'The edit box has a 1px border, whose color changes depending on focus
    Dim editBoxBorderColor As Long
    
    If m_HasFocus Then
        editBoxBorderColor = g_Themer.GetThemeColor(PDTC_ACCENT_DEFAULT)
    Else
        editBoxBorderColor = g_Themer.GetThemeColor(PDTC_GRAY_DEFAULT)
    End If
    
    'Draw the border
    GDI_Plus.GDIPlusDrawRectOutlineToDC m_BackBuffer.getDIBDC, 0, 0, m_BackBuffer.getDIBWidth - 1, m_BackBuffer.getDIBHeight - 1, editBoxBorderColor
    
    'Paint the buffer to the screen
    If g_IsProgramRunning Then cPainter.RequestRepaint Else BitBlt UserControl.hDC, 0, 0, m_BackBuffer.getDIBWidth, m_BackBuffer.getDIBHeight, m_BackBuffer.getDIBDC, 0, 0, vbSrcCopy

End Sub

'If a virtual key code is numeric, return TRUE.
Private Function isVirtualKeyNumeric(ByVal vKey As Long, Optional ByRef numericValue As Long = 0) As Boolean
    
    If (vKey >= VK_0) And (vKey <= VK_9) Then
        isVirtualKeyNumeric = True
        numericValue = vKey - VK_0
    Else
    
        If (vKey >= VK_NUMPAD0) And (vKey <= VK_NUMPAD9) Then
            isVirtualKeyNumeric = True
            numericValue = vKey - VK_NUMPAD0
        Else
            isVirtualKeyNumeric = False
        End If
    End If
    
End Function

'Given a virtual keycode, return TRUE if the keycode is a command key that must be manually forwarded to an edit box.  Command keys include
' arrow keys at present, but in the future, additional keys can be added to this list.
'
'NOTE FOR OUTSIDE USERS!  These key constants are declared publicly in PD, because they are used many places.  You can find virtual keycode
' declarations in PD's Public_Constants module, or at this MSDN link:
' http://msdn.microsoft.com/en-us/library/windows/desktop/dd375731%28v=vs.85%29.aspx
Private Function doesVirtualKeyRequireSpecialHandling(ByVal vKey As Long) As Boolean
    
    Select Case vKey
    
        Case VK_LEFT, VK_UP, VK_RIGHT, VK_DOWN
            doesVirtualKeyRequireSpecialHandling = True
            
        Case VK_TAB
            doesVirtualKeyRequireSpecialHandling = True
        
        Case Else
            doesVirtualKeyRequireSpecialHandling = False
    
    End Select
        
End Function

'Note that the vKey constant below is a virtual key mapping, not necessarily a standard VB key constant
Private Function IsVirtualKeyDown(ByVal vKey As Long) As Boolean
    IsVirtualKeyDown = GetAsyncKeyState(vKey) And &H8000
End Function

'Install a keyboard hook in our window
Private Sub InstallHookConditional()

    'Check for an existing hook
    If Not m_HasFocus Then
    
        'Note the time.  This is used to prevent keypresses occurring immediately prior to the hook, from being
        ' caught within our hook proc!
        m_TimeAtFocusEnter = GetTickCount
        
        'Note that this window is now active
        m_HasFocus = True
        
        'No hook exists.  Hook the control now.
        cSubclass.shk_SetHook WH_KEYBOARD, False, MSG_BEFORE, m_EditBoxHwnd, 2, Me, , True
        
        'Redraw the control to reflect focus state
        redrawBackBuffer
    
    End If

End Sub

Private Sub RemoveHookConditional()

    'Check for an existing hook
    If m_HasFocus Then
        
        'Note that this window is now considered inactive
        m_HasFocus = False
        
        'A hook exists.  Uninstall it now.
        cSubclass.shk_UnHook WH_KEYBOARD
        
        'Redraw the control to reflect focus state
        redrawBackBuffer
        
    End If
    
End Sub

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
    
    m_InHookNow = True
    bHandled = False
        
    If (lHookType = WH_KEYBOARD) And m_HasFocus Then
        
        'MSDN states that negative codes must be passed to the next hook, without processing
        ' (see http://msdn.microsoft.com/en-us/library/ms644984.aspx)
        If nCode >= 0 Then
            
            'Before proceeding, cache the status of all 256 keyboard keys.  This is important for non-Latin keyboards, which can
            ' produce Unicode characters in a variety of ways.  (For example, by holding down multiple keys at once.)  If we end
            ' up forwarding a key event to the default WM_KEYDOWN handler, it will need this information in order to parse any
            ' IME input.
            GetKeyboardState m_keyStateData(0)
            
            'Bit 31 of lParam is 0 if a key is being pressed, and 1 if it is being released.  We use this to raise
            ' separate KeyDown and KeyUp events, as necessary.
            If lParam < 0 Then
                
                'Raise a KeyUp event.  The caller can deny further handling by setting the appropriate param to TRUE.
                bHandled = False
                RaiseEvent KeyPress(wParam, bHandled)
                
                'If the user didn't forcibly override further key handling, proceed with default proc behavior
                If Not bHandled Then
                
                    'The default key handler works just fine for character keys.  However, dialog keys (e.g. arrow keys) get eaten
                    ' by VB, so we must manually catch them in this hook, and forward them direct to the edit control.
                    If doesVirtualKeyRequireSpecialHandling(wParam) Then
                        
                        'Key up events will be raised twice; once in a transitionary stage, and once again in a final stage.
                        ' To prevent double-raising of KeyUp events, we check the transitionary state before proceeding
                        If ((lParam And 1) <> 0) And ((lParam And 3) = 1) Then
                            
                            'Non-tab keys that require special handling are text-dependent keys (e.g. arrow keys).  Simply forward these
                            ' directly to the edit box, and it will take care of the rest.
                            
                            'WM_KEYUP requires that certain lParam bits be set.  See http://msdn.microsoft.com/en-us/library/windows/desktop/ms646281%28v=vs.85%29.aspx
                            SendMessage m_EditBoxHwnd, WM_KEYUP, wParam, ByVal (lParam And &HDFFFFF81 Or &HC0000000)
                            bHandled = True
                            
                        End If
                    
                    End If
                
                    'Another special case we must handle here in the hook is Alt+ key presses
                    'If we're not tracking sys key messages, start now
                    If m_AltKeyMode And ((lParam And 1) <> 0) And ((lParam And 3) = 1) And (lParam < 0) Then
                        
                        'See if the Alt key is being released.  If it is, submit the retrieved character code.
                        If wParam = VK_ALT Then
                            
                            m_AltKeyMode = False
                            
                            'Make sure the typed code is not longer than 10 chars (the limit for a Long)
                            If Len(assembledVirtualKeyString) <> 0 And Len(assembledVirtualKeyString) <= 10 Then
                            
                                'If the Alt+keycode is larger than an Int, we must submit it manually.
                                Dim charAsLong As Long
                            
                                'For better range checking, first copy the value into a Currency-type integer
                                Dim testRange As Currency
                                testRange = CCur(assembledVirtualKeyString)
                                
                                'Perform a second check, to make sure the value fits within a VB long.
                                If (testRange <= LONG_MAX) And (testRange > 65536) Then
                                    
                                    charAsLong = CLng(assembledVirtualKeyString)
                                    If charAsLong And &HFFFF0000 <> 0 Then
                                        
                                        'Convert it into two chars.  The code for this is rather involved; see http://en.wikipedia.org/wiki/UTF-16#Code_points_U.2B010000_to_U.2B10FFFF
                                        ' for details.
                                        charAsLong = charAsLong - &H10000
                                        
                                        Dim charHiWord As Long, charLoWord As Long
                                        charHiWord = ((charAsLong \ 1024) And &H7FF) + &HD800&
                                        charLoWord = (charAsLong And &H3FF) + &HDC00&
                                        
                                        'Send those chars to the edit box
                                        Dim tmpMsg As winMsg
                                        tmpMsg.hWnd = m_EditBoxHwnd
                                        tmpMsg.sysMsg = WM_CHAR
                                        tmpMsg.wParam = charHiWord
                                        tmpMsg.lParam = lParam
                                        tmpMsg.msgTime = GetTickCount()
                                        DispatchMessage tmpMsg
                                        
                                        tmpMsg.wParam = charLoWord
                                        tmpMsg.msgTime = GetTickCount()
                                        DispatchMessage tmpMsg
                                    
                                        assembledVirtualKeyString = ""
                                        bHandled = True
                                    
                                    End If
                                End If
                            End If
                            
                        'If we're already tracking sys key messages, continue assembling numeric keypresses
                        Else
                        
                            'Make sure the keypress is numeric.  If it is, continue assembling a virtual string.
                            Dim numCheck As Long
                            If isVirtualKeyNumeric(wParam, numCheck) Then
                                assembledVirtualKeyString = assembledVirtualKeyString & CStr(numCheck)
                            Else
                                If Len(assembledVirtualKeyString) <> 0 Then assembledVirtualKeyString = ""
                            End If
                            
                            bHandled = True
                            
                        End If
                
                    End If
                    
                End If
                
            Else
            
                'The default key handler works just fine for character keys.  However, dialog keys (e.g. arrow keys) get eaten
                ' by VB, so we must manually catch them in this hook, and forward them direct to the edit control.
                If doesVirtualKeyRequireSpecialHandling(wParam) Then
                    
                    'On a single-line control, the tab key should be used to redirect focus to a new window.
                    If (wParam = VK_TAB) Then
                        
                        'Multiline edit boxes accept tab keypresses.  Single line ones do not, so interpret TAB as a
                        ' request to forward (or reverse) focus.
                        If (Not m_Multiline) And ((GetTickCount - m_TimeAtFocusEnter) > 250) Then
                            
                            'We can handle this tab press one of two ways, based on the TabBehavior property
                            If m_TabMode = TabDefaultBehavior Then
                                
                                'Immediately forward focus to the next control
                                UserControl_Support.ForwardFocusToNewControl Me, Not IsVirtualKeyDown(VK_SHIFT)
                                
                            'Let the caller determine how to handle the keypress
                            Else
                                RaiseEvent TabPress(Not IsVirtualKeyDown(VK_SHIFT))
                            End If
                            
                            'Eat the keypress
                            bHandled = True
                            
                        End If
                        
                    Else
                    
                        'WM_KEYDOWN requires that certain bits be set.  See http://msdn.microsoft.com/en-us/library/windows/desktop/ms646280%28v=vs.85%29.aspx
                        SendMessage m_EditBoxHwnd, WM_KEYDOWN, wParam, ByVal (lParam And &H51111111)
                        bHandled = True
                        
                    End If
                    
                End If
                
                'Another special case we must handle here in the hook is Alt+ keypresses.  These work fine for values below 255.
                ' They are not operational for character values above that range.
                If (Not m_AltKeyMode) And IsVirtualKeyDown(VK_ALT) Then
                    m_AltKeyMode = True
                    assembledVirtualKeyString = ""
                End If
                
            End If
            
        End If
            
    End If
    
    'Per MSDN, return the value of CallNextHookEx, contingent on whether or not we handled the keypress internally.
    ' Note that if we do not manually handle a keypress, this behavior allows the default keyhandler to deal with
    ' the pressed keys (and raise its own WM_CHAR events, etc).
    If (Not bHandled) Then
        lReturn = CallNextHookEx(0, nCode, wParam, lParam)
    Else
        lReturn = 1
    End If
    
    m_InHookNow = False
    
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
    
    'Now comes a really messy bit of VB-specific garbage.
    '
    'Normally, a Unicode window (e.g. one created with CreateWindowW/ExW) automatically receives Unicode window messages.
    ' Keycodes are an exception to this, because they use a unique message chain that also involves the parent window
    ' of the Unicode window - and in the case of VB, that parent window's message pump is *not* Unicode-aware, so it
    ' doesn't matter that our window *is*!
    '
    'Why involve the parent window in the keyboard processing chain?  In a nutshell, an initial WM_KEYDOWN message contains only
    ' a virtual key code (which can be thought of as representing the ID of a physical keyboard key, not necessarily a specific
    ' character or letter).  This virtual key code must be translated into either a character or an accelerator, based on the use
    ' of Alt/Ctrl/Shift/Tab etc keys, any IME mapping, other accessibility features, and more.  In the case of accelerators, the
    ' parent window must be involved in the translation step, because it is most likely the window with an accelerator table (as
    ' would be used for menu shortcuts, among other things).  So a child window can't avoid its parent window being involved in
    ' key event handling.
    '
    'Anyway, the moral of the story is that we have to do a shitload of extra work to bypass the default message translator.
    ' Without this, IME entry methods (easily tested via the Windows on-screen keyboard and some non-Latin language) result in
    ' ???? chars, despite use of a Unicode window - and that ultimately defeats the whole point of a Unicode text box, no?
    Select Case uMsg
        
        'The parent receives this message for all kinds of things; we subclass it to track when the edit box's contents have changed.
        ' (And when we don't handle the message, it is *very important* that we forward it correctly!
        Case WM_COMMAND
            
            'Make sure the command is relative to *our* edit box, and not another one
            If lParam = m_EditBoxHwnd Then
            
                'Check for the EN_UPDATE flag; if present, raise the CHANGE event
                If (wParam \ &H10000) = EN_UPDATE Then
                    RaiseEvent Change
                    bHandled = True
                End If
                
            End If
        
        'The parent receives this message, prior to the edit box being drawn.  The parent can use this to assign text and
        ' background colors to the edit box.
        Case WM_CTLCOLOREDIT
            
            'Make sure the command is relative to *our* edit box, and not another one
            If lParam = m_EditBoxHwnd Then
            
                'We can set the text color directly, using the API
                If g_IsProgramRunning Then
                    SetTextColor wParam, g_Themer.GetThemeColor(PDTC_TEXT_EDITBOX)
                Else
                    SetTextColor wParam, RGB(0, 0, 128)
                End If
                
                'We return the background brush
                bHandled = True
                lReturn = m_EditBoxBrush
                
            End If
        
        Case WM_CHAR
            'On a single-line text box, pressing Enter will cause a ding.  (It's insane that this is the default behavior.)
            ' To prevent this, we capture return presses, and forcibly terminate them.  If the user wants to do something with
            ' Enter keypresses on these controls, they will have already handled the key event in the hook proc, above.
            If (Not m_Multiline) And ((wParam = VK_RETURN) Or (wParam = VK_TAB)) Then
                bHandled = True
                lReturn = 0
            End If
            
            'Debug.Print "WM_CHAR: " & wParam & "," & lParam
        
        'WM_UNICHAR messages are never sent by Windows.  However, third-party IMEs may send them.  Before allowing WM_UNICHAR
        ' messages to pass, Windows will first probe a window by sending the UNICODE_NOCHAR value.  If a window responds with 1
        ' (instead of 0, as DefWindowProc does), Windows will allow WM_UNICHAR messages to pass.
        Case WM_UNICHAR
            If wParam = UNICODE_NOCHAR Then
                bHandled = True
                lReturn = 1
            End If
            'Debug.Print "UNICODE char received: " & wParam & "," & lParam
        
        'Manually dispatch WM_KEYDOWN messages.
        Case WM_KEYDOWN
            
            'Because we will be dispatching our own WM_CHAR messages with any processed Unicode characters, we must manually
            ' assemble a full window message.  All messages will be sent to the API edit box we've created, so we can mark the
            ' message's hWnd and message type just once, at the start of the dispatch loop.
            Dim tmpMsg As winMsg
            tmpMsg.hWnd = m_EditBoxHwnd
            tmpMsg.sysMsg = WM_CHAR
            
            'Normally, we would next retrieve the status of all 256 keyboard keys.  However, our hook proc, above, has already
            ' done this for us.  The results are cached inside m_keyStateData().
            
            'Next, we need to prepare a string buffer to receive the Unicode translation of the current virtual key.
            ' This is tricky because ToUnicode/Ex do not specify a max buffer size they may write.  Michael Kaplan's
            ' definitive article series on this topic (dead link on MSDN; I found it here: http://www.siao2.com/2006/03/23/558674.aspx)
            ' uses a 10-char buffer.  That should be sufficient for our purposes as well.
            Dim tmpString As String
            tmpString = String$(10, vbNullChar)
            
            Dim unicodeResult As Long, tmpLong As Long
            
            'Before proceeding, see if a dead key was pressed.  Per MapVirtualKey's documentation (http://msdn.microsoft.com/en-us/library/windows/desktop/ms646306%28v=vs.85%29.aspx)
            ' "Dead keys (diacritics) are indicated by setting the top bit of the return value," we we can easily check in VB
            ' as it will seem to be a negative number.
            '
            'Because it is not possible to detect a dead char via ToUnicode without also removing that char from the key buffer, this
            ' is the only safe way (I know of) to detect and preprocess dead chars.  Note also that we SKIP THIS STEP if we already
            ' have a dead char in the buffer.  This allows repeat presses of a dead key (e.g. `` on a U.S. International keyboard) to
            ' pass through as double characters, which is the expected behavior.
            If (MapVirtualKey(wParam, MAPVK_VK_TO_CHAR) < 0) And (m_DeadCharVal = 0) Then
                
                'A dead key was pressed.  Rather than send this character to the edit box immediately (which defeats the whole
                ' purpose of a dead key!), we will make a note of it, and reinsert it to the key queue on a subsequent WM_KEYDOWN.
                
                'Update our stored dead key values
                m_DeadCharVal = wParam
                m_DeadCharScanCode = MapVirtualKey(wParam, MAPVK_VK_TO_VSC)
                GetKeyboardState m_DeadCharKeyStateData(0)
                
                'Setting unicodeResult to 0 prevents further processing by this proc
                unicodeResult = 0
            
            'The current key is NOT a dead char, or it is but we already have a dead char in the buffer, so we have no choice
            ' but to process it immediately.
            Else
                
                'If a dead char was initiated previously, and the current key press is NOT another dead char, insert the original
                ' dead char back into the key state buffer.  Note that we don't care about the return value, as we know it will
                ' be -1 since a dead key is being added!
                If (m_DeadCharVal <> 0) Then
                    ToUnicode m_DeadCharVal, m_DeadCharScanCode, m_DeadCharKeyStateData(0), StrPtr(tmpString), Len(tmpString), 0
                End If
                
                'Perform a Unicode translation using the pressed virtual key (wParam), a buffer of all previous relevant characters
                ' (e.g. dead chars from previous steps), and a full buffer of all current key states.
                unicodeResult = ToUnicode(wParam, MapVirtualKey(wParam, MAPVK_VK_TO_VSC), m_keyStateData(0), StrPtr(tmpString), Len(tmpString), 0)
                
                'Reset any dead character tracking
                m_DeadCharVal = 0
                
            End If
            
            'ToUnicode has four possible return values:
            ' -1: the char is an accent or diacritic.  If possible, it has been translated to a standalone spacing version
            '     (always a UTF-16 entry point), and placed in the output buffer.  Generally speaking, we don't want to treat
            '     this as an actual character until we receive the *next* character input.  This allows us to properly assemble
            '     mixed characters (for example `a should map to à, while `c maps to just `c - but we can't know how to use a
            '     dead key until the next keypress is received).  This behavior is affected by the current keyboard layout.
            '
            ' 0: function failed.  This is not a bad thing; many East Asian IMEs will merge multiple keypresses into a single
            '    character, so the preceding keypresses will not return anything.
            '
            ' 1: success.  A single Unicode character was written to the buffer.
            
            ' 2+: also success.  Multiple Unicode characters were written to the buffer, typically when a matching ligature was
            '    not found for a relevant multi-glyph input.  This is a valid return, and all characters in the buffer should
            '    be sent to the text box.  IMPORTANT NOTE: the string buffer can contain more values than the return value specifies,
            '    so it's important to handle the buffer using *this return value*, and *not the buffer's actual contents*.
            
            'We will now proceed to parse the results of ToUnicode, using its return value as our guide.
            Select Case unicodeResult
            
                'Dead character, meaning an accent or other diacritic.  Because we deal with the "dead key" case explicitly
                ' in previous steps, we don't have to deal with it here.
                Case -1
                    
                    'Note that the message was handled successfully.  (This may not be necessary, but I'm including it
                    ' "just in case", to work around potential oddities in dead key handling.)
                    bHandled = True
                    lReturn = 0
                    
                'Failure; no Unicode result.  This can happen if an IME is still assembling characters, and no action
                ' is required on our part.
                Case 0
                    
                    
                '1 or more chars were successfully processed and returned.
                Case 1 To Len(tmpString)
                    
                    'Send each processed Unicode character in turn, using DispatchMessage to completely bypass VB's
                    ' default handler.  This prevents forcible down-conversion to ANSI.
                    Dim i As Long
                    For i = 1 To unicodeResult
                        
                        'Each unprocessed character will have left a pending WM_KEYDOWN message in this window's queue.
                        ' To prevent other handlers from getting that original message (which is no longer valid), we are
                        ' going to iterate through each message in turn, replacing them with our own custom dispatches.
                        PeekMessage tmpMsg, m_EditBoxHwnd, WM_KEYFIRST, WM_KEYLAST, PM_REMOVE
                        
                        'Retrieve the unsigned Int value of this string char
                        CopyMemory tmpLong, ByVal StrPtr(tmpString) + ((i - 1) * 2), 2
                        
                        'Assemble a new window message, passing the retrieved string char as the wParam
                        tmpMsg.wParam = tmpLong
                        tmpMsg.lParam = lParam
                        tmpMsg.msgTime = GetTickCount()
                        
                        'Dispatch the message directly, bypassing TranslateMessage entirely.  NOTE!  This prevents accelerators
                        ' from working while the text box has focus, but I do not currently know a better way around this.  A custom
                        ' accelerator solution for PD would work fine, but we would need to do our own mapping from inside this proc.
                        DispatchMessage tmpMsg
                        
                    Next i
                                        
                    'Note that the message was handled successfully
                    bHandled = True
                    lReturn = 0
                
                'This case should never fire
                Case Else
                    Debug.Print "Excessively large Unicode buffer value returned: " & unicodeResult
                
            End Select
        
        'On mouse activation, the previous VB window/control that had focus will not be redrawn to reflect its lost focus state.
        ' (Presumably, this is because VB handles focus internally, rather than using standard window messages.)  To avoid the
        ' appearance of two controls simultaneously having focus, we re-set focus to the underlying user control, which forces
        ' VB to redraw the lost focus state of whatever control previously had focus.
        Case WM_MOUSEACTIVATE
            If Not m_HasFocus Then UserControl.SetFocus
            
        'When the control receives focus, initialize a keyboard hook.  This prevents accelerators from working, but it is the
        ' only way to bypass VB's internal message translator, which will forcibly convert certain Unicode chars to ANSI.
        Case WM_SETFOCUS
            
            'Forcibly disable PD's main accelerator control
            FormMain.pdHotkeys.DeactivateHook
            
            'Mark the control-wide focus state
            If Not m_ControlHasFocus Then
                m_ControlHasFocus = True
                RaiseEvent GotFocusAPI
            End If
            
            'Start hooking keypresses so we can grab Unicode chars before VB eats 'em
            InstallHookConditional
                        
        Case WM_KILLFOCUS
        
            'Re-enable PD's main accelerator control
            FormMain.pdHotkeys.ActivateHook
            
            'Mark the control-wide focus state
            If m_ControlHasFocus Then
                m_ControlHasFocus = False
                RaiseEvent LostFocusAPI
            End If
            
            'Release our hook.  In some circumstances, we can't do this immediately, so we set a timer that will release the hook
            ' as soon as the system allows.
            If m_InHookNow Then
                tmrHookRelease.Enabled = True
            Else
                RemoveHookConditional
            End If
            
        'Other messages??
        Case Else
    
    End Select



' *************************************************************
' C A U T I O N   C A U T I O N   C A U T I O N   C A U T I O N
' -------------------------------------------------------------
' DO NOT ADD ANY OTHER CODE BELOW THE "END SUB" STATEMENT BELOW
'   add this warning banner to the last routine in your class
' *************************************************************
End Sub



