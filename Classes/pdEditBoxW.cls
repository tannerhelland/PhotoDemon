VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdEditBoxW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Unicode Text Box control
'Copyright 2014-2016 by Tanner Helland
'Created: 03/November/14
'Last updated: 07/February/14
'Last update: migrate the API edit box code into a standalone class, rather than tying it explicitly to a .ctl instance
'
'This class contains code that was formerly implemented directly inside the pdTextBox user control.  To simplify
' the embedding of this window into other user controls (like pdSpinner), I've separated the text box events into
' their own class, and tried to self-contain as many bits as possible.
'
'For interaction details, look inside pdTextBox.
'
'A few additional notes on this edit box implementation:
'
' 1) I deliberately do not expose all text box properties (just the ones most relevant to PD).
' 2) Focus is currently problematic.  A fix is in-progress.  TODO!
' 3) To allow use of arrow keys and other control keys, this control must hook the keyboard.  (If it does not,
'    any parent VB window will eat control keypresses, because it doesn't know about windows created via the API!)
'    A byproduct of this is that accelerators flat-out WILL NOT WORK while this control has focus.  I haven't yet
'    settled on a good way to handle this; what I may end up doing is manually forwarding any key combinations that
'    use Alt to the default window handler, but I'm not sure this will help.  TODO!
' 4) Dynamic hooking can occasionally cause trouble in the IDE, particularly when used with break points.  It should
'    be perfectly stable once compiled.
' 5) This class relies on a number of functions in the UserControl_Support module, such as sharing GDI brushes and
'    fonts between all pdEditBoxW instances (to reduce GDI object count).
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'By design, PD textboxes raises fewer events than standard text boxes
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

'Rather than use an StdFont container (which requires VB to create redundant font objects), we track font properties manually,
' via dedicated properties.
Private m_FontSize As Single

'The following property changes require creation/destruction of the text box.  PD will automatically backup the edit box's text
' prior to recreating it, but note that text cannot be non-destructively saved when toggling the multiline property if linefeed
' characters are in use!
Private m_Multiline As Boolean

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
Private Declare Function SetBkMode Lib "gdi32" (ByVal targetDC As Long, ByVal nBkMode As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal targetDC As Long, ByVal nBkColor As Long) As Long
Private Const SBKM_OPAQUE = 2
Private Const SBKM_TRANSPARENT = 1

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
Private Declare Function GetWindowRect Lib "user32" (ByVal hndWindow As Long, ByRef lpRect As winRect) As Long

'Getting/setting window data
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextW" (ByVal hWnd As Long, ByVal lpStringPointer As Long, ByVal nMaxCount As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthW" (ByVal hWnd As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function InvalidateRect Lib "user32" (ByVal targetHwnd As Long, ByRef lpRect As winRect, ByVal bErase As Long) As Long
Private Declare Function UpdateWindow Lib "user32" (ByVal targetHwnd As Long) As Long

'Handle to the system edit box wrapped by this control
Private m_EditBoxHwnd As Long

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
Private Const WM_CTLCOLORSTATIC As Long = &H138

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
Private m_ParentHasBeenSubclassed As Boolean, m_ParentHWnd As Long

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

'Dynamic hooking requires us to track focus events with care.  When focus is lost, we must relinquish control of the keyboard.
' This value will be set to TRUE if the API edit box currently has focus.
Private m_HasFocus As Boolean

'Because our API edit box is not tied into VB's default tab stop handling, we must jump through some hoops to forward focus correctly.
' Our hook proc will capture the Tab key that causes focus to enter the control, and mistakenly assume it is a Tab keypress from
' *within* the control.  To prevent this from happening, we enforce a slight time delay from when our hook procedure begins, to when
' we capture Tab keypresses.  This prevents faulty Tab-key handling.
Private m_TimeAtFocusEnter As Long

'The system handles drawing of the edit box.  These persistent brush and font handles are passed to the relevant
' window messages, and WAPI uses them when rendering the edit box and associated text.
Private m_EditBoxBrush As Long, m_EditBoxFont As Long

'While inside the hook event, this will be set to TRUE.  Because we raise events directly from the hook, we sometimes need to postpone
' crucial actions (like releasing the hook) until the hook proc has exited.
Private m_InHookNow As Boolean

'If the user attempts to set the Text property before the edit box is created (e.g. when the control is invisible), we will back up the
' text to this string.  When the edit box is created, this text will be automatically placed inside the control.
Private m_TextBackup As String

'Basic control constants, cached here for easier tracking
Private m_Enabled As Boolean, m_Visible As Boolean
Private m_BackColor As Long, m_TextColor As Long

'If something forces us to release our keyboard hook while in the midst of a hook, we want to delay the request until
' the hook exits.  If we don't do this, we will crash.
Private WithEvents m_HookReleaseTimer As pdTimer
Attribute m_HookReleaseTimer.VB_VarHelpID = -1

Public Property Get BackColor() As Long
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal newColor As Long)
    m_BackColor = Colors.ConvertSystemColor(newColor)
    If m_EditBoxHwnd <> 0 Then CreateEditBoxGDIObjects
End Property

Public Property Get Enabled() As Boolean
    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal newValue As Boolean)
    m_Enabled = newValue
    If m_EditBoxHwnd <> 0 Then EnableWindow m_EditBoxHwnd, IIf(m_Enabled, 1, 0)
End Property

'Font properties; only a subset are used, as PD handles most font settings automatically
Public Property Get FontSize() As Single
    FontSize = m_FontSize
End Property

'Edit box sizes and margins are typically enforced by the system, at creation time, so it's easiest to just recreate
' the box from scratch if the font settings change.  (In other words, try to avoid font changes at run-time.)
Public Property Let FontSize(ByVal newSize As Single)
    If newSize <> m_FontSize Then
        m_FontSize = newSize
        CreateEditBox m_ParentHWnd
    End If
End Property

Public Property Get HasFocus() As Boolean
    HasFocus = m_HasFocus
End Property

'FYI: if the returned hWnd is zero, an API edit box has not been created yet
Public Property Get hWnd() As Long
    hWnd = m_EditBoxHwnd
End Property

Public Property Get Multiline() As Boolean
    Multiline = m_Multiline
End Property

'Changing the multiline property requires a full recreation of the edit box (e.g. it cannot be changed via
' window message alone).  Changing the multiline property enforces a number of different edit boxes behaviors,
' including non-obvious ones (like text margins), so avoid run-time changes if you can.
Public Property Let Multiline(ByVal newState As Boolean)
    If newState <> m_Multiline Then
        m_Multiline = newState
        CreateEditBox m_ParentHWnd
    End If
End Property

'External functions can call this to fully select the text box's contents
Public Sub SelectAll()
    If m_EditBoxHwnd <> 0 Then SendMessage m_EditBoxHwnd, EM_SETSEL, ByVal 0&, ByVal -1&
End Sub

'SelStart is used by some PD functions to control caret positioning after automatic text updates
' (as used in the spinner contol, for example)
Public Property Get SelStart() As Long
    If m_EditBoxHwnd <> 0 Then
        Dim startPos As Long, endPos As Long, retVal As Long
        retVal = SendMessage(m_EditBoxHwnd, EM_GETSEL, startPos, endPos)
        SelStart = startPos
    End If
End Property

Public Property Let SelStart(ByVal newPosition As Long)
    If m_EditBoxHwnd <> 0 Then SendMessage m_EditBoxHwnd, EM_SETSEL, newPosition, newPosition
End Property

'For performance reasons, the text is not always stored locally.  It is retrieved, as needed, from the common control instance.
Public Property Get Text() As String
    
    'If the API edit box has been created, retrieve the latest text directly from it
    If m_EditBoxHwnd <> 0 Then
    
        'Retrieve the length of the edit box's current text.  Note that this is not necessarily the *actual* length.
        ' Instead, it is an interop-friendly measurement that represents the maximum possible size of the buffer,
        ' when accounting for mixed ANSI and Unicode chars (among other things).
        Dim bufLength As Long
        bufLength = GetWindowTextLength(m_EditBoxHwnd) + 1
        
        'Prepare a local string buffer at that size
        Text = Space$(bufLength)
        
        'Retrieve the text.  Note that the retrieval function will return that actual length of the buffer (not counting
        ' the null terminator).  On the off chance that the actual length differs from the buffer we were initially given,
        ' trim the string to match.
        Dim actualBufLength As Long
        actualBufLength = GetWindowText(m_EditBoxHwnd, StrPtr(Text), bufLength)
        If actualBufLength <> bufLength Then Text = Left$(Text, actualBufLength)
    
    'If an API edit box has not been created, return any text the caller may have previously cached
    Else
        Text = m_TextBackup
    End If
    
End Property

Public Property Let Text(ByRef newString As String)

    'Unfortunately, we cannot use SetWindowText here.  SetWindowText does not expand tab characters in a string,
    ' so our only option is to manually send a WM_SETTEXT message to the text box.
    If m_EditBoxHwnd <> 0 Then
        SendMessage m_EditBoxHwnd, WM_SETTEXT, 0&, ByVal StrPtr(newString)
    
    'If an edit box doesn't exist, cache the string locally.  (We'll automatically set it anew whenever we get around
    ' to creating the edit box.)
    Else
        m_TextBackup = newString
    End If
    
    If g_IsProgramRunning Then RaiseEvent Change
    
End Property

Public Property Get textColor() As Long
    textColor = m_TextColor
End Property

Public Property Let textColor(ByVal newColor As Long)
    m_TextColor = Colors.ConvertSystemColor(newColor)
    If m_EditBoxHwnd <> 0 Then CreateEditBoxGDIObjects
End Property

Public Property Get Visible() As Boolean
    Visible = m_Visible
End Property

Public Property Let Visible(ByVal newVisibility As Boolean)
    m_Visible = newVisibility
    If m_Visible Then
        If m_EditBoxHwnd = 0 Then
            CreateEditBox m_ParentHWnd
        Else
            ShowWindow m_EditBoxHwnd, SW_SHOWNA
        End If
    Else
        If m_EditBoxHwnd <> 0 Then ShowWindow m_EditBoxHwnd, SW_HIDE
    End If
End Property

'In order to give our owner some control over when the edit box is actually created, they are required to call the
' CreateEditBox function at least once.  (The Show() event is a decent place, for example.)  After the edit box has
' been created at least once, it will re-create itself as necessary (e.g. if property changes require it).
'
'Note that (x, y) values for the edit box can be specified.  These are technically optional, but you DO need to supply
' a valid x2 and y2 if you want the edit box to actually be visible!  Alternatively, you can set the optional
' "determineHeightForMe" parameter to TRUE to have the edit box size itself vertically, according to the current font.
Public Function CreateEditBox(ByVal parentHwnd As Long, Optional ByVal x1 As Long = 0, Optional ByVal y1 As Long = 0, Optional ByVal x2 As Long = 0, Optional ByVal y2 As Long = 0, Optional ByVal determineHeightForMe As Boolean = True) As Boolean
    
    'Ignore all creation requests until we receive one with a valid parent hWnd
    If parentHwnd <> 0 Then
    
        'If the edit box already exists, copy its text prior to killing it
        Dim curText As String
        If m_EditBoxHwnd <> 0 Then curText = Text Else curText = m_TextBackup
        DestroyEditBox
        
        'The edit box requires some GDI handles so it can auto-paint its own backgruond and text.  Create those objects now.
        CreateEditBoxGDIObjects
        
        'Some window style flags vary on user-editable properties; calculate those now
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
        
        'Multiline text boxes can have any height.  Single-line text boxes should not; we prefer to force them to their
        ' ideal height, using the current font as our guide.
        If determineHeightForMe Then y2 = y1 + Me.SuggestedHeight
        
        'Create the window!
        m_EditBoxHwnd = CreateWindowEx(flagsWinStyleExtended, ByVal StrPtr("EDIT"), ByVal StrPtr(""), flagsWinStyle Or flagsEditControl, _
            x1, y1, x2, y2, parentHwnd, 0, App.hInstance, ByVal 0&)
        
        'Now that the window exists, we need to pass it any relevant local settings
        Me.Enabled = m_Enabled
        Me.Visible = m_Visible
        
        'We need to subclass a number of different messages, both for the edit box window and for the parent control.
        If g_IsProgramRunning Then
            If Not (cSubclass Is Nothing) Then
                
                'Subclass the edit box
                cSubclass.ssc_Subclass m_EditBoxHwnd, 0, 1, Me, True, True, True
                cSubclass.ssc_AddMsg m_EditBoxHwnd, MSG_BEFORE, WM_KEYDOWN, WM_SETFOCUS, WM_KILLFOCUS, WM_CHAR, WM_UNICHAR, WM_MOUSEACTIVATE
                
                'Subclass the user control as well.  This is necessary for handling update messages from the edit box
                If Not m_ParentHasBeenSubclassed Then
                    m_ParentHWnd = parentHwnd
                    cSubclass.ssc_Subclass parentHwnd, 0, 1, Me, True, True, False
                    cSubclass.ssc_AddMsg parentHwnd, MSG_BEFORE, WM_CTLCOLOREDIT, WM_CTLCOLORSTATIC, WM_COMMAND
                    m_ParentHasBeenSubclassed = True
                End If
                
            End If
        End If
        
        'Pass an appropriate font handle to the edit box; GDI will use this font for text rendering.
        If m_EditBoxHwnd <> 0 Then SendMessage m_EditBoxHwnd, WM_SETFONT, m_EditBoxFont, IIf(m_Visible, 1, 0)
        
        'If the edit box had text before we killed it, restore that text now
        If Len(curText) <> 0 Then Text = curText
        
    End If
    
    'Return TRUE if we managed to create an API window
    CreateEditBox = CBool(m_EditBoxHwnd <> 0)
    
End Function

'The caller is welcome to kill the edit box at any time.  Relevant properties are stored in this class, so the actual
' API window associated with the class can be created/killed at your leisure.  (Killing the window instead of hiding
' it frees extra resources, for example.)
'
'NOTE: this function does not destroy associated resources, like GDI brushes or fonts.  Those are shared among all
'      pdEditBox instances, so the only way to free them is to free all pdEditBox instances throughout PD.
Public Function DestroyEditBox() As Boolean
    If m_EditBoxHwnd <> 0 Then
        If Not (cSubclass Is Nothing) Then
            cSubclass.ssc_UnSubclass m_EditBoxHwnd
            cSubclass.shk_TerminateHooks
        End If
        DestroyEditBox = CBool(DestroyWindow(m_EditBoxHwnd) <> 0)
    Else
        DestroyEditBox = True
    End If
End Function

Friend Function GetPositionRect(ByRef dstPositionRect As winRect) As Boolean
    GetWindowRect m_EditBoxHwnd, dstPositionRect
    GetPositionRect = CBool(m_EditBoxHwnd <> 0)
End Function

Public Function GetWidth() As Long
    Dim tmpRect As winRect
    GetWindowRect m_EditBoxHwnd, tmpRect
    GetWidth = tmpRect.x2 - tmpRect.x1
End Function

Public Function GetHeight() As Long
    Dim tmpRect As winRect
    GetWindowRect m_EditBoxHwnd, tmpRect
    GetHeight = tmpRect.y2 - tmpRect.y1
End Function

'Change the edit box position and size in one fell swoop.  Coordinates are RELATIVE TO THE PARENT WINDOW.
Public Sub Move(ByVal newLeft As Long, ByVal newTop As Long, ByVal newWidth As Long, ByVal newHeight As Long)
    If m_EditBoxHwnd <> 0 Then MoveWindow m_EditBoxHwnd, newLeft, newTop, newWidth, newHeight, 1&
End Sub

'This function actually does two things: makes the edit box the foreground window, then assigns it keyboard focus.
' By design, all non-multiline edit boxes in PD select all their own text upon an initial click.
Public Sub SetFocusToEditBox()
    If m_EditBoxHwnd <> 0 Then
        SetForegroundWindow m_EditBoxHwnd
        SetFocus m_EditBoxHwnd
        If Not m_Multiline Then Me.SelectAll
    End If
End Sub

'Given the current font size, return the suggested height of an edit box.  Parent controls can use this to auto-size
' themselves in preparation of actually creating the API window.
Public Function SuggestedHeight() As Long
    SuggestedHeight = Font_Management.GetDefaultStringHeight(m_FontSize)
End Function

'Create the brush and font used to render the edit box background and text
Private Sub CreateEditBoxGDIObjects()

    If m_EditBoxBrush <> 0 Then UserControl_Support.ReleaseSharedGDIBrushByHandle m_EditBoxBrush
    If m_EditBoxFont <> 0 Then UserControl_Support.ReleaseSharedGDIFontByHandle m_EditBoxFont
    
    If g_IsProgramRunning Then
        m_EditBoxBrush = UserControl_Support.GetSharedGDIBrush(Me.BackColor)
    Else
        m_EditBoxBrush = UserControl_Support.GetSharedGDIBrush(RGB(0, 255, 0))
    End If
    
    m_EditBoxFont = UserControl_Support.GetSharedGDIFont(m_FontSize)
    If m_EditBoxHwnd <> 0 Then SendMessage m_EditBoxHwnd, WM_SETFONT, m_EditBoxFont, IIf(m_Visible, 1, 0)

End Sub

'Force the edit box to re-paint itself
Private Sub RepaintEditBox()
    If m_EditBoxHwnd <> 0 Then
        Dim tmpRect As winRect
        GetWindowRect m_EditBoxHwnd, tmpRect
        InvalidateRect m_EditBoxHwnd, tmpRect, 1
        UpdateWindow m_EditBoxHwnd
    End If
End Sub

Private Sub Class_Initialize()
    Set m_HookReleaseTimer = New pdTimer
    If g_IsProgramRunning Then Set cSubclass = New cSelfSubHookCallback
    m_Enabled = True
    m_Visible = True
End Sub

Private Sub Class_Terminate()

    'Release any GDI objects associated with the edit box
    If m_EditBoxBrush <> 0 Then
        UserControl_Support.ReleaseSharedGDIBrushByHandle m_EditBoxBrush
        m_EditBoxBrush = 0
    End If
    
    If m_EditBoxFont <> 0 Then
        UserControl_Support.ReleaseSharedGDIFontByHandle m_EditBoxFont
        m_EditBoxFont = 0
    End If
    
    'Destroy the edit box, as necessary
    DestroyEditBox
    
    'Release our extra subclasser(s)
    If Not cSubclass Is Nothing Then cSubclass.ssc_Terminate
    
End Sub

'If a hook is active, this timer will repeatedly try to kill it.  Do not enable it until you are certain the hook
' needs to be released.  (This is used as a failsafe if we cannot immediately release the hook when focus is lost,
' for example if we are currently inside an external event, as happens in the Layer toolbox, which hides the active
' text box inside the KeyPress event.)
Private Sub m_HookReleaseTimer_Timer()
    If (Not m_InHookNow) Then
        RemoveKeyboardHook
        m_HookReleaseTimer.StopTimer
    End If
End Sub

'If a virtual key code is numeric, return TRUE.
Private Function IsVirtualKeyNumeric(ByVal vKey As Long, Optional ByRef numericValue As Long = 0) As Boolean
    
    If (vKey >= VK_0) And (vKey <= VK_9) Then
        IsVirtualKeyNumeric = True
        numericValue = vKey - VK_0
    Else
    
        If (vKey >= VK_NUMPAD0) And (vKey <= VK_NUMPAD9) Then
            IsVirtualKeyNumeric = True
            numericValue = vKey - VK_NUMPAD0
        Else
            IsVirtualKeyNumeric = False
        End If
    End If
    
End Function

'Given a virtual keycode, return TRUE if the keycode is a command key that must be manually forwarded to an edit box.  Command keys include
' arrow keys at present, but in the future, additional keys can be added to this list.
'
'NOTE FOR OUTSIDE USERS!  These key constants are declared publicly in PD, because they are used many places.  You can find virtual keycode
' declarations in PD's Public_Constants module, or at this MSDN link:
' http://msdn.microsoft.com/en-us/library/windows/desktop/dd375731%28v=vs.85%29.aspx
Private Function DoesVirtualKeyRequireSpecialHandling(ByVal vKey As Long) As Boolean
    
    Select Case vKey
    
        Case VK_LEFT, VK_UP, VK_RIGHT, VK_DOWN
            DoesVirtualKeyRequireSpecialHandling = True
            
        Case VK_TAB
            DoesVirtualKeyRequireSpecialHandling = True
        
        Case Else
            DoesVirtualKeyRequireSpecialHandling = False
    
    End Select
        
End Function

'Note that the vKey constant below is a virtual key mapping, not necessarily a standard VB key constant
Private Function IsVirtualKeyDown(ByVal vKey As Long) As Boolean
    IsVirtualKeyDown = GetAsyncKeyState(vKey) And &H8000
End Function

'Install a keyboard hook in our window
Private Sub InstallKeyboardHook()

    'Check for an existing hook
    If Not m_HasFocus Then
    
        'Note that this window is now active
        m_HasFocus = True
        
        'Note the time.  This is used to prevent keypresses occurring immediately prior to the hook, from being
        ' caught within our hook proc!
        m_TimeAtFocusEnter = GetTickCount
        
        'No hook exists.  Hook the control now.
        If Not (cSubclass Is Nothing) Then cSubclass.shk_SetHook WH_KEYBOARD, False, MSG_BEFORE, m_EditBoxHwnd, 2, Me, , True
        
    End If

End Sub

'If a hook exists, uninstall it.  DO NOT CALL THIS FUNCTION if the class is currently inside the hook proc.
Private Sub RemoveKeyboardHook()
    If Not (cSubclass Is Nothing) Then cSubclass.shk_UnHook WH_KEYBOARD
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
                    If DoesVirtualKeyRequireSpecialHandling(wParam) Then
                        
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
                            If IsVirtualKeyNumeric(wParam, numCheck) Then
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
                If DoesVirtualKeyRequireSpecialHandling(wParam) Then
                    
                    'On a single-line control, the tab key should be used to redirect focus to a new window.
                    If (wParam = VK_TAB) Then
                        
                        'Multiline edit boxes accept tab keypresses.  Single line ones do not, so interpret TAB as a
                        ' request to forward (or reverse) focus.
                        If (Not m_Multiline) And ((GetTickCount - m_TimeAtFocusEnter) > 250) Then
                            
                            'Raise the special tab key event
                            RaiseEvent TabPress(Not IsVirtualKeyDown(VK_SHIFT))
                            
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
        
        'The parent receives this message prior to an ENABLED edit box being drawn.  The parent can use this to assign text and
        ' background colors to the edit box.
        Case WM_CTLCOLOREDIT
            
            'Make sure the command is relative to *our* edit box, and not another one
            If lParam = m_EditBoxHwnd Then
            
                'Assign colors directly, using the API
                SetTextColor wParam, Me.textColor
                SetBkMode wParam, SBKM_OPAQUE
                SetBkColor wParam, Me.BackColor
                
                'We return the background brush
                bHandled = True
                lReturn = m_EditBoxBrush
                
            End If
        
        'The parent receives this message prior to a DISABLED edit box being drawn.  The parent can use this to assign text and
        ' background colors to the edit box.
        Case WM_CTLCOLORSTATIC
        
            'Make sure the command is relative to *our* edit box, and not another one
            If lParam = m_EditBoxHwnd Then
            
                'Assign colors directly, using the API
                SetTextColor wParam, Me.textColor
                SetBkMode wParam, SBKM_OPAQUE
                SetBkColor wParam, Me.BackColor
                
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
            
        'WM_UNICHAR messages are never sent by Windows.  However, third-party IMEs may send them.  Before allowing WM_UNICHAR
        ' messages to pass, Windows will first probe a window by sending the UNICODE_NOCHAR value.  If a window responds with 1
        ' (instead of 0, as DefWindowProc does), Windows will allow WM_UNICHAR messages to pass.
        Case WM_UNICHAR
            If wParam = UNICODE_NOCHAR Then
                bHandled = True
                lReturn = 1
            End If
            
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
            If (Not m_HasFocus) And (m_ParentHWnd <> 0) Then SetFocus m_ParentHWnd
            
        'When the control receives focus, initialize a keyboard hook.  This prevents accelerators from working, but it is the
        ' only way to bypass VB's internal message translator, which will forcibly convert certain Unicode chars to ANSI.
        Case WM_SETFOCUS
            
            'Forcibly disable PD's main accelerator control
            FormMain.pdHotkeys.DeactivateHook
            
            'Start hooking keypresses so we can grab Unicode chars before VB eats 'em.  Not that this function actually sets
            ' the m_HasFocus value to TRUE.
            If Not m_HasFocus Then
                InstallKeyboardHook
                RaiseEvent GotFocusAPI
            End If
            
        Case WM_KILLFOCUS
            
            'Release our hook.  In some circumstances, we can't do this immediately, so we set a timer that will release the hook
            ' as soon as the system allows.
            If m_InHookNow Then
                m_HookReleaseTimer.StartTimer
            Else
                RemoveKeyboardHook
            End If
            
            'Re-enable PD's main accelerator control
            FormMain.pdHotkeys.ActivateHook
            
            'Immediately mark the control-wide focus state
            If m_HasFocus Then
                m_HasFocus = False
                RaiseEvent LostFocusAPI
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



