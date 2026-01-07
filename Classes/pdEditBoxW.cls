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
'Copyright 2014-2026 by Tanner Helland
'Created: 03/November/14
'Last updated: 10/May/19
'Last update: raise key events for arrow keys; in rare circumstances (e.g. pdSearchBox) our parent window
'              may want to do something special with these
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
' 2) To allow use of arrow keys and other control keys, this control must hook the keyboard.  (If it does not,
'    any parent VB window will eat control keypresses, because it doesn't know about windows created via WAPI.)
' 3) Dynamic hooking can occasionally cause trouble in the IDE, particularly when used with break points.  Avoid 'em.
' 4) This class relies on a number of functions in the UserControls module, such as sharing GDI brushes and
'    fonts between all pdEditBoxW instances (to reduce GDI object count).
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'If a property change requires the creation of an API window (e.g. visible = true), but the control has not yet been created,
' it will raise this sevent, notifying the parent to call the Create function immediately.
Public Event CreateTextBoxNow()

'By design, PD textboxes raises fewer events than standard text boxes
Public Event Change()
Public Event KeyPress(ByVal Shift As ShiftConstants, ByVal vKey As Long, ByRef preventFurtherHandling As Boolean)
Public Event KeyDown(ByVal Shift As ShiftConstants, ByVal vKey As Long, ByRef preventFurtherHandling As Boolean)
Public Event KeyUp(ByVal Shift As ShiftConstants, ByVal vKey As Long, ByRef preventFurtherHandling As Boolean)
Public Event Resize()
Public Event GotFocusAPI()
Public Event LostFocusAPI()

'While this control doesn't have any internal use for mouse events, it can relay them to its owner.  This allows PD to make
' UI rendering decisions based on the mouse state of the edit box (e.g. special borders on hover vs inactive).
Public Event MouseLeave(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
Public Event MouseEnter(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
Public Event MouseHover(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
Public Event MouseWheelVertical(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal scrollAmount As Double)
Public Event MouseWheelHorizontal(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal scrollAmount As Double)
Public Event MouseWheelZoom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal zoomAmount As Double)
Private WithEvents m_MouseEvents As pdInputMouse
Attribute m_MouseEvents.VB_VarHelpID = -1

'Window styles
Private Enum Win32_WindowStyles
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
    WS_EX_MDICHILD = &H40&
    WS_EX_WINDOWEDGE = &H100&
    WS_EX_CLIENTEDGE = &H200&
    WS_EX_CONTEXTHELP = &H400&
    WS_EX_RIGHT = &H1000&
    WS_EX_LEFT = &H0&
    WS_EX_RTLREADING = &H2000&
    WS_EX_LTRREADING = &H0&
    WS_EX_LEFTSCROLLBAR = &H4000&
    WS_EX_RIGHTSCROLLBAR = &H0&
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

Private Enum Win32_EditBoxStyles
    ES_AUTOHSCROLL = &H80
    ES_AUTOVSCROLL = &H40
    ES_CENTER = &H1
    ES_LEFT = &H0
    ES_LOWERCASE = &H10
    ES_MULTILINE = &H4
    ES_NOHIDESEL = &H100
    ES_NUMBER = &H2000
    ES_OEMCONVERT = &H400
    ES_PASSWORD = &H20
    ES_READONLY = &H800
    ES_RIGHT = &H2
    ES_UPPERCASE = &H8
    ES_WANTRETURN = &H1000
End Enum

#If False Then
    Private Const ES_AUTOHSCROLL = &H80, ES_AUTOVSCROLL = &H40, ES_CENTER = &H1, ES_LEFT = &H0, ES_LOWERCASE = &H10, ES_MULTILINE = &H4, ES_NOHIDESEL = &H100, ES_NUMBER = &H2000, ES_OEMCONVERT = &H400, ES_PASSWORD = &H20, ES_READONLY = &H800, ES_RIGHT = &H2, ES_UPPERCASE = &H8, ES_WANTRETURN = &H1000
#End If

'Updating the font is done via WM_SETFONT
Private Const WM_SETFONT = &H30
Private Declare Function SetBkMode Lib "gdi32" (ByVal targetDC As Long, ByVal nBkMode As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal targetDC As Long, ByVal nBkColor As Long) As Long
Private Const SBKM_OPAQUE = 2

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

'To avoid VB's ANSI message pump, we need to peek at character codes before VB eats them
Private Type TagMSG
    hWnd As Long
    msgID As Long
    wParam As Long
    lParam As Long
    msgTime As Long
    msgPoint As PointAPI
End Type

'System window handling APIs
Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExW" (ByVal dwExStyle As Long, ByVal lpClassName As Long, ByVal lpWindowName As Long, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByRef lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal bEnable As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hndWindow As Long, ByRef lpRect As winRect) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hndWindow As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function SetFocus Lib "user32" (ByVal hndWindow As Long) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hndWindow As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hndWindow As Long, ByVal nCmdShow As ShowWindowOptions) As Long

'Getting/setting window data
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextW" (ByVal hWnd As Long, ByVal lpStringPointer As Long, ByVal nMaxCount As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthW" (ByVal hWnd As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long

'Message routing
Private m_CharCache As Long
Private Const PM_NOREMOVE As Long = &H0
Private Declare Function PeekMessage Lib "user32" Alias "PeekMessageW" (ByRef lpMsg As TagMSG, ByVal hWnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long

'Handle to the system edit box wrapped by this control
Private m_EditBoxHwnd As Long

Private Const WM_KEYDOWN As Long = &H100
Private Const WM_KEYUP As Long = &H101
Private Const WM_CHAR As Long = &H102
Private Const WM_IME_CHAR As Long = &H286
Private Const WM_UNICHAR As Long = &H109
Private Const UNICODE_NOCHAR As Long = &HFFFF&
Private Const WM_SETFOCUS As Long = &H7
Private Const WM_KILLFOCUS As Long = &H8
Private Const WM_SETTEXT As Long = &HC
Private Const WM_COMMAND As Long = &H111
Private Const WM_CTLCOLOREDIT As Long = &H133
Private Const WM_CTLCOLORSTATIC As Long = &H138
Private Const WM_DESTROY As Long = &H2
Private Const EN_UPDATE As Long = &H400
Private Const EM_GETSEL As Long = &HB0
Private Const EM_SETSEL As Long = &HB1

'Obviously, we're going to be doing a lot of subclassing inside this control.
Implements ISubclass
Private m_EditBoxSubclassed As Boolean

'Unlike the edit box, which may be recreated multiple times as properties change, we only need to subclass the parent window once.
' After it has been subclassed, this will be set to TRUE.
Private m_ParentHasBeenSubclassed As Boolean, m_ParentHWnd As Long

'Dynamic hooking requires us to track focus events with care.  When focus is lost, we must relinquish control of the keyboard.
' This value will be set to TRUE if the API edit box currently has focus.
Private m_HasFocus As Boolean

'The system handles drawing of the edit box.  These persistent brush and font handles are passed to the relevant
' window messages, and WAPI uses them when rendering the edit box and associated text.
Private m_EditBoxBrush As Long, m_EditBoxFont As Long

'If the user attempts to set the Text property before the edit box is created (e.g. when the control is invisible), we will back up the
' text to this string.  When the edit box is created, this text will be automatically placed inside the control.
Private m_TextBackup As String

'Basic control constants, cached here for easier tracking
Private m_Enabled As Boolean, m_Visible As Boolean
Private m_BackColor As Long, m_TextColor As Long

'On the first MOUSE_UP event after the control receives focus, we select all text if the user did not make a manual text selection.
Private m_JustGotFocus As Boolean

'Keyboard hooking is used to bypass VB's propensity for eating arrow keypresses
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Private m_HookID As Long

'Our parent can query last key-code and scan-code values (from either key up or down)
Private m_lastScanCodeDown As Long, m_lastKeyCodeDown As Long
Private m_lastScanCodeUp As Long, m_lastKeyCodeUp As Long

Friend Property Get BackColor() As Long
    BackColor = m_BackColor
End Property

Friend Property Let BackColor(ByVal newBackColor As Long)
    newBackColor = Colors.ConvertSystemColor(newBackColor)
    If (newBackColor <> m_BackColor) Then
        m_BackColor = newBackColor
        If (m_EditBoxHwnd <> 0) Then CreateEditBoxGDIObjects
    End If
End Property

Friend Property Get Enabled() As Boolean
    Enabled = m_Enabled
End Property

Friend Property Let Enabled(ByVal newValue As Boolean)
    m_Enabled = newValue
    If (m_EditBoxHwnd <> 0) Then
        EnableWindow m_EditBoxHwnd, IIf(m_Enabled, 1, 0)
        CreateEditBoxGDIObjects True
    End If
End Property

'Font properties; only a subset are used, as PD handles most font settings automatically
Friend Property Get FontSize() As Single
    FontSize = m_FontSize
End Property

'Edit box sizes and margins are typically enforced by the system, at creation time, so I do not advise changing the font size of an
' edit box *after* it's been created.
Friend Property Let FontSize(ByVal newSize As Single)
    If (newSize <> m_FontSize) Then
        m_FontSize = newSize
        CreateEditBoxGDIObjects True
    End If
End Property

Friend Property Get HasFocus() As Boolean
    HasFocus = m_HasFocus
End Property

'FYI: if the returned hWnd is zero, an API edit box has not been created yet
Friend Property Get hWnd() As Long
    hWnd = m_EditBoxHwnd
End Property

Friend Property Get Multiline() As Boolean
    Multiline = m_Multiline
End Property

'Changing the multiline property requires a full recreation of the edit box (e.g. it cannot be changed via
' window message alone).  Changing the multiline property enforces a number of different edit boxes behaviors,
' including non-obvious ones (like text margins), so avoid run-time changes if you can.
Friend Property Let Multiline(ByVal newState As Boolean)
    If (newState <> m_Multiline) Then
        m_Multiline = newState
        If (m_EditBoxHwnd <> 0) Then CreateEditBox m_ParentHWnd
    End If
End Property

'External functions can call this to fully select the text box's contents
Friend Sub SelectAll()
    If (m_EditBoxHwnd <> 0) Then VBHacks.SendMsgW m_EditBoxHwnd, EM_SETSEL, 0&, -1&
End Sub

'SelLength is currently treated as read-only
Friend Property Get SelLength() As Long
    If (m_EditBoxHwnd <> 0) Then
        Dim startPos As Long, endPos As Long, retVal As Long
        retVal = VBHacks.SendMsgW(m_EditBoxHwnd, EM_GETSEL, VarPtr(startPos), VarPtr(endPos))
        SelLength = Abs(startPos - endPos)
    End If
End Property

'SelStart is used by some PD functions to control caret positioning after automatic text updates
' (as used in the spinner contol, for example)
Friend Property Get SelStart() As Long
    If (m_EditBoxHwnd <> 0) Then
        Dim startPos As Long, endPos As Long, retVal As Long
        retVal = VBHacks.SendMsgW(m_EditBoxHwnd, EM_GETSEL, VarPtr(startPos), VarPtr(endPos))
        SelStart = startPos
    End If
End Property

Friend Property Let SelStart(ByVal newPosition As Long)
    If (m_EditBoxHwnd <> 0) Then VBHacks.SendMsgW m_EditBoxHwnd, EM_SETSEL, newPosition, newPosition
End Property

'For performance reasons, the text is not always stored locally.  It is retrieved, as needed, from the common control instance.
Friend Property Get Text() As String
    
    'If the API edit box has been created, retrieve the latest text directly from it
    If (m_EditBoxHwnd <> 0) Then
    
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
        If (actualBufLength <> bufLength) Then Text = Left$(Text, actualBufLength)
        
    'If an API edit box has not been created, return any text the caller may have previously cached
    Else
        Text = m_TextBackup
    End If
    
End Property

Friend Property Let Text(ByRef newString As String)

    'Unfortunately, we cannot use SetWindowText here.  SetWindowText does not expand tab characters in a string,
    ' so our only option is to manually send a WM_SETTEXT message to the text box.
    If (m_EditBoxHwnd <> 0) Then
        VBHacks.SendMsgW m_EditBoxHwnd, WM_SETTEXT, 0&, StrPtr(newString)
    
    'If an edit box doesn't exist, cache the string locally.  (We'll automatically set it anew whenever we get around
    ' to creating the edit box.)
    Else
        m_TextBackup = newString
    End If
    
    If PDMain.IsProgramRunning() Then RaiseEvent Change
    
End Property

Friend Property Get TextColor() As Long
    TextColor = m_TextColor
End Property

Friend Property Let TextColor(ByVal newTextColor As Long)
    newTextColor = Colors.ConvertSystemColor(newTextColor)
    If (newTextColor <> m_TextColor) Then
        m_TextColor = newTextColor
        If (m_EditBoxHwnd <> 0) Then CreateEditBoxGDIObjects
    End If
End Property

Friend Property Get Visible() As Boolean
    Visible = m_Visible
End Property

Friend Property Let Visible(ByVal newVisibility As Boolean)
    m_Visible = newVisibility
    If m_Visible And PDMain.IsProgramRunning() Then
        If (m_EditBoxHwnd = 0) Then RaiseEvent CreateTextBoxNow
        ShowWindow m_EditBoxHwnd, SW_SHOWNA
    Else
        If (m_EditBoxHwnd <> 0) Then ShowWindow m_EditBoxHwnd, SW_HIDE
    End If
End Property

'In order to give our owner some control over when the edit box is actually created, they are required to call the
' CreateEditBox function at least once.  (The Show() event is a decent place, for example.)  After the edit box has
' been created at least once, it will re-create itself as necessary (e.g. if property changes require it).
'
'Note that (x, y) values for the edit box can be specified.  These are technically optional, but you DO need to supply
' a valid x2 and y2 if you want the edit box to actually be visible!  Alternatively, you can set the optional
' "determineHeightForMe" parameter to TRUE to have the edit box size itself vertically, according to the current font.
Friend Function CreateEditBox(ByVal parentHwnd As Long, Optional ByVal x1 As Long = 0, Optional ByVal y1 As Long = 0, Optional ByVal xWidth As Long = 0, Optional ByVal yHeight As Long = 0, Optional ByVal determineHeightForMe As Boolean = True) As Boolean
    
    'Ignore all creation requests until we receive one with a valid parent hWnd
    If (parentHwnd <> 0) Then
    
        'If the edit box already exists, copy its text prior to killing it
        Dim curText As String
        If (m_EditBoxHwnd <> 0) Then curText = Me.Text Else curText = m_TextBackup
        DestroyEditBox
        
        'Some window style flags vary on user-editable properties; calculate those now
        Dim flagsWinStyle As Long, flagsWinStyleExtended As Long, flagsEditControl As Win32_EditBoxStyles
        flagsWinStyle = WS_VISIBLE Or WS_CHILD
        flagsWinStyleExtended = 0
        flagsEditControl = 0
        
        If m_Multiline Then
            flagsWinStyle = flagsWinStyle Or WS_VSCROLL
            flagsEditControl = flagsEditControl Or ES_MULTILINE Or ES_WANTRETURN Or ES_AUTOVSCROLL
        Else
            flagsEditControl = flagsEditControl Or ES_AUTOHSCROLL
        End If
        
        'Multiline text boxes can have any height.  Single-line text boxes should not; we prefer to force them to their
        ' ideal height, using the current font as our guide.
        If determineHeightForMe Then yHeight = Me.SuggestedHeight
        
        'Create the window!
        If PDMain.IsProgramRunning() Then
            m_EditBoxHwnd = CreateWindowEx(flagsWinStyleExtended, StrPtr("EDIT"), 0&, flagsWinStyle Or flagsEditControl, _
                x1, y1, xWidth, yHeight, parentHwnd, 0&, App.hInstance, ByVal 0&)
            If (m_EditBoxHwnd <> 0) Then UserControls.NotifyAPIWindowCreated
        End If
        
        If (m_EditBoxHwnd <> 0) Then
        
            'Now that the window exists, we need to pass it any relevant local settings
            Me.Enabled = m_Enabled
            Me.Visible = m_Visible
            
            'We need to subclass a number of different messages, both for the edit box window and for the parent control.
            If PDMain.IsProgramRunning() Then
                
                'Subclass the edit box itself
                m_EditBoxSubclassed = VBHacks.StartSubclassing(m_EditBoxHwnd, Me)
                
                'Subclass the edit box's parent window.  This is necessary for handling certain update messages.
                If (Not m_ParentHasBeenSubclassed) Then
                    m_ParentHWnd = parentHwnd
                    m_ParentHasBeenSubclassed = VBHacks.StartSubclassing(parentHwnd, Me)
                End If
                
                'Add a mouse handler (which performs its own subclassing)
                If (m_MouseEvents Is Nothing) Then Set m_MouseEvents = New pdInputMouse
                m_MouseEvents.AddInputTracker m_EditBoxHwnd, , True
                m_MouseEvents.SetDefWindowProcFallback True
                
            End If
            
            'Pass an appropriate font handle to the edit box; GDI will use this font for text rendering.
            CreateEditBoxGDIObjects
            If (m_EditBoxFont <> 0) Then VBHacks.SendMsgW m_EditBoxHwnd, WM_SETFONT, m_EditBoxFont, IIf(m_Visible, 1, 0)
            
            'If the edit box had text before we killed it, restore that text now
            If (LenB(curText) <> 0) Then Me.Text = curText
            
        End If
        
    End If
    
    'Return TRUE if we managed to create an API window
    CreateEditBox = (m_EditBoxHwnd <> 0)
    
End Function

'The caller is welcome to kill the edit box at any time.  Relevant properties are stored in this class, so the actual
' API window associated with the class can be created/killed at your leisure.  (Killing the window instead of hiding
' it frees extra resources, for example.)
'
'NOTE: this function does not destroy associated resources, like GDI brushes or fonts.  Those are shared among all
'      pdEditBox instances, so the only way to free them is to free all pdEditBox instances throughout PD.
Friend Function DestroyEditBox() As Boolean
    
    ReleaseEditBoxGDIObjects
    
    If (m_EditBoxHwnd <> 0) Then
    
        If (Not m_MouseEvents Is Nothing) Then Set m_MouseEvents = Nothing
        
        If m_EditBoxSubclassed Then
            VBHacks.StopSubclassing m_EditBoxHwnd, Me
            m_EditBoxSubclassed = False
        End If
        
        If m_ParentHasBeenSubclassed Then
            If (m_ParentHWnd <> 0) Then VBHacks.StopSubclassing m_ParentHWnd, Me
            m_ParentHasBeenSubclassed = False
            m_ParentHWnd = 0
        End If
        
        DestroyEditBox = (DestroyWindow(m_EditBoxHwnd) <> 0)
        
        If DestroyEditBox Then
            UserControls.NotifyAPIWindowDestroyed
        Else
            PDDebug.LogAction "WARNING!  pdEditBoxW.DestroyEditBox failed.  System edit box #" & m_EditBoxHwnd & " *was not* destroyed!  DLL error was #" & Err.LastDllError
        End If
        
        m_EditBoxHwnd = 0
        
    Else
        DestroyEditBox = True
    End If
    
End Function

Friend Function GetPositionRect(ByRef dstPositionRect As winRect) As Boolean
    GetWindowRect m_EditBoxHwnd, dstPositionRect
    GetPositionRect = (m_EditBoxHwnd <> 0)
End Function

Friend Function GetWidth() As Long
    Dim tmpRect As winRect
    GetWindowRect m_EditBoxHwnd, tmpRect
    GetWidth = (tmpRect.x2 - tmpRect.x1)
End Function

'IMPORTANT NOTE: if the edit box has not yet been created, this function will return the control's *suggested* height, by default.
Friend Function GetHeight() As Long
    If (m_EditBoxHwnd <> 0) Then
        Dim tmpRect As winRect
        GetWindowRect m_EditBoxHwnd, tmpRect
        GetHeight = (tmpRect.y2 - tmpRect.y1)
    Else
        GetHeight = Me.SuggestedHeight
    End If
End Function

'Change the edit box position and size in one fell swoop.  Coordinates are RELATIVE TO THE PARENT WINDOW.
Friend Sub Move(ByVal newLeft As Long, ByVal newTop As Long, ByVal newWidth As Long, ByVal newHeight As Long)
    If (m_EditBoxHwnd <> 0) Then
        Dim moveSuccess As Boolean
        moveSuccess = (MoveWindow(m_EditBoxHwnd, newLeft, newTop, newWidth, newHeight, 1&) <> 0)
        If (Not moveSuccess) Then PDDebug.LogAction "WARNING: pdEditBoxW.Move failed with code " & Err.LastDllError
    End If
End Sub

'INSIDE A KEY EVENT (the only place this is valid), you can call this function to retrieve
' the unprocessed key code and scan code from the keyboardproc.
Friend Sub GetLastKeyAndScanCode(ByRef dstKeyCode As Long, ByRef dstScanCode As Long, Optional ByVal useKeyDownValues As Boolean = True)
    If useKeyDownValues Then
        dstKeyCode = m_lastKeyCodeDown
        dstScanCode = m_lastScanCodeDown
    Else
        dstKeyCode = m_lastKeyCodeUp
        dstScanCode = m_lastScanCodeUp
    End If
End Sub

'This function actually does two things: makes the edit box the foreground window, then assigns it keyboard focus.
' By design, all non-multiline edit boxes in PD select all their own text upon an initial click.
Friend Sub SetFocusToEditBox(Optional ByVal selectTextToo As Boolean = False)
    If (m_EditBoxHwnd <> 0) Then
        SetForegroundWindow m_EditBoxHwnd
        SetFocus m_EditBoxHwnd
        If selectTextToo Then Me.SelectAll
    End If
End Sub

'Given the current font size, return the suggested height of an edit box.  Parent controls can use this to auto-size
' themselves in preparation of actually creating the API window.
Friend Function SuggestedHeight() As Long
    SuggestedHeight = Fonts.GetDefaultStringHeight(m_FontSize) + 1
End Function

'Create the brush and font used to render the edit box background and text.  Normally, we only sync these values with the API window
' if absolutely necessary, but if you want to forcibly set them (at some performance penalty), set forceSync to TRUE.
Private Sub CreateEditBoxGDIObjects(Optional ByVal forceSync As Boolean = False)

    Dim RepaintRequired As Boolean
    
    If (m_EditBoxBrush <> UserControls.GetSharedGDIBrush(Me.BackColor)) Or forceSync Then
        If (m_EditBoxBrush <> 0) Then UserControls.ReleaseSharedGDIBrushByHandle m_EditBoxBrush
        m_EditBoxBrush = UserControls.GetSharedGDIBrush(Me.BackColor)
        RepaintRequired = True
    End If
    
    If (m_EditBoxFont <> UserControls.GetSharedGDIFont(m_FontSize)) Or RepaintRequired Or forceSync Then
        If (m_EditBoxFont <> 0) Then UserControls.ReleaseSharedGDIFontByHandle m_EditBoxFont
        m_EditBoxFont = UserControls.GetSharedGDIFont(m_FontSize)
        If (m_EditBoxHwnd <> 0) Then VBHacks.SendMsgW m_EditBoxHwnd, WM_SETFONT, m_EditBoxFont, IIf(m_Visible, 1, 0)
    End If

End Sub

'Call this function to release any created GDI edit boxes.  Note that this will be done automatically by the Create function, so you
' only need to call this dedicated function when the edit box is destroyed.
Private Sub ReleaseEditBoxGDIObjects()

    'Release any GDI objects associated with the edit box
    If (m_EditBoxBrush <> 0) Then
        UserControls.ReleaseSharedGDIBrushByHandle m_EditBoxBrush
        m_EditBoxBrush = 0
    End If
    
    If (m_EditBoxFont <> 0) Then
        UserControls.ReleaseSharedGDIFontByHandle m_EditBoxFont
        m_EditBoxFont = 0
    End If
    
End Sub

Private Sub m_MouseEvents_MouseEnter(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    RaiseEvent MouseEnter(Button, Shift, x, y)
    m_MouseEvents.SetCursor_System IDC_IBEAM
End Sub

Private Sub m_MouseEvents_MouseHover(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    RaiseEvent MouseHover(Button, Shift, x, y)
End Sub

Private Sub m_MouseEvents_MouseLeave(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    RaiseEvent MouseLeave(Button, Shift, x, y)
    m_MouseEvents.SetCursor_System IDC_DEFAULT
End Sub

'On first click, we select all text (but only if the user hasn't manually selected text with their initial mouse action)
Private Sub m_MouseEvents_MouseUpCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal clickEventAlsoFiring As Boolean, ByVal timeStamp As Long)
    'If m_JustGotFocus And Me.SelLength = 0 Then Me.SelectAll
    m_JustGotFocus = False
End Sub

Private Sub m_MouseEvents_MouseWheelHorizontal(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal scrollAmount As Double)
    RaiseEvent MouseWheelHorizontal(Button, Shift, x, y, scrollAmount)
End Sub

Private Sub m_MouseEvents_MouseWheelVertical(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal scrollAmount As Double)
    RaiseEvent MouseWheelVertical(Button, Shift, x, y, scrollAmount)
End Sub

Private Sub m_MouseEvents_MouseWheelZoom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal zoomAmount As Double)
    RaiseEvent MouseWheelZoom(Button, Shift, x, y, zoomAmount)
End Sub

Private Sub Class_Initialize()
    If PDMain.IsProgramRunning() Then Set m_MouseEvents = New pdInputMouse
    m_Enabled = True
    m_Visible = True
End Sub

Private Sub Class_Terminate()
    
    'Destroy the edit box, as necessary.  (Note that this also frees our internal subclassers)
    DestroyEditBox
    
    'Release any extra subclasser(s)
    If (Not m_MouseEvents Is Nothing) Then Set m_MouseEvents = Nothing
    
End Sub

Private Function HandleKeyDownUpMsgs(ByVal hWnd As Long, ByVal uiMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByRef msgEaten As Boolean, ByVal keyIsDown As Boolean) As Long
    
    'Inside keydown/up messages, we want to cache WM_CHAR, if one exists.  (If we don't do this, VB will eat unicode
    ' character codes inside its TranslateAccelerator handler.)
    Dim tmpMsg As TagMSG
    If (PeekMessage(tmpMsg, hWnd, WM_CHAR, WM_CHAR, PM_NOREMOVE) <> 0) Then m_CharCache = tmpMsg.wParam
    HandleKeyDownUpMsgs = 0
    
    'Cache key and scan code(s)
    If keyIsDown Then
        m_lastKeyCodeDown = wParam
        m_lastScanCodeDown = (lParam \ 65536) And &HFF&
    Else
        m_lastKeyCodeUp = wParam
        m_lastScanCodeUp = (lParam \ 65536) And &HFF&
    End If
    'Private m_lastScanCodeDown As Long, m_lastKeyCodeDown As Long
    'Private m_lastScanCodeUp As Long, m_lastKeyCodeUp As Long
    
    'Raise a corresponding keydown/up event; note that PD rarely uses these - instead, it preferentially uses
    ' KeyPress events.
    Dim retShiftConstants As ShiftConstants
    If IsVirtualKeyDown(VK_SHIFT) Then retShiftConstants = retShiftConstants Or vbShiftMask
    If IsVirtualKeyDown(VK_CONTROL) Then retShiftConstants = retShiftConstants Or vbCtrlMask
    If IsVirtualKeyDown(VK_ALT) Then retShiftConstants = retShiftConstants Or vbAltMask
    If keyIsDown Then
        RaiseEvent KeyDown(retShiftConstants, wParam, msgEaten)
    Else
        RaiseEvent KeyUp(retShiftConstants, wParam, msgEaten)
    End If
    
End Function

Private Function HandleEditBoxMsgs(ByVal hWnd As Long, ByVal uiMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByRef msgEaten As Boolean) As Long

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
    'Anyway, the moral of the story is that we have to do extra work to bypass the default message translator. Without this,
    ' IME entry methods (easily tested via the Windows on-screen keyboard and some non-Latin language) result in ???? chars,
    ' despite use of a Unicode window - and that ultimately defeats the whole point of a Unicode text box, no?
    
    'Manually pull key modifier states (shift, control, alt/menu) in advance; these are standard for all key events
    Dim retShiftConstants As ShiftConstants
    If IsVirtualKeyDown(VK_SHIFT) Then retShiftConstants = retShiftConstants Or vbShiftMask
    If IsVirtualKeyDown(VK_CONTROL) Then retShiftConstants = retShiftConstants Or vbCtrlMask
    If IsVirtualKeyDown(VK_ALT) Then retShiftConstants = retShiftConstants Or vbAltMask
    
    If (uiMsg = WM_CHAR) Then
        
        'On a single-line text box, pressing Enter will cause a ding.  (It's insane that this is the default behavior.)
        ' To prevent this, we capture return presses and forcibly terminate them.  If the user wants to do something with
        ' Enter keypresses on these controls, they will have already handled the key event in the hook proc.
        '
        'Note that we also catch the weird case of "wParam = 1"; this prevents a ding when Ctrl+A is used
        ' to select all text.  (For some reason, this causes a left-button keycode to be sent to the window?)
        If (Not m_Multiline) Then
            
            If ((wParam = VK_RETURN) Or (wParam = VK_TAB) Or (wParam = VK_ESCAPE) Or (wParam = 1)) Then
                HandleEditBoxMsgs = 0
                msgEaten = True
            End If
            
            'Notify our owner of this message.  (FYI, I realize this is all sorts of messed up - we're reporting a
            ' char code instead of a key code, contrary to the way the event is defined.  This mimics the way VB6 does it,
            ' and because we don't actually rely on virtual key IDs at present, it doesn't matter.  But keep this in mind
            ' if you decide to actually handle and respond to these events.)
            RaiseEvent KeyPress(retShiftConstants, wParam, msgEaten)
            
        Else
            
            'The "wParam = 1" case mentioned above still needs to be handled here, alas - dings happen in
            ' multiline edit boxes too!
            If (wParam = 1) Then
                HandleEditBoxMsgs = 0
                msgEaten = True
            End If
            
            RaiseEvent KeyPress(retShiftConstants, wParam, msgEaten)
            
        End If
        
        'If we previously cached a Unicode character in WM_KEYDOWN, cast it back to wParam now.
        If (Not msgEaten) Then
            
            'Some WM_CHAR messages may not originate from the keyboard; instead, PD may raise them manually from
            ' window messages that the default window proc may garble (e.g. WM_IME_CHAR, below).  If our character
            ' cache is empty, it means this is one of those messages; grab the current wParam and use it instead
            ' of our (blank) cached value.
            If (m_CharCache = 0) Then m_CharCache = wParam
            
            msgEaten = True
            HandleEditBoxMsgs = VBHacks.DefaultSubclassProc(hWnd, uiMsg, m_CharCache And &HFFFF&, lParam)
            m_CharCache = 0
            
        End If
        
    'Potential bug-fix added June 2017; manually handle IME-originating chars as well, just in case the def
    ' subclass proc fails to process them correctly.
    ElseIf (uiMsg = WM_IME_CHAR) Then
        HandleEditBoxMsgs = VBHacks.SendMsgW(hWnd, WM_CHAR, wParam, lParam)
        msgEaten = True
    
    'WM_UNICHAR messages are never sent by Windows.  However, third-party IMEs may send them.  Before allowing WM_UNICHAR
    ' messages to pass, Windows will first probe a window by sending the UNICODE_NOCHAR value.  If a window responds with 1
    ' (instead of 0, as DefWindowProc does), Windows will allow WM_UNICHAR messages to pass.
    ElseIf (uiMsg = WM_UNICHAR) Then
        
        'Probing for WM_UNICHAR support; respond with TRUE
        If (wParam = UNICODE_NOCHAR) Then
            msgEaten = True
            HandleEditBoxMsgs = 1
        
        'Passing an actual unicode character; forward this to our default handler
        Else
            m_CharCache = wParam
            VBHacks.SendMsgW hWnd, WM_CHAR, wParam, lParam
            msgEaten = True
            HandleEditBoxMsgs = 0
        End If
            
    'Manually dispatch WM_KEYDOWN messages.
    ElseIf (uiMsg = WM_KEYDOWN) Or (uiMsg = WM_KEYUP) Then
        HandleEditBoxMsgs = HandleKeyDownUpMsgs(hWnd, uiMsg, wParam, lParam, msgEaten, uiMsg = WM_KEYDOWN)
        
    'When the control receives focus, initialize a keyboard hook.  This prevents accelerators from working, but it is the
    ' only way to bypass VB's internal message translator, which will forcibly convert certain Unicode chars to ANSI.
    ElseIf (uiMsg = WM_SETFOCUS) Then
        
        'Forcibly disable PD's main accelerator control
        FormMain.HotkeyManager.DeactivateHook
        
        InstallKeyboardHook
        
        'Raise a matching notification
        If (Not m_HasFocus) Then
            m_HasFocus = True
            UserControls.PDControlReceivedFocus hWnd
            RaiseEvent GotFocusAPI
        End If
        
        'This value is important; on MOUSE_UP, we test it, and if the control just received focus, we select all text.
        m_JustGotFocus = True
        
    ElseIf (uiMsg = WM_KILLFOCUS) Then
        
        RemoveKeyboardHook
        
        'Re-enable PD's main accelerator control
        FormMain.HotkeyManager.ActivateHook
        
        'Raise a matching notification
        If m_HasFocus Then
            m_HasFocus = False
            UserControls.PDControlLostFocus hWnd
            RaiseEvent LostFocusAPI
        End If
            
    'Depending how our parent control is unloaded, the API edit box may receive a WM_DESTROY message before
    ' this class is terminated.  If that happens, we'll manually invoke our own internal DestroyWindow code.
    ElseIf (uiMsg = WM_DESTROY) Then
        
        RemoveKeyboardHook
        
        If DestroyEditBox Then
            HandleEditBoxMsgs = 0
            msgEaten = True
        End If
        
    'Failsafe window destruction check
    ElseIf (uiMsg = WM_NCDESTROY) Then
    
        RemoveKeyboardHook
        VBHacks.StopSubclassing m_ParentHWnd, Me
        m_ParentHasBeenSubclassed = False
        'Note that we deliberately *do not* eat this message, as VB's default wndproc needs to perform its own teardown
    
    End If
    
End Function

'Subclassed *parent window* messages are handled here (like WM_COMMAND)
Private Function HandleParentMsgs(ByVal hWnd As Long, ByVal uiMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByRef msgEaten As Boolean) As Long
    
    'The parent receives this message for all kinds of things; we subclass it to track when the edit box's contents have changed.
    ' (And when we don't handle the message, it is *very important* that we forward it correctly!
    If (uiMsg = WM_COMMAND) Then
        
        'Make sure the command is relative to *our* edit box, and not another one
        If (lParam = m_EditBoxHwnd) Then
            
            'Check for the EN_UPDATE flag; if present, raise the CHANGE event
            If ((wParam \ &H10000) = EN_UPDATE) Then
                RaiseEvent Change
                HandleParentMsgs = 0
                msgEaten = True
            End If
            
        End If
        
    'The parent receives this message prior to an ENABLED edit box being drawn.  The parent can use this to assign text and
    ' background colors to the edit box.
    ElseIf (uiMsg = WM_CTLCOLOREDIT) Then
        
        'Make sure the command is relative to *our* edit box, and not another one
        If (lParam = m_EditBoxHwnd) Then
            
            'wParam contains the relevant hDC
            
            'Assign colors directly, using the API
            SetTextColor wParam, Me.TextColor
            SetBkMode wParam, SBKM_OPAQUE
            SetBkColor wParam, Me.BackColor
            
            'Return the background brush we set
            HandleParentMsgs = m_EditBoxBrush
            msgEaten = True
            
        End If
        
    'The parent receives this message prior to a DISABLED edit box being drawn.  The parent can use this to assign text and
    ' background colors to the edit box.
    ElseIf (uiMsg = WM_CTLCOLORSTATIC) Then
    
        'Make sure the command is relative to *our* edit box, and not another one
        If lParam = m_EditBoxHwnd Then
        
            'Assign colors directly, using the API
            SetTextColor wParam, Me.TextColor
            SetBkMode wParam, SBKM_OPAQUE
            SetBkColor wParam, Me.BackColor
            
            'Return the background brush, per MSDN instructions
            HandleParentMsgs = m_EditBoxBrush
            msgEaten = True
            
        End If
        
    ElseIf (uiMsg = WM_NCDESTROY) Then
        VBHacks.StopSubclassing m_ParentHWnd, Me
        m_ParentHasBeenSubclassed = False
        'Note that we deliberately *do not* eat this message, as VB's default wndproc needs to perform its own teardown
    
    End If
        
End Function

Private Function ISubclass_WindowMsg(ByVal hWnd As Long, ByVal uiMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal dwRefData As Long) As Long
    
    Dim msgEaten As Boolean: msgEaten = False
    
    'Subclass handling is broken into two chunks: one for the edit box itself, and one for the parent window (which is also subclassed).
    If (hWnd = m_EditBoxHwnd) Then
        ISubclass_WindowMsg = HandleEditBoxMsgs(hWnd, uiMsg, wParam, lParam, msgEaten)
        
    ElseIf (hWnd = m_ParentHWnd) Then
        ISubclass_WindowMsg = HandleParentMsgs(hWnd, uiMsg, wParam, lParam, msgEaten)
    
    End If
    
    If (Not msgEaten) Then ISubclass_WindowMsg = VBHacks.DefaultSubclassProc(hWnd, uiMsg, wParam, lParam)
    
End Function

'On focus events, keyboard hooks are set or removed accordingly.  See the VBHacks module for details.
Private Sub InstallKeyboardHook()
    If PDMain.IsProgramRunning() Then
        If (m_HookID = 0) Then m_HookID = VBHacks.NotifyEditBoxHookNeeded(Me)
    End If
End Sub

Private Sub RemoveKeyboardHook()
    If (m_HookID <> 0) Then
        UnhookWindowsHookEx m_HookID
        m_HookID = 0
        VBHacks.NotifyEditBoxHookNotNeeded ObjPtr(Me)
    End If
End Sub

Private Function HandleKeyProc(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long, ByRef msgEaten As Boolean) As Long
    
    'Only respond to actual key presses (not just message peeks), and ignore key events if a menu is dropped
    If (nCode = 0) Then
    If (Not Menus.IsMainMenuActive()) Then
        
        'For the most part, we can safely ignore key presses, and instead let the default window handler handle them.
        ' The main outlier is arrow keys; VB intercepts and eat these (under the faulty assumption that they're
        ' meant for dialog navigation).  As such, we have to watch for incoming arrow keypresses, and when they occur,
        ' manually forward them to the edit box.
        
        'Bit 31 of lParam is 0 if a key is being pressed, and 1 if it is being released.  We check this in advance,
        ' so we don't double-raise key events.
        If (lParam >= 0) Then
        
            'Key up/down events are raised twice; once in a transitionary stage, and once again in a final stage.
            ' To prevent double-raising events, check the transitionary state before proceeding
            If ((lParam And 1) <> 0) And ((lParam And 3) = 1) Then
                
                'Manually pull key modifier states (shift, control, alt/menu)
                Dim retShiftConstants As ShiftConstants
                If IsVirtualKeyDown(VK_SHIFT) Then retShiftConstants = retShiftConstants Or vbShiftMask
                If IsVirtualKeyDown(VK_CONTROL) Then retShiftConstants = retShiftConstants Or vbCtrlMask
                If IsVirtualKeyDown(VK_ALT) Then retShiftConstants = retShiftConstants Or vbAltMask
                
                'Check for arrow keys and forward them manually.  (Also, eat these messages so no one else receives them.)
                If (wParam = VK_LEFT) Or (wParam = VK_UP) Or (wParam = VK_RIGHT) Or (wParam = VK_DOWN) Then
                    RaiseEvent KeyPress(retShiftConstants, wParam, msgEaten)
                    If (Not msgEaten) Then VBHacks.SendMsgW m_EditBoxHwnd, WM_KEYDOWN, wParam, (lParam And &H51111111)
                    msgEaten = True
                End If
                
                'Check for tab presses and raise them manually.  (Note that this step may need to
                ' be skipped in multiline edit boxes - for now, PD assumes TAB is always a navigation
                ' command, but perhaps we'll want to use a property to toggle this in the future?)
                If (wParam = VK_TAB) Then
                    RaiseEvent KeyPress(retShiftConstants, wParam, msgEaten)
                    msgEaten = True
                End If
            
            End If
        End If
    End If
    End If
    
End Function

'VBHacks will forward hooked messages here, while a hook is active
Friend Function EditBoxKeyboardProc(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
    'If we've shut down our internal hook handling code, ignore this event completely if.
    If (m_HookID <> 0) Then
        
        Dim msgEaten As Boolean: msgEaten = False
        
        'MSDN states that negative codes must be passed to the next hook, without processing
        ' (see http://msdn.microsoft.com/en-us/library/ms644984.aspx)
        If (nCode >= 0) Then
            EditBoxKeyboardProc = HandleKeyProc(nCode, wParam, lParam, msgEaten)
        End If
        
        'Per MSDN, return the value of CallNextHookEx, contingent on whether or not we handled the keypress internally.
        ' Note that if we do not manually handle a keypress, this behavior allows the default keyhandler to deal with
        ' the pressed keys (and raise its own corresponding WM_CHAR events, etc).
        If (Not msgEaten) Then
            EditBoxKeyboardProc = CallNextHookEx(0&, nCode, wParam, lParam)
        Else
            EditBoxKeyboardProc = 1&
        End If
        
    End If

End Function
