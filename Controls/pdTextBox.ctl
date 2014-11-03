VERSION 5.00
Begin VB.UserControl pdTextBox 
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
End
Attribute VB_Name = "pdTextBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Unicode Text Box control
'Copyright ©2013-2014 by Tanner Helland
'Created: 03/November/14
'Last updated: 03/November/14
'Last update: initial build
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
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************


Option Explicit

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

'Many APIs are required for this control
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExW" (ByVal dwExStyle As Long, ByVal lpClassName As Long, ByVal lpWindowName As Long, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByRef lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long

'Handle to the system edit box wrapped by this control
Private m_EditBoxHwnd As Long

'pdFont handles the creation and maintenance of the font used to render the text box.  It is also used to determine control width for
' single-line text boxes, as the control is auto-sized to fit the current font.
Private curFont As pdFont

'Rather than use an StdFont container (which requires VB to create redundant font objects), we track font properties manually,
' via dedicated properties.
Private m_FontSize As Single

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

'GetKeyboardState fills a [256] array with the state of all keyboard keys  Rather than constantly redimming an array for holding those
' return values, we simply declare one at a module level.
Private Declare Function GetKeyboardState Lib "user32" (ByRef pbKeyState As Byte) As Long
Private m_keyStateData(0 To 255) As Byte

Private Declare Function ToUnicode Lib "user32" (ByVal uVirtKey As Long, ByVal uScanCode As Long, lpKeyState As Byte, ByVal pwszBuff As Long, ByVal cchBuff As Long, ByVal wFlags As Long) As Long
Private Declare Function DispatchMessage Lib "user32" Alias "DispatchMessageW" (lpMsg As winMsg) As Long
Private Declare Function PeekMessage Lib "user32" Alias "PeekMessageW" (ByRef lpMsg As winMsg, ByVal hWnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long

Private Const WM_KEYDOWN As Long = &H100
Private Const WM_CHAR As Long = &H102
Private cSubclass As cSelfSubHookCallback

'Additional helpers for rendering themed and multiline tooltips
Private m_ToolTip As clsToolTip
Private m_ToolString As String

'hWnds aren't exposed by default
Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

'Container hWnd must be exposed for external tooltip handling
Public Property Get containerHwnd() As Long
    containerHwnd = UserControl.containerHwnd
End Property

'Font properties; only a subset are used, as PD handles most font settings automatically
Public Property Get FontSize() As Single
    FontSize = m_FontSize
End Property

Public Property Let FontSize(ByVal newSize As Single)
    If newSize <> m_FontSize Then
        m_FontSize = newSize
        refreshFont
    End If
End Property

'When the user control is hidden, we must hide the edit box window as well
' TODO!
Private Sub UserControl_Hide()
    
    
    
End Sub

Private Sub UserControl_Initialize()

    m_EditBoxHwnd = 0
    
    Set curFont = New pdFont
    m_FontSize = 10
    
    'Create an initial font object
    refreshFont
    
    'At run-time, initialize a subclasser
    If g_UserModeFix Then Set cSubclass = New cSelfSubHookCallback
    
End Sub

Private Sub UserControl_InitProperties()
    FontSize = 10
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    With PropBag
        FontSize = .ReadProperty("FontSize", 10)
    End With

End Sub

'When the user control is resized, the text box must be resized to match
' TODO: use a single helper function for calculating the edit box window rect.  We may want to draw our own border around the text box,
' for theming purposes, so we don't want multiple functions calculating their own window rect.
Private Sub UserControl_Resize()

    If m_EditBoxHwnd <> 0 Then
        MoveWindow m_EditBoxHwnd, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, 1
    End If

End Sub

'TODO: When the user control is shown, we must show the edit box window as well
Private Sub UserControl_Show()

    'If we have not yet created the edit box, do so now
    If m_EditBoxHwnd = 0 Then
    
        If Not createEditBox() Then Debug.Print "Edit box could not be created!"
    
    'The edit box has already been created, so we just need to show it
    Else
    
    End If
    
    'When the control is first made visible, remove the control's tooltip property and reassign it to the checkbox
    ' using a custom solution (which allows for linebreaks and theming).  Note that this has the ugly side-effect of
    ' permanently erasing the extender's tooltip, so FOR THIS CONTROL, TOOLTIPS MUST BE SET AT RUN-TIME!
    '
    'TODO!  Add helper functions for setting the tooltip to the created hWnd, instead of the VB control
    m_ToolString = Extender.ToolTipText

    If m_ToolString <> "" Then

        Set m_ToolTip = New clsToolTip
        With m_ToolTip

            .Create Me
            .MaxTipWidth = PD_MAX_TOOLTIP_WIDTH
            .AddTool Me, m_ToolString
            Extender.ToolTipText = ""

        End With

    End If

End Sub

'As the wrapped system edit box may need to be recreated when certain properties are changed, this function is used to
' automate the process of destroying an existing window (if any) and recreating it anew.
Private Function createEditBox() As Boolean

    'If the edit box already exists, kill it
    destroyEditBox
    
    'Figure out which flags to use, based on the control's properties
    Dim flagsWinStyle As Long, flagsWinStyleExtended As Long, flagsEditControl As Long
    flagsWinStyle = WS_VISIBLE Or WS_CHILD
    flagsWinStyleExtended = 0
    flagsEditControl = ES_AUTOHSCROLL 'Or ES_NOHIDESEL
    
    m_EditBoxHwnd = CreateWindowEx(flagsWinStyleExtended, ByVal StrPtr("EDIT"), ByVal StrPtr(""), flagsWinStyle Or flagsEditControl, _
        0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, UserControl.hWnd, 0, App.hInstance, ByVal 0&)
    
    'Assign a subclasser to enable IME support
    If g_UserModeFix Then
        If Not cSubclass Is Nothing Then
            cSubclass.ssc_Subclass m_EditBoxHwnd, 0, 1, Me, True, True, True
            cSubclass.ssc_AddMsg m_EditBoxHwnd, MSG_BEFORE, WM_KEYDOWN
        Else
            Debug.Print "subclasser could not be initialized for text box!"
        End If
    End If
        
    
    'Assign the default font to the edit box
    refreshFont True
    
    'Return TRUE if successful
    createEditBox = (m_EditBoxHwnd <> 0)

End Function

Private Function destroyEditBox() As Boolean

    If m_EditBoxHwnd <> 0 Then
        cSubclass.ssc_UnSubclass m_EditBoxHwnd
        DestroyWindow m_EditBoxHwnd
    End If
    
    destroyEditBox = True

End Function

Private Sub UserControl_Terminate()
    
    'Destroy the edit box, as necessary
    destroyEditBox
    
    'Release any extra subclasser(s)
    If Not cSubclass Is Nothing Then cSubclass.ssc_Terminate
    
End Sub

'When the font used for the edit box changes in some way, it can be recreated (refreshed) using this function.  Note that font
' creation is expensive, so it's worthwhile to avoid this step as much as possible.
Private Sub refreshFont(Optional ByVal forceRefresh As Boolean = False)
    
    Dim fontRefreshRequired As Boolean
    fontRefreshRequired = curFont.hasFontBeenCreated
    
    'Update each font parameter in turn.  If one (or more) requires a new font object, the font will be recreated as the final step.
    
    'Font face is always set automatically, to match the current program-wide font
    If (Len(g_InterfaceFont) > 0) And (StrComp(curFont.getFontFace, g_InterfaceFont, vbBinaryCompare) <> 0) Then
        fontRefreshRequired = True
        curFont.setFontFace g_InterfaceFont
    End If
    
    'In the future, I may switch to GDI+ for font rendering, as it supports floating-point font sizes.  In the meantime, we check
    ' parity using an Int() conversion, as GDI only supports integer font sizes.
    If Int(m_FontSize) <> Int(curFont.getFontSize) Then
        fontRefreshRequired = True
        curFont.setFontSize m_FontSize
    End If
        
    'Request a new font, if one or more settings have changed
    If fontRefreshRequired Or forceRefresh Then
        curFont.createFontObject
        
        'Whenever the font is recreated, we need to reassign it to the text box.  This is done via the WM_SETFONT message.
        If m_EditBoxHwnd <> 0 Then SendMessage m_EditBoxHwnd, WM_SETFONT, curFont.getFontHandle, IIf(UserControl.Extender.Visible, 1, 0)
        
    End If
    
    'Also, the back buffer needs to be rebuilt to reflect the new font metrics
    'updateControlSize

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    'Store all associated properties
    With PropBag
        .WriteProperty "FontSize", m_FontSize, 10
    End With
    
End Sub

'External functions can call this to request a redraw.  This is helpful for live-updating theme settings, as in the Preferences dialog.
Public Sub updateAgainstCurrentTheme()
    
    If g_UserModeFix Then
        
        'Update the current font, as necessary
        refreshFont
        
        'Force an immediate repaint
        'updateControlSize
                
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
    
    'Now comes a really messy bit of VB-specific garbage.
    
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
    
    'Anyway, the moral of the story is that we have to do a shitload of extra work to bypass the default message translator.
    ' Without this, IME entry methods (easily tested via the Windows on-screen keyboard and a language like Kazakh) result in
    ' ???? chars, despite use of a Unicode window - and that ultimately defeats the whole point of a Unicode text box, no?
    
    'Manually dispatch WM_KEYDOWN messages
    If uMsg = WM_KEYDOWN Then
        
        'Start by retrieving a full copy of the WM_KEYDOWN message contents, and also removing the WM_KEYDOWN message from the stack.
        Dim tmpMsg As winMsg
        If PeekMessage(tmpMsg, m_EditBoxHwnd, 0, 0, 1) <> 0 Then
            
            'Note that PeekMessage, above, should never fail since we are literally calling it from within the wndProc - but better safe
            ' than sorry, right??
            
            'Next, we need to retrieve the status of all 256 keyboard keys.  This is important for non-Latin keyboards, which can
            ' produce Unicode characters in a variety of ways.  (For example, by holding down multiple keys at once.)
            If GetKeyboardState(m_keyStateData(0)) <> 0 Then
            
                'Next, we need to prepare a string buffer to receive the Unicode translation of the current virtual key.
                ' This is tricky because ToUnicode/Ex do not specify a max buffer size they may write.  Michael Kaplan's
                ' definitive article series on this topic (dead link on MSDN; I found it here: http://www.siao2.com/2006/03/23/558674.aspx)
                ' uses a 10-char buffer.  That should be sufficient for our purposes as well.
                Dim tmpString As String
                tmpString = String$(10, vbNullChar)
                
                'Perform a Unicode translation using the pressed virtual key (wParam) and the buffer of all current key states
                Dim unicodeResult As Long
                unicodeResult = ToUnicode(wParam, 0, m_keyStateData(0), StrPtr(tmpString), Len(tmpString), 0)
                
                'ToUnicode has four possible return values:
                ' -1: the char is an accent or diacritic.  If possible, it has been translated to a standalone spacing version
                '     (always a UTF-16 entry point), and placed in the output buffer.  For our purposes, we'll just retrieve the
                '     UTF-16 entry point and call it good.
                ' 0: function failed
                ' 1: success; a single Unicode character was written to the buffer
                ' 2+: success; multiple Unicode characters were written to the buffer, typically when a matching ligature was
                '    not found for a relevant multi-glyph input.  This is a valid return, and all specified characters should
                '    be sent to the text box, if possible.
                If unicodeResult = -1 Then unicodeResult = 2
                                
                'IMPORTANT!  The string buffer can contain more values than the return value specified, so it's important to
                ' shrink the buffer using the *return value*, and *not the buffer's actual contents*.
                Select Case unicodeResult
                
                    'Dead character, meaning a single UTF-16 entry point.  This case was forcibly forwarded to type 2,
                    ' above, so this case will never raise - I just include it here for reference.
                    Case -1
                    
                    'Failure; no Unicode result
                    Case 0
                    
                    '1 to 4 chars
                    Case 1 To 4
                    
                        'Retrieve the relevant portion of the string
                        tmpString = Left$(tmpString, unicodeResult)
                        
                        'This is the problematic part.  For single character strings, AscW should work just fine.  However, I'm not sure
                        ' what to do for longer buffers.
                        ' TODO: investigate http://www.cyberactivex.com/UnicodeTutorialVb.htm#SurrogatePairs as a possible solution
                        tmpMsg.wParam = CLng(AscW(tmpString))
                        
                        'Convert the message type to WM_CHAR
                        tmpMsg.sysMsg = WM_CHAR
                        
                        'Dispatch the message
                        DispatchMessage tmpMsg
                        
                        'Note that the message was handled successfully
                        bHandled = True
                        lReturn = 0
                    
                    Case Else
                        Debug.Print "Excessively large Unicode buffer value returned: " & unicodeResult
                    
                End Select
                
            End If
            
        Else
            Debug.Print "peek message failed"
        End If
        
    End If



' *************************************************************
' C A U T I O N   C A U T I O N   C A U T I O N   C A U T I O N
' -------------------------------------------------------------
' DO NOT ADD ANY OTHER CODE BELOW THE "END SUB" STATEMENT BELOW
'   add this warning banner to the last routine in your class
' *************************************************************
End Sub


