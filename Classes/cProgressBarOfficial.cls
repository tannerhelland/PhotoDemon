VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cProgressBarOfficial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Note: this file has been modified for use within PhotoDemon.

'To the best of my knowledge, this class is ©2011 LaVolpe, but released into the public domain.

'You may download the original version of this code from the following link (good as of September '13):
' http://www.vbforums.com/showthread.php?462116-Marquee-Progressbar

'This class differs from the other progress bar class because it creates an actual instance of a Windows progress
' bar, instead of simply opening the system's progress bar theme but doing all painting itself.  This is preferable
' because you can support things like marquee (endless) style, and the "lighting" animations present on Vista+.
' The plan is to eventually switch the main form's progress bar to this class, but it will take a bit of work to
' rewrite the message text rendering.

Option Explicit

' Must call CreateProgressBar before any other properties/functions can be called

Private Declare Sub InitCommonControls Lib "comctl32" ()
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByRef lpParam As Any) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, ByRef lpPoint As POINTAPI) As Long
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, ByRef lprcUpdate As Any, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long

' http://msdn.microsoft.com/en-us/library/ff485990(v=VS.85).aspx
' ^^ Documented progress bar messages
Private Const PROGRESS_CLASS As String = "msctls_progress32"
Private Const WM_USER As Long = &H400
Private Const CCM_FIRST As Long = &H2000
Private Const CCM_SETBKCOLOR As Long = CCM_FIRST + 1
Private Const PBM_DELTAPOS As Long = WM_USER + 3
Private Const PBM_GETPOS As Long = WM_USER + 8
Private Const PBM_GETRANGE As Long = WM_USER + 7
Private Const PBM_SETBARCOLOR As Long = WM_USER + 9
Private Const PBM_SETBKCOLOR As Long = CCM_SETBKCOLOR
Private Const PBM_SETPOS As Long = WM_USER + 2
Private Const PBM_SETRANGE As Long = WM_USER + 1
Private Const PBM_SETRANGE32 As Long = WM_USER + 6
Private Const PBS_VERTICAL As Long = &H4
Private Const PBM_SETMARQUEE = WM_USER + 10
Private Const PBM_SETSTATE As Long = WM_USER + 16   ' vista+ only - starts/pauses marquee?
Private Const PBM_GETSTATE As Long = WM_USER + 17   ' vista+ only
Private Const PBM_GETBARCOLOR As Long = WM_USER + 15
Private Const PBST_NORMAL As Long = 1               ' control must be themed; vista+ only
Private Const PBST_PAUSED As Long = 3               ' control must be themed; vista+ only
Private Const PBM_SETSTEP As Long = WM_USER + 4
Private Const PBM_GETSTEP As Long = WM_USER + 13
Private Const PBM_STEPIT As Long = WM_USER + 5
' http://msdn.microsoft.com/en-us/library/bb760820%28v=VS.85%29.aspx
' ^^ documented progressbar window styles
Private Const PBS_SMOOTH As Long = &H1
Private Const PBS_MARQUEE As Long = &H8             ' control must be themed & v6 of common controls installed
Private Const PBS_SMOOTHREVERSE As Long = &H10      ' vista+ only
Private Const WS_CHILD As Long = &H40000000
Private Const WS_VISIBLE As Long = &H10000000
Private Const GWL_STYLE As Long = -16
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private m_Hwnd As Long
Private m_Speed As Long         ' marquee speed
Private m_ForeColor As Long     ' cached no way of getting value in WinXP and lower
Private m_BackColor As Long     ' cached no way of getting value in WinXP and lower


Public Function CreateProgressBar(ByVal ContainerHwnd As Long, _
                        ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, _
                        ByVal Horizontal As Boolean, ByVal Smooth As Boolean, SmoothReverse As Boolean, ByVal Visible As Boolean) As Boolean
                        
    ' creates an API progress bar
    Dim lStyle As Long
    If m_Hwnd Then DestroyWindow m_Hwnd: m_Hwnd = 0&
    
    If Visible = True Then lStyle = WS_VISIBLE
    If Horizontal = False Then lStyle = lStyle Or PBS_VERTICAL
    If Smooth Then lStyle = lStyle Or PBS_SMOOTH
    If SmoothReverse Then lStyle = lStyle Or PBS_SMOOTHREVERSE  ' Vista+ only else ignored
    
    m_Hwnd = CreateWindowEx(0&, PROGRESS_CLASS, vbNullString, WS_CHILD Or lStyle, X, Y, Width, Height, ContainerHwnd, 0&, App.hInstance, ByVal 0&)
    If m_Hwnd = 0& Then
        InitCommonControls
        m_Hwnd = CreateWindowEx(0&, PROGRESS_CLASS, vbNullString, WS_CHILD Or lStyle, X, Y, Width, Height, ContainerHwnd, 0&, App.hInstance, ByVal 0&)
    End If
    If m_Hwnd Then
        m_ForeColor = &HFF000000 ' defaults
        m_BackColor = m_ForeColor
        m_Speed = 30
        CreateProgressBar = True
    End If
    
End Function

Public Property Let Horizontal(isHorizontal As Boolean)
    ' sets orientation at horizontal or vertical
    If isHorizontal <> Me.Horizontal Then
        pvCopyBarProperties isHorizontal, Me.SmoothDisplay, Me.SmoothReverse, True
    End If
End Property
Public Property Get Horizontal() As Boolean
    Horizontal = ((GetWindowLong(m_Hwnd, GWL_STYLE) And PBS_VERTICAL) = 0&)
End Property

Public Property Let SmoothDisplay(isSmooth As Boolean)
    ' sets display smooth or segmented; does not apply if control is themed
    If isSmooth <> Me.SmoothDisplay Then
        pvCopyBarProperties Me.Horizontal, isSmooth, Me.SmoothReverse, False
    End If
End Property
Public Property Get SmoothDisplay() As Boolean
    SmoothDisplay = ((GetWindowLong(m_Hwnd, GWL_STYLE) And PBS_SMOOTH) = PBS_SMOOTH)
End Property

Public Property Let SmoothReverse(isSmooth As Boolean) ' Vista+ only
    ' determines how bar is displayed when moving from higher to lower values: smooth or jumpy
    If isSmooth <> Me.SmoothReverse Then
        pvCopyBarProperties Me.Horizontal, Me.SmoothDisplay, isSmooth, False
    End If
End Property
Public Property Get SmoothReverse() As Boolean
    SmoothReverse = ((GetWindowLong(m_Hwnd, GWL_STYLE) And PBS_SMOOTHREVERSE) = PBS_SMOOTHREVERSE)
End Property


Public Property Let BarColor(newColor As Long) ' does not apply if control is themed
    If newColor = -1& Then m_ForeColor = &HFF000000 Else m_ForeColor = newColor And &HFFFFFF
    SendMessage m_Hwnd, PBM_SETBARCOLOR, 0&, ByVal m_ForeColor
End Property
Public Property Get BarColor() As Long
    If m_ForeColor = &HFF000000 Then BarColor = -1& Else BarColor = m_ForeColor
    ' note that  PBM_GETBARCOLOR is Vista+ only
End Property

Public Property Let BackColor(newColor As Long) ' does not apply if control is themed
    If newColor = -1& Then m_BackColor = &HFF000000 Else m_BackColor = newColor And &HFFFFFF
    SendMessage m_Hwnd, PBM_SETBKCOLOR, 0&, ByVal m_BackColor
End Property
Public Property Get BackColor() As Long
    If m_BackColor = &HFF000000 Then BackColor = -1& Else BackColor = m_BackColor
    ' note that PBM_GETBKCOLOR is Vista+ only
End Property

Public Property Let Visible(isVisible As Boolean) ' sets visibility
    ShowWindow m_Hwnd, Abs(isVisible)
End Property
Public Property Get Visible() As Boolean
    Visible = ((GetWindowLong(m_Hwnd, GWL_STYLE) And WS_VISIBLE) = WS_VISIBLE)
End Property

Public Property Get Left() As Long          ' returns/sets Left position in pixels
    Dim wRect As RECT, tPt As POINTAPI
    GetWindowRect m_Hwnd, wRect
    tPt.X = wRect.Left
    ScreenToClient GetParent(m_Hwnd), tPt
    Left = tPt.X
End Property
Public Property Let Left(newValue As Long)
    Dim wRect As RECT, tPt As POINTAPI
    GetWindowRect m_Hwnd, wRect
    tPt.X = wRect.Left: tPt.Y = wRect.Top
    ScreenToClient GetParent(m_Hwnd), tPt
    MoveWindow m_Hwnd, newValue, tPt.Y, wRect.Right - wRect.Left, wRect.Bottom - wRect.Top, True
End Property

Public Property Get Top() As Long           ' returns/sets Top position in pixels
    Dim wRect As RECT, tPt As POINTAPI
    GetWindowRect m_Hwnd, wRect
    tPt.Y = wRect.Top
    ScreenToClient GetParent(m_Hwnd), tPt
    Top = tPt.Y
End Property
Public Property Let Top(newValue As Long)
    Dim wRect As RECT, tPt As POINTAPI
    GetWindowRect m_Hwnd, wRect
    tPt.X = wRect.Left: tPt.Y = wRect.Top
    ScreenToClient GetParent(m_Hwnd), tPt
    MoveWindow m_Hwnd, tPt.X, newValue, wRect.Right - wRect.Left, wRect.Bottom - wRect.Top, True
End Property

Public Property Get Height() As Long        ' returns/sets Height in pixels
    Dim wRect As RECT
    GetWindowRect m_Hwnd, wRect
    Height = wRect.Bottom - wRect.Top
End Property
Public Property Let Height(newValue As Long)
    Dim wRect As RECT
    GetWindowRect m_Hwnd, wRect
    MoveWindow m_Hwnd, wRect.Left, wRect.Top, wRect.Right - wRect.Left, newValue, True
End Property

Public Property Get Width() As Long         ' returns/sets Width in pixels
    Dim wRect As RECT
    GetWindowRect m_Hwnd, wRect
    Width = wRect.Right - wRect.Left
End Property
Public Property Let Width(newValue As Long)
    Dim wRect As RECT
    GetWindowRect m_Hwnd, wRect
    MoveWindow m_Hwnd, wRect.Left, wRect.Top, newValue, wRect.Bottom - wRect.Top, True
End Property

Public Sub Move(ByVal Left As Long, ByVal Top As Long, Optional ByVal Width As Long, Optional ByVal Height As Long)
    ' passed parameters must be in pixels
    If Width = 0& Then Width = Me.Width
    If Height = 0& Then Height = Me.Height
    MoveWindow m_Hwnd, Left, Top, Width, Height, True

End Sub

Public Property Let Marquee(isMarquee As Boolean)   ' toggles whether marquee mode is on/off
    ' when turning off, the Value is reset to 0
    If Me.Marquee = False Then
        If isMarquee = True Then
            SetWindowLong m_Hwnd, GWL_STYLE, GetWindowLong(m_Hwnd, GWL_STYLE) Or PBS_MARQUEE
            SendMessage m_Hwnd, PBM_SETMARQUEE, Abs(isMarquee), ByVal m_Speed
        End If
    ElseIf isMarquee = False Then
        SendMessage m_Hwnd, PBM_SETMARQUEE, Abs(isMarquee), ByVal 0&
        SetWindowLong m_Hwnd, GWL_STYLE, (GetWindowLong(m_Hwnd, GWL_STYLE) And Not PBS_MARQUEE)
    End If
    
End Property
Public Property Get Marquee() As Boolean
    Marquee = ((GetWindowLong(m_Hwnd, GWL_STYLE) And PBS_MARQUEE) = PBS_MARQUEE)
End Property

Public Property Let MarqueeScrollRate(newValue As Long) ' pauses or resets/restarts marquee
    If newValue <> m_Speed And newValue >= 0& Then
        m_Speed = newValue
        If Me.Marquee Then SendMessage m_Hwnd, PBM_SETMARQUEE, Abs(m_Speed > 0&), ByVal m_Speed
    End If
End Property
Public Property Get MarqueeScrollRate() As Long
    MarqueeScrollRate = m_Speed
End Property

Public Property Let MarqueePaused(isPaused As Boolean) ' Vista+ only
    SendMessage m_Hwnd, PBM_SETSTATE, Abs(isPaused) * 2& + 1&, ByVal 0&
End Property
Public Property Get MarqueePaused() As Boolean
    MarqueePaused = (SendMessage(m_Hwnd, PBM_GETSTATE, 0&, ByVal 0&) = PBST_PAUSED)
End Property

Public Property Let Min(newValue As Long)   ' sets minimum bar value
    SendMessage m_Hwnd, PBM_SETRANGE32, newValue, ByVal Me.Max
End Property
Public Property Get Min() As Long
    Min = SendMessage(m_Hwnd, PBM_GETRANGE, 1&, ByVal 0&)
End Property

Public Property Let Max(newValue As Long)   ' sets maximum bar value
    SendMessage m_Hwnd, PBM_SETRANGE32, Me.Min, ByVal newValue
End Property
Public Property Get Max() As Long
    Max = SendMessage(m_Hwnd, PBM_GETRANGE, 0&, ByVal 0&)
End Property

Public Property Let Value(newValue As Long) ' returns/sets bar value
    SendMessage m_Hwnd, PBM_SETPOS, newValue, ByVal 0&
End Property
Public Property Get Value() As Long
    Value = SendMessage(m_Hwnd, PBM_GETPOS, 0&, ByVal 0&)
End Property

Public Sub Refresh()                        ' redraws the control
    RedrawWindow m_Hwnd, ByVal 0&, 0&, &H101&
End Sub

Public Property Get hWnd() As Long          ' returns window handle
    hWnd = m_Hwnd
End Property

Public Property Let StepValue(newValue As Long) ' sets/returns the current step value
    SendMessage m_Hwnd, PBM_SETSTEP, newValue, ByVal 0&
End Property
Public Property Get StepValue() As Long
    StepValue = SendMessage(m_Hwnd, PBM_GETSTEP, 0&, ByVal 0&)
    If StepValue = 0& Then StepValue = 10& ' default set by windows
End Property

Public Function StepIt(Optional ByVal IncrementBy As Long) As Long ' increments value by StepValue
    ' resets Value to StepValue if current Value + StepValue > Max
    If IncrementBy = 0& Then
        StepIt = SendMessage(m_Hwnd, PBM_STEPIT, 0&, ByVal 0&)
    Else
        StepIt = SendMessage(m_Hwnd, PBM_DELTAPOS, IncrementBy, ByVal 0&)
    End If
End Function

Private Sub pvCopyBarProperties(Horizontal As Boolean, ByVal Smooth As Boolean, ByVal SmoothReverse As Boolean, ToggleOrientation As Boolean)

    ' changing smooth/standard display and vertical/horizontal orientation requires
    ' creating a new progress bar. Copies attributes from previous bar & applies to new bar

    Dim wRect As RECT, tPt As POINTAPI
    Dim oldHwnd As Long, newHwnd As Long, lProp As Long
    Dim lColor1 As Long, lColor2 As Long
    
    GetWindowRect m_Hwnd, wRect
    tPt.X = wRect.Left: tPt.Y = wRect.Top
    ScreenToClient GetParent(m_Hwnd), tPt
    lColor1 = m_ForeColor: lColor2 = m_BackColor
    lProp = m_Speed
    oldHwnd = m_Hwnd: m_Hwnd = 0&
    If ToggleOrientation Then
        CreateProgressBar GetParent(oldHwnd), tPt.X, tPt.Y, wRect.Bottom - wRect.Top, wRect.Right - wRect.Left, Horizontal, Smooth, SmoothReverse, False
    Else
        CreateProgressBar GetParent(oldHwnd), tPt.X, tPt.Y, wRect.Right - wRect.Left, wRect.Bottom - wRect.Top, Horizontal, Smooth, SmoothReverse, False
    End If
    If m_Hwnd Then
        newHwnd = m_Hwnd
        m_Hwnd = oldHwnd: SendMessage newHwnd, PBM_SETRANGE32, Me.Min, ByVal Me.Max
        lProp = Me.Value: m_Hwnd = newHwnd: Me.Value = lProp
        m_Hwnd = oldHwnd: lProp = Me.StepValue: m_Hwnd = newHwnd: Me.StepValue = lProp
        m_Hwnd = oldHwnd: lProp = Me.Marquee: m_Hwnd = newHwnd: Me.Marquee = lProp
        If lColor1 = &HFF000000 Then Me.BarColor = -1& Else Me.BarColor = lColor1
        If lColor2 = &HFF000000 Then Me.BackColor = -1& Else Me.BackColor = lColor2
        m_Hwnd = oldHwnd
        If Me.Marquee Then
            m_Hwnd = newHwnd: Me.Marquee = True: Me.MarqueeScrollRate = m_Speed
            m_Hwnd = oldHwnd: lProp = Me.MarqueePaused: m_Hwnd = newHwnd: Me.MarqueePaused = lProp
        End If
        m_Hwnd = oldHwnd: lProp = Me.Visible: m_Hwnd = newHwnd: Me.Visible = lProp
        DestroyWindow oldHwnd
    Else
        m_Hwnd = oldHwnd
    End If

End Sub


