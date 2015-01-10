VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsToolTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Note: this file has been modified for use within PhotoDemon, including fixing of several critical bugs.

'To my knowledge, this class is Copyright 2010 Dipak Auddy, with modifications by Roger Gilchrist

'You may download the original version of this code from the following link (good as of March '13):
' http://www.vb-helper.com/howto_multi_line_tooltip.html

'Many thanks also to Rod Stephens for his excellent VB-helper website.


Option Explicit
'cslToolTip
'By Dipak Auddy
' Mail : auddy@gawab.com
'MODIFICATIONS
'by Roger Gilchrist
' Mail : rojagilkrist@hotmail.com
'Released with Dipak's permission
'
' Defaults:
' DelayInitial   = 500  (1/2 sec)
' DelayAutoPopup = 5000  (5 secs)
' DelayReshow    = 100 (1/10 sec)
' MaxTipWidth    = 0
' all Margins    = 0

'MODIFICATIONS
'moved all the support Types and declates into the class
'makes it much more portable
Private Type RECT
    Left                            As Long
    Top                             As Long
    Right                           As Long
    Bottom                          As Long
End Type

Private Const WM_USER             As Long = &H400
Private Const TOOLTIPS_CLASS      As String = "tooltips_class32"
Private Const TTS_ALWAYSTIP       As Long = &H1
Private Const TTM_SETTITLE As Long = (WM_USER + 32)

#Const WIN32_IE = &H400
Private Type TOOLINFO
    cbSize                          As Long
    uFlags                          As TT_Flags
    hWnd                            As Long
    uId                             As Long
    RECT                            As RECT
    hInst                           As Long
    lpszText                        As String
#If (WIN32_IE >= &H300) Then
    lParam                          As Long
#End If
End Type

Public Enum TT_Flags
    TTF_IDISHWND = &H1
    TTF_CENTERTIP = &H2
    TTF_RTLREADING = &H4
    TTF_SUBCLASS = &H10
#If (WIN32_IE >= &H300) Then
    TTF_TRACK = &H20
    TTF_ABSOLUTE = &H80
    TTF_TRANSPARENT = &H100
    TTF_DI_SETITEM = &H8000&
#End If
End Enum

#If False Then 'Trick preserves Case of Enums when typing in IDE
    Private TTF_IDISHWND, TTF_CENTERTIP, TTF_RTLREADING, TTF_SUBCLASS, TTF_TRACK, TTF_ABSOLUTE, TTF_TRANSPARENT, TTF_DI_SETITEM
#End If

Public Enum TT_DelayTime
    TTDT_AUTOMATIC = 0
    TTDT_RESHOW = 1
    TTDT_AUTOPOP = 2
    TTDT_INITIAL = 3
End Enum

#If False Then 'Trick preserves Case of Enums when typing in IDE
Private TTDT_AUTOMATIC, TTDT_RESHOW, TTDT_AUTOPOP, TTDT_INITIAL
#End If

Public Enum ttDelayTimeConstants
    ttDelayDefault = TTDT_AUTOMATIC '= 0
    ttDelayInitial = TTDT_INITIAL '= 3
    ttDelayShow = TTDT_AUTOPOP '= 2
    ttDelayReshow = TTDT_RESHOW '= 1
    ttDelayMask = 3
End Enum

#If False Then 'Trick preserves Case of Enums when typing in IDE
Private ttDelayDefault, ttDelayInitial, ttDelayShow, ttDelayReshow, ttDelayMask
#End If

Public Enum ttMarginConstants
    ttMarginLeft = 0
    ttMarginTop = 1
    ttMarginRight = 2
    ttMarginBottom = 3
End Enum

#If False Then 'Trick preserves Case of Enums when typing in IDE
Private ttMarginLeft, ttMarginTop, ttMarginRight, ttMarginBottom
#End If

Public Enum TT_Msgs
    TTM_ACTIVATE = (WM_USER + 1)
    TTM_SETDELAYTIME = (WM_USER + 3)
    TTM_RELAYEVENT = (WM_USER + 7)
    TTM_GETTOOLCOUNT = (WM_USER + 13)
    TTM_WINDOWFROMPOINT = (WM_USER + 16)
#If UNICODE Then
    TTM_ADDTOOL = (WM_USER + 50)
    TTM_DELTOOL = (WM_USER + 51)
    TTM_NEWTOOLRECT = (WM_USER + 52)
    TTM_GETTOOLINFO = (WM_USER + 53)
    TTM_SETTOOLINFO = (WM_USER + 54)
    TTM_HITTEST = (WM_USER + 55)
    TTM_GETTEXT = (WM_USER + 56)
    TTM_UPDATETIPTEXT = (WM_USER + 57)
    TTM_ENUMTOOLS = (WM_USER + 58)
    TTM_GETCURRENTTOOL = (WM_USER + 59)
#Else
    TTM_ADDTOOL = (WM_USER + 4)
    TTM_DELTOOL = (WM_USER + 5)
    TTM_NEWTOOLRECT = (WM_USER + 6)
    TTM_GETTOOLINFO = (WM_USER + 8)
    TTM_SETTOOLINFO = (WM_USER + 9)
    TTM_HITTEST = (WM_USER + 10)
    TTM_GETTEXT = (WM_USER + 11)
    TTM_UPDATETIPTEXT = (WM_USER + 12)
    TTM_ENUMTOOLS = (WM_USER + 14)
    TTM_GETCURRENTTOOL = (WM_USER + 15)
#End If
#If (WIN32_IE >= &H300) Then
    TTM_TRACKACTIVATE = (WM_USER + 17)
    TTM_TRACKPOSITION = (WM_USER + 18)
    TTM_SETTIPBKCOLOR = (WM_USER + 19)
    TTM_SETTIPTEXTCOLOR = (WM_USER + 20)
    TTM_GETDELAYTIME = (WM_USER + 21)
    TTM_GETTIPBKCOLOR = (WM_USER + 22)
    TTM_GETTIPTEXTCOLOR = (WM_USER + 23)
    TTM_SETMAXTIPWIDTH = (WM_USER + 24)
    TTM_GETMAXTIPWIDTH = (WM_USER + 25)
    TTM_SETMARGIN = (WM_USER + 26)
    TTM_GETMARGIN = (WM_USER + 27)
    TTM_POP = (WM_USER + 28)
#End If
#If (WIN32_IE >= &H400) Then
    TTM_UPDATE = (WM_USER + 29)
#End If
End Enum

#If False Then 'Trick preserves Case of Enums when typing in IDE
    Private TTM_ACTIVATE, TTM_SETDELAYTIME, TTM_RELAYEVENT, TTM_GETTOOLCOUNT, TTM_WINDOWFROMPOINT
    Private TTM_ADDTOOL, TTM_DELTOOL, TTM_NEWTOOLRECT, TTM_GETTOOLINFO, TTM_SETTOOLINFO, TTM_HITTEST
    Private TTM_GETTEXT, TTM_UPDATETIPTEXT, TTM_ENUMTOOLS, TTM_GETCURRENTTOOL, TTM_TRACKACTIVATE
    Private TTM_TRACKPOSITION, TTM_SETTIPBKCOLOR, TTM_SETTIPTEXTCOLOR, TTM_GETDELAYTIME, TTM_GETTIPBKCOLOR
    Private TTM_GETTIPTEXTCOLOR, TTM_SETMAXTIPWIDTH, TTM_GETMAXTIPWIDTH, TTM_SETMARGIN, TTM_GETMARGIN
    Private TTM_POP, TTM_UPDATE
#End If

Public Enum TT_Notifications
    TTN_FIRST = -520&
    TTN_LAST = -549&
#If UNICODE Then
    TTN_NEEDTEXT = (TTN_FIRST - 10)
#Else
    TTN_NEEDTEXT = (TTN_FIRST - 0)
#End If
    TTN_SHOW = (TTN_FIRST - 1)
    TTN_POP = (TTN_FIRST - 2)
End Enum

#If False Then 'Trick preserves Case of Enums when typing in IDE
Private TTN_FIRST, TTN_LAST, TTN_NEEDTEXT, TTN_SHOW, TTN_POP
#End If
'
' Exported by Comctl32.dll >= v4.00.950
' Ensures that the common control dynamic
' link library (DLL) is loaded.
'
' NOTE: API replaced by InitCommonControlsEx
Private mnlgHwndTT                As Long
Private mnlgMaxTip                As Long
Private m_TTHeader                As String
Private m_ShowHeader              As Boolean
Private Declare Function SendMessageT Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)

'Added by Tanner: due to complexities with always-on-top windows, we must also make tooltips always-on-top
Private Const WS_EX_TOPMOST As Long = &H8&

Public Function AddTool(ByRef ctrl As Object, Optional ByVal strText As String) As Boolean

    Dim ti As TOOLINFO
    
    'The IDE may throw random hissy fits about the Container.hWnd workaround property below.  To avoid this,
    ' simply ignore errors in this function.
    On Error Resume Next
    
    'MODIFICATION
    'if a control has no hWnd to use for tooltip
    If Not fIsWindow(ctrl) Then
        ''MsgBox ctrl.Name & " cannot have a tool tip", vbCritical
        Exit Function
    End If
    
    If (mnlgHwndTT = 0) Then
        Exit Function
    End If
  
    If (fGetToolInfo(ctrl.hWnd, ti) = False) Then
    
        With ti
            .cbSize = Len(ti)
            '
            ' TTF_IDISHWND must be set to tell the tooltip
            ' control to retrieve the control's rect from
            ' it's hWnd specified in uId.
            '
            .uFlags = TTF_SUBCLASS Or TTF_IDISHWND
            If TypeOf ctrl Is Form Then
                .hWnd = ctrl.hWnd
            Else
            
                If TypeOf ctrl Is smartOptionButton Then
                    .hWnd = ctrl.containerHwnd
                ElseIf TypeOf ctrl Is smartCheckBox Then
                    .hWnd = ctrl.containerHwnd
                Else
                    .hWnd = ctrl.Container.hWnd
                End If
                
            End If
            .uId = ctrl.hWnd
            
            If Len(strText) <> 0 Then

                'MODIFICATION
                'sending a blank string to ToolText deletes the control from the collection
                'MODIFICATION
                'add ToolTipHeader to strText if requested
                
                If m_ShowHeader Then
                    If LenB(m_TTHeader) Then strText = m_TTHeader & vbNewLine & strText
                End If
                
                .lpszText = strText
'            Else
'                .lpszText = "Tool" & ToolCount + 1
            End If
'
            ' Maintain the maximum tip text
            ' length for fGetToolInfo.
            '
            mnlgMaxTip = fMax(mnlgMaxTip, Len(.lpszText) + 1)
    
        End With
        
        ' Returns 1 on success, 0 on failure
        AddTool = SendMessageT(mnlgHwndTT, TTM_ADDTOOL, 0, ti)
        
    End If

End Function

Public Function AddToolFromHwnd(ByVal srcHwnd As Long, Optional ByVal strText As String) As Boolean

    Dim ti As TOOLINFO
    
    'Ignore hWnds of 0
    If (srcHwnd = 0) Then Exit Function
    
    If (fGetToolInfo(srcHwnd, ti) = False) Then
    
        With ti
            .cbSize = Len(ti)
            '
            ' TTF_IDISHWND must be set to tell the tooltip
            ' control to retrieve the control's rect from
            ' it's hWnd specified in uId.
            '
            .uFlags = TTF_SUBCLASS Or TTF_IDISHWND
            .hWnd = srcHwnd
            .uId = srcHwnd
            
            If Len(strText) <> 0 Then

                'MODIFICATION
                'sending a blank string to ToolText deletes the control from the collection
                'MODIFICATION
                'add ToolTipHeader to strText if requested
                
                If m_ShowHeader Then
                    If LenB(m_TTHeader) Then strText = m_TTHeader & vbNewLine & strText
                End If
                
                .lpszText = strText
            End If
            
            mnlgMaxTip = fMax(mnlgMaxTip, Len(.lpszText) + 1)
    
        End With
        
        ' Returns 1 on success, 0 on failure
        AddToolFromHwnd = SendMessageT(mnlgHwndTT, TTM_ADDTOOL, 0, ti)
        
    End If

End Function

Public Property Get BackColor() As OLE_COLOR

  If (mnlgHwndTT = 0) Then
    Exit Property
  End If
'
' OLE_COLOR is defined in stdole2.tlb
'
  BackColor = SendMessageT(mnlgHwndTT, TTM_GETTIPBKCOLOR, 0, 0)

End Property

Public Property Let BackColor(clr As OLE_COLOR)

  If (mnlgHwndTT = 0) Then
    Exit Property
  End If
  Call SendMessageT(mnlgHwndTT, TTM_SETTIPBKCOLOR, clr, 0)

End Property

Private Sub Class_Initialize()

  m_ShowHeader = False

End Sub

Private Sub Class_Terminate()

  If mnlgHwndTT > 0 Then
    Call DestroyWindow(mnlgHwndTT)
  End If

End Sub

Public Function Create(ByRef frm As Object) As Boolean

  If (mnlgHwndTT = 0) Then
    'Call InitCommonControls
'
' The hwndParent param lets the tooltip window
' be owned by the specified form and be destroyed
' along with it. We'll cleanup in Class_Terminate anyway.
' No WS_EX_TOPMOST or TTS_ALWAYSTIP per Win95 UI rules.
'
    mnlgHwndTT = CreateWindowEx(0, TOOLTIPS_CLASS, vbNullString, TTS_ALWAYSTIP Or WS_EX_TOPMOST, 0, 0, 0, 0, frm.hWnd, 0, App.hInstance, ByVal 0)
  End If
  Create = CBool(mnlgHwndTT)

End Function

Public Property Get DelayTime(dwType As ttDelayTimeConstants) As Long

  If (mnlgHwndTT = 0) Then
    Exit Property
  End If
  DelayTime = SendMessageT(mnlgHwndTT, TTM_GETDELAYTIME, (dwType And ttDelayMask), 0&)

End Property

Public Property Let DelayTime(dwType As ttDelayTimeConstants, _
                              dwMilliSecs As Long)

  If (mnlgHwndTT = 0) Then
    Exit Property
  End If
  Call SendMessageT(mnlgHwndTT, TTM_SETDELAYTIME, (dwType And ttDelayMask), ByVal dwMilliSecs)
' no rtn val

End Property

Private Function fGetStrFromBuffer(ByVal strValue As String) As String

  If InStr(strValue, vbNullChar) Then
    fGetStrFromBuffer = Left$(strValue, InStr(strValue, vbNullChar) - 1)
   Else
'
' If strValue had no null char, the Left$ function
' above would rtn a zero length string ("").
'
    fGetStrFromBuffer = strValue
  End If

End Function

Private Function fGetToolInfo(ByVal lnghwndTool As Long, _
                              ti As TOOLINFO, _
                              Optional ByVal fGetText As Boolean = False) As Boolean

  Dim nItems As Long
  Dim i      As Long

  ti.cbSize = Len(ti)
  If fGetText Then
    ti.lpszText = String$(mnlgMaxTip, 0)
  End If
  nItems = ToolCount
  For i = 0 To nItems - 1
'
' Returns 1 on success, 0 on failure.
'
    If SendMessageT(mnlgHwndTT, TTM_ENUMTOOLS, (i), ti) Then
      If (lnghwndTool = ti.uId) Then
        fGetToolInfo = True
        Exit Function
      End If
    End If
  Next i

End Function

Public Function fIsWindow(ByRef ctrl As Object) As Boolean

  On Error GoTo ErrorHandler
  fIsWindow = CBool(ctrl.hWnd)
ErrorHandler:

End Function

Private Function fLowWord(ByVal lngValue As Long) As Long

'
' Returns the low-order word from a 32-bit value.
'

  Call MoveMemory(fLowWord, lngValue, 2)

End Function

Private Function fMax(ByVal lngParm1 As Long, _
                      ByVal lngParm2 As Long) As Long

'
' Returns the larger of the two values.
'

  If lngParm1 > lngParm2 Then
    fMax = lngParm1
   Else
    fMax = lngParm2
  End If

End Function

Public Property Get ForeColor() As OLE_COLOR

  If (mnlgHwndTT = 0) Then
    Exit Property
  End If
  ForeColor = SendMessageT(mnlgHwndTT, TTM_SETTIPTEXTCOLOR, 0, 0)

End Property

Public Property Let ForeColor(clr As OLE_COLOR)

  If (mnlgHwndTT = 0) Then
    Exit Property
  End If
  Call SendMessageT(mnlgHwndTT, TTM_SETTIPTEXTCOLOR, clr, 0)   ' no rtn val

End Property

Public Function HasToolTip(ByRef ctrl As Control) As Boolean

'MODIFICATION
'Test if control already has aToolTip
'See Command1_Click on demo from for why

  Dim ti As TOOLINFO

  HasToolTip = fGetToolInfo(ctrl.hWnd, ti)

End Function

Public Property Get hWnd() As Long

  hWnd = mnlgHwndTT

End Property

Public Property Get Margin(dwType As ttMarginConstants) As Long

  Dim rc As RECT

  If (mnlgHwndTT = 0) Then
    Exit Property
  End If
  Call SendMessageT(mnlgHwndTT, TTM_GETMARGIN, 0, rc)
  Select Case dwType
   Case ttMarginLeft
    Margin = rc.Left
   Case ttMarginTop
    Margin = rc.Top
   Case ttMarginRight
    Margin = rc.Right
   Case ttMarginBottom
    Margin = rc.Bottom
  End Select

End Property

Public Property Let Margin(dwType As ttMarginConstants, _
                           ByVal cPixels As Long)

  Dim rc As RECT

  If (mnlgHwndTT = 0) Then
    Exit Property
  End If
  Call SendMessageT(mnlgHwndTT, TTM_GETMARGIN, 0, rc)
  Select Case dwType
   Case ttMarginLeft
    rc.Left = cPixels
   Case ttMarginTop
    rc.Top = cPixels
   Case ttMarginRight
    rc.Right = cPixels
   Case ttMarginBottom
    rc.Bottom = cPixels
  End Select
  Call SendMessageT(mnlgHwndTT, TTM_SETMARGIN, 0, rc)

End Property

Public Property Get MaxTipWidth() As Long

'
' If MaxTipWidth is -1, there is no word wrapping and
' text control characters are printed and not
' evaluated (i.e. a vbCrLf shows up as "||")
'

  If (mnlgHwndTT = 0) Then
    Exit Property
  End If
  MaxTipWidth = fLowWord(SendMessageT(mnlgHwndTT, TTM_GETMAXTIPWIDTH, 0, 0))

End Property

Public Property Let MaxTipWidth(ByVal lngWidth As Long)

'
' If MaxTipWidth is -1, there is no word wrapping and
' text control characters are printed and not
' evaluated (i.e. a vbCrLf shows up as "||")
'

  If mnlgHwndTT = 0 Then Exit Property
  
  If lngWidth < 1 Then
    lngWidth = -1
  End If
  
  'NOTE FROM TANNER: in the original version, this line was incorrect - lngWidth MUST be prefaced with ByVal for
  ' the MaxTipWidth function to work.  (Otherwise, the address of lngWidth is passed.)
  Call SendMessageT(mnlgHwndTT, TTM_SETMAXTIPWIDTH, 0, ByVal lngWidth)

End Property

Public Function RemoveTool(ByRef ctrl As Object) As Boolean

  Dim ti As TOOLINFO

  If (mnlgHwndTT = 0) Then
    Exit Function
  End If
  If fGetToolInfo(ctrl.hWnd, ti) Then
    Call SendMessageT(mnlgHwndTT, TTM_DELTOOL, 0, ti)
    RemoveTool = True
  End If

End Function

Public Function RemoveToolByHwnd(ByVal srcHwnd As Long) As Boolean

  Dim ti As TOOLINFO

  If (mnlgHwndTT = 0) Then
    Exit Function
  End If
  If fGetToolInfo(srcHwnd, ti) Then
    Call SendMessageT(mnlgHwndTT, TTM_DELTOOL, 0, ti)
    RemoveToolByHwnd = True
  End If

End Function

Public Property Get ToolCount() As Long

  If (mnlgHwndTT = 0) Then
    Exit Property
  End If
  ToolCount = SendMessageT(mnlgHwndTT, TTM_GETTOOLCOUNT, 0, 0)

End Property

Public Property Get ToolText(ByRef ctrl As Object) As String

  Dim ti As TOOLINFO

  If (mnlgHwndTT = 0) Then
    Exit Property
  End If
  If fGetToolInfo(ctrl.hWnd, ti, True) Then
    ToolText = ToolText & fGetStrFromBuffer(ti.lpszText)
  End If

End Property

Public Property Let ToolText(ByRef ctrl As Object, _
                             ByVal strText As String)

  Dim ti As TOOLINFO

  If (mnlgHwndTT = 0) Then
    Exit Property
  End If
'MODIFICATION
'sending a blank string to ToolText deletes the control from the collection
  If LenB(strText) = 0 Then
    'RemoveTool ctrl
    'Exit Property
  End If
'MODIFICATION
'add ToolTipHeader to strText if requested
  If m_ShowHeader Then
    If LenB(m_TTHeader) Then
      SendMessageT mnlgHwndTT, TTM_SETTITLE, 0&, ByVal m_TTHeader
      'strText = m_TTHeader & vbNewLine & strText
    Else
        SendMessageT mnlgHwndTT, TTM_SETTITLE, 0&, ByVal vbNullString
    End If
  End If
  If fGetToolInfo(ctrl.hWnd, ti) Then
    ti.lpszText = strText
    mnlgMaxTip = fMax(mnlgMaxTip, Len(strText) + 1)
'
' The tooltip won't appear for the control
' if lpszText is an empty string
'
    Call SendMessageT(mnlgHwndTT, TTM_UPDATETIPTEXT, 0, ti)
  End If

End Property

Public Property Get ToolTipHandle() As Long

  ToolTipHandle = mnlgHwndTT

End Property

Public Property Get ToolTipHeader() As String

  ToolTipHeader = m_TTHeader

End Property

Public Property Let ToolTipHeader(ByVal vNewValue As String)

'MODIFICATION
' allows you to set a standard opening header for all tooltips

  m_TTHeader = vNewValue
'if you supply a Header then set it to default On
  m_ShowHeader = True

End Property

Public Property Get ToolTipHeaderShow() As Boolean

  ToolTipHeaderShow = m_ShowHeader

End Property

Public Property Let ToolTipHeaderShow(ByVal vNewValue As Boolean)

'MODIFICATION
' turn the ToolTipHeader on/off on a per control basis

  m_ShowHeader = vNewValue

End Property

':)Code Fixer V2.1.8 (9/06/2004 2:59:58 PM) 198 + 413 = 611 Lines Thanks Ulli for inspiration and lots of code.

