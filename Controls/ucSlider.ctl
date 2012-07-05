VERSION 5.00
Begin VB.UserControl ucSlider 
   AutoRedraw      =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   510
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   720
   ClipControls    =   0   'False
   LockControls    =   -1  'True
   ScaleHeight     =   34
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   48
   Begin VB.Image imgSlider 
      Height          =   240
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgRail 
      Height          =   255
      Left            =   285
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "ucSlider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'========================================================================================
' User control:  ucSlider.ctl
' Author:        Carles P.V. ©2001-2005
' Dependencies:
' Last revision: 2005.05.29 (Original code date: 2001)
' Version:       1.2.0
'========================================================================================

Option Explicit

'-- API:

Private Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, lpRect As RECT2, ByVal lEdge As Long, ByVal grfFlags As Long) As Long

Private Const BDR_SUNKEN      As Long = &HA
Private Const BDR_RAISED      As Long = &H5
Private Const BDR_SUNKENOUTER As Long = &H2
Private Const BDR_RAISEDINNER As Long = &H4
Private Const BF_RECT         As Long = &HF

Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
                         
Private Const HWND_TOP       As Long = 0
Private Const HWND_TOPMOST   As Long = -1
Private Const HWND_NOTOPMOST As Long = -2
Private Const SWP_NOSIZE     As Long = &H1
Private Const SWP_NOMOVE     As Long = &H2
Private Const SWP_NOACTIVATE As Long = &H10
Private Const SWP_SHOWWINDOW As Long = &H40
                         
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT2) As Long

Private Type RECT2
    x1 As Long
    y1 As Long
    x2 As Long
    y2 As Long
End Type

'-- Public enums.:
Public Enum sOrientationConstants
    [Horizontal] = 0
    [Vertical]
End Enum
Public Enum sRailStyleConstants
    [Sunken] = 0
    [Raised]
    [SunkenSoft]
    [RaisedSoft]
    [ByPicture] = 99
End Enum

'-- Private types:
Private Type Point
    x As Single
    y As Single
End Type

'-- Private variables:
Private pv_bSliderHooked As Boolean ' imgSlider hooked
Private pv_uSliderOffset As Point   ' imgSlider anchor point
Private pv_uRailRect     As RECT2   ' Rail rectangle
Private pv_uSliderlRect  As RECT2   ' Slider rectangle
Private pv_lAbsCount     As Long    ' pv_lAbsCount = Max - Min
Private pv_lLastValue    As Long    ' Last slider value
Private pv_lTPPx         As Long    ' TwipsPerPixelX
Private pv_lTPPy         As Long    ' TwipsPerPixelY

'-- Default property values:
Private Const m_def_Enabled      As Boolean = True
Private Const m_def_Orientation  As Long = [Vertical]
Private Const m_def_RailStyle    As Long = [Sunken]
Private Const m_def_ShowValueTip As Boolean = True
Private Const m_def_Min          As Long = 0
Private Const m_def_Max          As Long = 10
Private Const m_def_Value        As Long = 0

'-- Property variables:
Private m_Enabled      As Boolean
Private m_Orientation  As sOrientationConstants
Private m_RailStyle    As sRailStyleConstants
Private m_ShowValueTip As Boolean
Private m_Min          As Long
Private m_Max          As Long
Private m_Value        As Long

'-- Event declarations:
Public Event Click()
Public Event ArrivedFirst()
Public Event ArrivedLast()
Public Event Change()
Public Event MouseDown(Shift As Integer)
Public Event MouseUp(Shift As Integer)



'========================================================================================
' Usercontrol initialization/termination
'========================================================================================

Private Sub UserControl_Initialize()
    
    pv_lTPPx = Screen.TwipsPerPixelX
    pv_lTPPy = Screen.TwipsPerPixelY
End Sub

Private Sub UserControl_Terminate()
    
    'On Error Resume Next
    'If (Not ucSliderTip Is Nothing) Then
    '    Call Unload(ucSliderTip)
    '    Set ucSliderTip = Nothing
    'End If
    'On Error GoTo 0
End Sub

'========================================================================================
' Drawing
'========================================================================================

Private Sub UserControl_Show()
    '-- Draw control
    Call Refresh
End Sub

Private Sub UserControl_Resize()

    On Error Resume Next
    
    '-- Resize control
    If (m_RailStyle = 99 And imgRail.Picture.handle <> 0) Then
    
        Select Case m_Orientation
            
            Case 0 '-- Horizontal
                If (imgSlider.Height < imgRail.Height) Then
                    SIZE (imgRail.Width + 4) * pv_lTPPx, imgRail.Height * pv_lTPPx
                  Else
                    SIZE (imgRail.Width + 4) * pv_lTPPx, imgSlider.Height * pv_lTPPx
                End If
                
            Case 1 '-- Vertical
                If (imgSlider.Width < imgRail.Width) Then
                    SIZE imgRail.Width * pv_lTPPy, (imgRail.Height + 4) * pv_lTPPy
                  Else
                    SIZE imgSlider.Width * pv_lTPPy, (imgRail.Height + 4) * pv_lTPPy
                End If
        End Select
    
      Else
        Select Case m_Orientation
            
            Case 0 '-- Horizontal
                If (Width = 0) Then Width = imgSlider.Width * pv_lTPPx
                Height = imgSlider.Height * pv_lTPPy
                    
            Case 1 '-- Vertical
                If (Height = 0) Then Height = imgSlider.Height * pv_lTPPy
                Width = (imgSlider.Width) * pv_lTPPx
        End Select
    
    End If
    
    '-- Update slider position
    Select Case m_Orientation
    
        Case 0 '-- Horizontal
            If (imgSlider.Height < imgRail.Height And m_RailStyle = 99 And imgRail <> 0) Then
                imgSlider.TOp = (imgRail.Height - imgSlider.Height) \ 2
              Else
                imgSlider.TOp = 0
            End If
            imgSlider.Left = (m_Value - m_Min) * (ScaleWidth - imgSlider.Width) / pv_lAbsCount
        
        Case 1 '-- Vertical
            If (imgSlider.Width < imgRail.Width And m_RailStyle = 99 And imgRail <> 0) Then
                imgSlider.Left = (imgRail.Width - imgSlider.Width) \ 2
              Else
                imgSlider.Left = 0
            End If
            imgSlider.TOp = ScaleHeight - imgSlider.Height - (m_Value - m_Min) * (ScaleHeight - imgSlider.Height) / pv_lAbsCount
    End Select
    
    '-- Define rail rectangle
    Select Case m_Orientation
        
        Case 0 '-- Horizontal
            With pv_uRailRect
                .y1 = (imgSlider.Height - 4) \ 2
                .y2 = .y1 + 4
                .x1 = imgSlider.Width \ 2 - 2
                .x2 = .x1 + ScaleWidth - imgSlider.Width + 4
            End With
                
        Case 1 '-- Vertical
            With pv_uRailRect
                .y1 = imgSlider.Height \ 2 - 2
                .y2 = .y1 + ScaleHeight - imgSlider.Height + 4
                .x1 = (imgSlider.Width - 4) \ 2
                .x2 = .x1 + 4
            End With
    End Select
    
    '-- Refresh control
    Call Refresh
    
    On Error GoTo 0
End Sub

Private Sub Refresh()
    
    '-- Clear control
    Call Cls
    
    '-- Draw rail...
    On Error Resume Next
    
    If (m_RailStyle = 99) Then
    
        Select Case m_Orientation
        
            Case 0 '-- Horizontal
                Call PaintPicture(imgRail, 2, (ScaleHeight - imgRail.Height) \ 2)
                 
            Case 1 '-- Vertical
                Call PaintPicture(imgRail, (ScaleWidth - imgRail.Width) \ 2, 2)
        End Select
        
      Else
        Call DrawEdge(hDC, pv_uRailRect, Choose(m_RailStyle + 1, BDR_SUNKEN, BDR_RAISED, BDR_SUNKENOUTER, BDR_RAISEDINNER, 0), BF_RECT)
   End If
    
    '-- ...and slider
    If (imgSlider.Picture) Then
        
        '-- Paint image
        Call PaintPicture(imgSlider, imgSlider.Left, imgSlider.TOp)
      
      Else
        
        With pv_uSliderlRect
            .x1 = imgSlider.Left
            .y1 = imgSlider.TOp
            .x2 = .x1 + imgSlider.Width
            .y2 = .y1 + imgSlider.Height
            
            '-- Erase background
            Line (.x1, .y1)-(.x2 - 1, .y2 - 1), vbButtonFace, BF
            
            '-- 3D edge
            Call DrawEdge(hDC, pv_uSliderlRect, BDR_RAISED, BF_RECT)
        End With
    End If
    
    '-- Show value tip
    If (m_ShowValueTip And pv_bSliderHooked) Then
        Call pvShowTip
    End If
    
    On Error GoTo 0
End Sub

'========================================================================================
' Scrolling
'========================================================================================

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If (Me.Enabled) Then
    
        With imgSlider
            
            '-- Hook slider, get offsets and show tip
            If (Button = vbLeftButton) Then
               
                pv_bSliderHooked = True
                
                '-- Mouse over slider
                If (x >= .Left And x < .Left + .Width And y >= .TOp And y < .TOp + .Height) Then
                   
                    pv_uSliderOffset.x = x - .Left
                    pv_uSliderOffset.y = y - .TOp
                
                Else
                '-- Mouse over rail
                    pv_uSliderOffset.x = .Width \ 2
                    pv_uSliderOffset.y = .Height \ 2
                    Call UserControl_MouseMove(Button, Shift, x, y)
                End If
                
                '-- Show tip
                If (m_ShowValueTip) Then
                    Call pvShowTip
                End If
                
                RaiseEvent MouseDown(Shift)
            End If
        End With
    End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If (pv_bSliderHooked) Then
        
        '-- Check limits
        With imgSlider
        
            Select Case m_Orientation
            
                Case 0 '-- Horizontal
                    If (x - pv_uSliderOffset.x < 0) Then
                        .Left = 0
                      ElseIf (x - pv_uSliderOffset.x > ScaleWidth - .Width) Then
                        .Left = ScaleWidth - .Width
                      Else
                        .Left = x - pv_uSliderOffset.x
                    End If
            
                Case 1 '-- Vertical
                    If (y - pv_uSliderOffset.y < 0) Then
                        .TOp = 0
                      ElseIf (y - pv_uSliderOffset.y > ScaleHeight - .Height) Then
                        .TOp = ScaleHeight - .Height
                      Else
                        .TOp = y - pv_uSliderOffset.y
                    End If
            End Select
        End With
        
        '-- Get value from imgSlider position
        Value = pvGetValue
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    '-- Click event (If mouse over control area)
    If (x >= 0 And x < ScaleWidth And y >= 0 And y < ScaleHeight And Button = vbLeftButton) Then
        RaiseEvent Click
    End If
    
    '-- MouseUp event (imgSlider has been hooked)
    If (pv_bSliderHooked) Then
        RaiseEvent MouseUp(Shift)
    End If
    
    '-- Unhook slider and hide value tip
    pv_bSliderHooked = False
    'Call Unload(ucSliderTip)
End Sub

'========================================================================================
' Private
'========================================================================================

Private Function pvGetValue() As Long
    
    On Error Resume Next
    
    Select Case m_Orientation
    
        Case 0 '-- Horizontal
            pvGetValue = imgSlider.Left / (ScaleWidth - imgSlider.Width) * pv_lAbsCount + m_Min
            imgSlider.Left = (pvGetValue - m_Min) * (ScaleWidth - imgSlider.Width) / pv_lAbsCount
        
        Case 1 '-- Vertical
            pvGetValue = (ScaleHeight - imgSlider.Height - imgSlider.TOp) / (ScaleHeight - imgSlider.Height) * pv_lAbsCount + m_Min
            imgSlider.TOp = ScaleHeight - imgSlider.Height - (pvGetValue - m_Min) * (ScaleHeight - imgSlider.Height) / pv_lAbsCount
    End Select
    
    On Error GoTo 0
End Function

Private Sub pvResetSlider()

    Select Case m_Orientation
        
        Case 0 '-- Horizontal
            Call imgSlider.Move(0, 0)
             
        Case 1 '-- Vertical
            Call imgSlider.Move(0, ScaleHeight - imgSlider.Height)
    End Select
End Sub

Private Sub pvShowTip()
    
'  Dim uRect As RECT2
'  Dim x     As Long
'  Dim y     As Long

'    On Error Resume Next
'
'    Call GetWindowRect(hWnd, uRect)
'
'    With ucSliderTip
'
'        .lblTip.Width = .TextWidth(m_Value)
'        .lblTip.Caption = m_Value
'        Call .lblTip.Refresh
'
'        Select Case m_Orientation
'
'            Case 0 '-- Horizontal
'                x = uRect.x1 + imgSlider.Left + (imgSlider.Width - .lblTip.Width - 4) \ 2
'                y = uRect.y1 + imgSlider.Top - .lblTip.Height - 5
'
'            Case 1 '-- Vertical
'                x = uRect.x1 + imgSlider.Left - .lblTip.Width - 6
'                y = uRect.y1 + imgSlider.Top + (imgSlider.Height - .lblTip.Height - 4) \ 2
'        End Select
'
'        '-- Set Tip position...
'        Call .Move(x * pv_lTPPx, y * pv_lTPPy, (.lblTip.Width + 4) * pv_lTPPx, (.lblTip.Height + 3) * pv_lTPPy)
'
'        '-- ...and show it
'        Call SetWindowPos(.hWnd, HWND_TOP, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW)
'    End With
'
'    On Error GoTo 0
End Sub

'========================================================================================
' Init/Read/Write properties
'========================================================================================

Private Sub UserControl_InitProperties()

    m_Enabled = m_def_Enabled
    m_Orientation = m_def_Orientation
    m_RailStyle = m_def_RailStyle
    m_ShowValueTip = m_def_ShowValueTip
    m_Min = m_def_Min
    m_Max = m_def_Max
    m_Value = m_def_Value
    
    pv_lAbsCount = 10
    pv_lLastValue = m_Value
    Call pvResetSlider
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    With PropBag
    
        UserControl.BackColor = .ReadProperty("BackColor", vbButtonFace)
        UserControl.MousePointer = .ReadProperty("MousePointer", vbDefault)
        Set MouseIcon = .ReadProperty("MouseIcon", Nothing)
        m_Enabled = .ReadProperty("Enabled", m_def_Enabled)
        m_Orientation = .ReadProperty("Orientation", m_def_Orientation)
        m_RailStyle = .ReadProperty("RailStyle", m_def_RailStyle)
        m_ShowValueTip = .ReadProperty("ShowValueTip", m_def_ShowValueTip)
        m_Min = .ReadProperty("Min", m_def_Min)
        m_Max = .ReadProperty("Max", m_def_Max)
        m_Value = .ReadProperty("Value", m_def_Value)
        
        Set imgSlider.Picture = .ReadProperty("SliderImage", Nothing)
        Set imgRail = .ReadProperty("RailImage", Nothing)
        
        '-- Get absolute count and set imgSlider position
        pv_lAbsCount = m_Max - m_Min
        pv_lLastValue = m_Value
        imgSlider.Left = (m_Value - m_Min) * (ScaleWidth - imgSlider.Width) / pv_lAbsCount
        imgSlider.TOp = (ScaleHeight - imgSlider.Height) - (m_Value - m_Min) * (ScaleHeight - imgSlider.Height) / pv_lAbsCount
    End With
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    With PropBag
        Call .WriteProperty("BackColor", UserControl.BackColor, vbButtonFace)
        Call .WriteProperty("Enabled", m_Enabled, m_def_Enabled)
        Call .WriteProperty("MouseIcon", MouseIcon, Nothing)
        Call .WriteProperty("MousePointer", UserControl.MousePointer, vbDefault)
        Call .WriteProperty("SliderImage", imgSlider.Picture, Nothing)
        Call .WriteProperty("Orientation", m_Orientation, m_def_Orientation)
        Call .WriteProperty("RailImage", imgRail, Nothing)
        Call .WriteProperty("RailStyle", m_RailStyle, m_def_RailStyle)
        Call .WriteProperty("ShowValueTip", m_ShowValueTip, m_def_ShowValueTip)
        Call .WriteProperty("Min", m_Min, m_def_Min)
        Call .WriteProperty("Max", m_Max, m_def_Max)
        Call .WriteProperty("Value", m_Value, m_def_Value)
    End With
End Sub

'========================================================================================
' Properties
'========================================================================================

Public Property Get Enabled() As Boolean
    Enabled = m_Enabled
End Property
Public Property Let Enabled(ByVal New_Enabled As Boolean)
    m_Enabled = New_Enabled
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property
Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    Call Refresh
End Property

Public Property Get MousePointer() As MousePointerConstants
    MousePointer = UserControl.MousePointer
End Property
Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)
    UserControl.MousePointer() = New_MousePointer
End Property

Public Property Get MouseIcon() As StdPicture
    Set MouseIcon = UserControl.MouseIcon
End Property
Public Property Set MouseIcon(ByVal New_MouseIcon As StdPicture)
    Set UserControl.MouseIcon = New_MouseIcon
End Property

Public Property Get Max() As Long
    Max = m_Max
End Property
Public Property Let Max(ByVal New_Max As Long)
    If (New_Max <= m_Min) Then Call Err.Raise(380)
    m_Max = New_Max
    pv_lAbsCount = m_Max - m_Min
End Property

Public Property Get Min() As Long
    Min = m_Min
End Property
Public Property Let Min(ByVal New_Min As Long)
    If (New_Min >= m_Max) Then Err.Raise 380
    m_Min = New_Min
    Value = New_Min
    pv_lAbsCount = m_Max - m_Min
End Property

Public Property Get Value() As Long
Attribute Value.VB_UserMemId = 0
Attribute Value.VB_MemberFlags = "200"
    Value = m_Value
End Property
Public Property Let Value(ByVal New_Value As Long)

    If (New_Value < m_Min Or New_Value > m_Max) Then Call Err.Raise(380)
    
    m_Value = New_Value
        
    If (m_Value <> pv_lLastValue) Then
        
        If (Not pv_bSliderHooked) Then
                   
            Select Case m_Orientation

                Case 0 '-- Horizontal
                    imgSlider.Left = (New_Value - m_Min) * (ScaleWidth - imgSlider.Width) / pv_lAbsCount
                
                Case 1 '-- Vertical
                    imgSlider.TOp = ScaleHeight - imgSlider.Height - (New_Value - m_Min) * (ScaleHeight - imgSlider.Height) / pv_lAbsCount
            End Select
        End If
        
        Call Refresh
        pv_lLastValue = m_Value
        
        RaiseEvent Change
        If (m_Value = m_Max) Then RaiseEvent ArrivedLast
        If (m_Value = m_Min) Then RaiseEvent ArrivedFirst
    End If
End Property

Public Property Get Orientation() As sOrientationConstants
    Orientation = m_Orientation
End Property
Public Property Let Orientation(ByVal New_Orientation As sOrientationConstants)
    m_Orientation = New_Orientation
    Call pvResetSlider
    Call UserControl_Resize
End Property

Public Property Get RailStyle() As sRailStyleConstants
    RailStyle = m_RailStyle
End Property
Public Property Let RailStyle(ByVal New_RailStyle As sRailStyleConstants)
    m_RailStyle = New_RailStyle
    Call UserControl_Resize
End Property

Public Property Get SliderImage() As Picture
    Set SliderImage = imgSlider.Picture
End Property
Public Property Set SliderImage(ByVal New_SliderImage As Picture)
    Set imgSlider.Picture = New_SliderImage
    Call UserControl_Resize
End Property

Public Property Get RailImage() As Picture
    Set RailImage = imgRail.Picture
End Property
Public Property Set RailImage(ByVal New_RailImage As Picture)
    Set imgRail.Picture = New_RailImage
    Call UserControl_Resize
End Property

Public Property Get ShowValueTip() As Boolean
    ShowValueTip = m_ShowValueTip
End Property
Public Property Let ShowValueTip(ByVal New_ShowValueTip As Boolean)
    m_ShowValueTip = New_ShowValueTip
End Property
