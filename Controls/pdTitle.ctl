VERSION 5.00
Begin VB.UserControl pdTitle 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   450
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3930
   DrawStyle       =   5  'Transparent
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HasDC           =   0   'False
   ScaleHeight     =   30
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   262
   ToolboxBitmap   =   "pdTitle.ctx":0000
End
Attribute VB_Name = "pdTitle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Collapsible Title Label+Button control
'Copyright 2014-2016 by Tanner Helland
'Created: 19/October/14
'Last updated: 12/February/16
'Last update: integrate with pdUCSupport, which cuts a ton of redundant code
'
'In a surprise to precisely no one, PhotoDemon has some unique needs when it comes to user controls - needs that
' the intrinsic VB controls can't handle.  These range from the obnoxious (lack of an "autosize" property for
' anything but labels) to the critical (no Unicode support).
'
'As such, I've created many of my own UCs for the program.  All are owner-drawn, with the goal of maintaining
' visual fidelity across the program, while also enabling key features like Unicode support.
'
'A few notes on this "title" control, specifically:
'
' 1) Captioning is (mostly) handled by the pdCaption class, so autosizing of overlong text is supported.
' 2) High DPI settings are handled automatically.
' 3) A hand cursor is automatically applied, and clicks are returned via the Click event.
' 4) Coloration is automatically handled by PD's internal theming engine.
' 5) This title control is meant to be used above collapsible UI panels.  It auto-toggles between a "true" and
'     "false" state, and this state is returned directly inside the Click() event.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'This control really only needs one event raised - Click
Public Event Click(ByVal newState As Boolean)

'Because VB focus events are wonky, especially when we use CreateWindow within a UC, this control raises its own
' specialized focus events.  If you need to track focus, use these instead of the default VB functions.
Public Event GotFocusAPI()
Public Event LostFocusAPI()

'Rect where the caption is rendered.  This is calculated by UpdateControlLayout, and it needs to be revisited if the
' caption changes, or the control size changes.
Private m_CaptionRect As RECT

'Current title state (TRUE when arrow is pointing down, e.g. the associated container is "open")
Private m_TitleState As Boolean

'User control support class.  Historically, many classes (and associated subclassers) were required by each user control,
' but I've since attempted to wrap these into a single master control support class.
Private WithEvents ucSupport As pdUCSupport
Attribute ucSupport.VB_VarHelpID = -1

'Local list of themable colors.  This list includes all potential colors used by this class, regardless of state change
' or internal control settings.  The list is updated by calling the UpdateColorList function.
' (Note also that this list does not include variants, e.g. "BorderColor" vs "BorderColor_Hovered".  Variant values are
'  automatically calculated by the color management class, and they are retrieved by passing boolean modifiers to that
'  class, rather than treating every imaginable variant as a separate constant.)
Private Enum PDTITLE_COLOR_LIST
    [_First] = 0
    PDT_Background = 0
    PDT_Caption = 1
    PDT_Arrow = 2
    PDT_Border = 3
    [_Last] = 3
    [_Count] = 4
End Enum

'Color retrieval and storage is handled by a dedicated class; this allows us to optimize theme interactions,
' without worrying about the details locally.
Private m_Colors As pdThemeColors

'Caption is handled just like the common control label's caption property.  It is valid at design-time, and any translation,
' if present, will not be processed until run-time.
' IMPORTANT NOTE: only the ENGLISH caption is returned.  I don't have a reason for returning a translated caption (if any),
'                  but I can revisit in the future if it ever becomes relevant.
Public Property Get Caption() As String
Attribute Caption.VB_UserMemId = -518
    Caption = ucSupport.GetCaptionText
End Property

Public Property Let Caption(ByRef newCaption As String)
    
    ucSupport.SetCaptionText newCaption
    PropertyChanged "Caption"
    
    'Access keys must be handled manually.
    Dim ampPos As Long
    ampPos = InStr(1, newCaption, "&", vbBinaryCompare)
    
    If (ampPos > 0) And (ampPos < Len(newCaption)) Then
    
        'Get the character immediately following the ampersand, and dynamically assign it
        Dim accessKeyChar As String
        accessKeyChar = Mid$(newCaption, ampPos + 1, 1)
        UserControl.AccessKeys = accessKeyChar
    
    Else
        UserControl.AccessKeys = ""
    End If
    
End Property

'The Enabled property is a bit unique; see http://msdn.microsoft.com/en-us/library/aa261357%28v=vs.60%29.aspx
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_UserMemId = -514
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal newValue As Boolean)
    UserControl.Enabled = newValue
    RedrawBackBuffer
    PropertyChanged "Enabled"
End Property

Public Property Get FontBold() As Boolean
    FontBold = ucSupport.GetCaptionFontBold
End Property

Public Property Let FontBold(ByVal newValue As Boolean)
    ucSupport.SetCaptionFontBold newValue
    PropertyChanged "FontBold"
End Property

Public Property Get FontSize() As Single
    FontSize = ucSupport.GetCaptionFontSize
End Property

Public Property Let FontSize(ByVal newSize As Single)
    ucSupport.SetCaptionFontSize newSize
    PropertyChanged "FontSize"
End Property

'hWnds aren't exposed by default
Public Property Get hWnd() As Long
Attribute hWnd.VB_UserMemId = -515
    hWnd = UserControl.hWnd
End Property

'State is toggled on each click.  TRUE means the accompanying panel should be OPEN.
Public Property Get Value() As Boolean
Attribute Value.VB_UserMemId = 0
    Value = m_TitleState
End Property

Public Property Let Value(ByVal newState As Boolean)
    If newState <> m_TitleState Then
        m_TitleState = newState
        RedrawBackBuffer
        PropertyChanged "Value"
    End If
End Property

'To support high-DPI settings properly, we expose some specialized move+size functions
Public Function GetLeft() As Long
    GetLeft = ucSupport.GetControlLeft
End Function

Public Sub SetLeft(ByVal newLeft As Long)
    ucSupport.RequestNewPosition newLeft, , True
End Sub

Public Function GetTop() As Long
    GetTop = ucSupport.GetControlTop
End Function

Public Sub SetTop(ByVal newTop As Long)
    ucSupport.RequestNewPosition , newTop, True
End Sub

Public Function GetWidth() As Long
    GetWidth = ucSupport.GetControlWidth
End Function

Public Sub SetWidth(ByVal newWidth As Long)
    ucSupport.RequestNewSize newWidth, , True
End Sub

Public Function GetHeight() As Long
    GetHeight = ucSupport.GetControlHeight
End Function

Public Sub SetHeight(ByVal newHeight As Long)
    ucSupport.RequestNewSize , newHeight, True
End Sub

Public Sub SetPositionAndSize(ByVal newLeft As Long, ByVal newTop As Long, ByVal newWidth As Long, ByVal newHeight As Long)
    ucSupport.RequestFullMove newLeft, newTop, newWidth, newHeight, True
End Sub

'A few key events are also handled
Private Sub ucSupport_KeyUpCustom(ByVal Shift As ShiftConstants, ByVal vkCode As Long, markEventHandled As Boolean)

    'When space is released, redraw the button to match
    If (vkCode = VK_SPACE) Or (vkCode = VK_RETURN) Then

        If Me.Enabled Then
            m_TitleState = Not m_TitleState
            RedrawBackBuffer
            RaiseEvent Click(m_TitleState)
        End If
        
    End If

End Sub

'Only left clicks raise Click() events
Private Sub ucSupport_MouseDownCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    
    If Me.Enabled And ((Button And pdLeftButton) <> 0) Then
    
        'Toggle title state and redraw
        m_TitleState = Not m_TitleState
        
        'Note that drawing flags are handled by MouseDown/Up.  Click() is only used for raising a matching Click() event.
        RaiseEvent Click(m_TitleState)
        RedrawBackBuffer
        
    End If
    
End Sub

Private Sub ucSupport_MouseEnter(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    ucSupport.RequestCursor IDC_HAND
    RedrawBackBuffer
End Sub

'When the mouse leaves the UC, we must repaint the button (as it's no longer hovered)
Private Sub ucSupport_MouseLeave(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    RedrawBackBuffer
End Sub

Private Sub ucSupport_GotFocusAPI()
    RaiseEvent GotFocusAPI
    RedrawBackBuffer
End Sub

Private Sub ucSupport_LostFocusAPI()
    RaiseEvent LostFocusAPI
    RedrawBackBuffer
End Sub

Private Sub ucSupport_RepaintRequired(ByVal updateLayoutToo As Boolean)
    If updateLayoutToo Then UpdateControlLayout
    RedrawBackBuffer
End Sub

Private Sub ucSupport_WindowResize(ByVal newWidth As Long, ByVal newHeight As Long)
    UpdateControlLayout
End Sub

Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
    m_TitleState = Not m_TitleState
    RaiseEvent Click(m_TitleState)
End Sub

Private Sub UserControl_Initialize()
    
    'Initialize a master user control support class
    Set ucSupport = New pdUCSupport
    ucSupport.RegisterControl UserControl.hWnd
    
    'Request any control-specific functionality
    ucSupport.RequestExtraFunctionality True, True
    ucSupport.SpecifyRequiredKeys VK_SPACE, VK_RETURN
    ucSupport.RequestCaptionSupport
    ucSupport.SetCaptionAutomaticPainting False
    
    'Prep the color manager and load default colors
    Set m_Colors = New pdThemeColors
    Dim colorCount As PDTITLE_COLOR_LIST: colorCount = [_Count]
    m_Colors.InitializeColorList "PDTitle", colorCount
    If Not g_IsProgramRunning Then UpdateColorList
    
    'Update the control size parameters at least once
    UpdateControlLayout
                
End Sub

'Set default properties
Private Sub UserControl_InitProperties()
    Caption = ""
    FontBold = False
    FontSize = 10
    Value = True
End Sub

'At run-time, painting is handled by PD's pdWindowPainter class.  In the IDE, however, we must rely on VB's internal paint event.
Private Sub UserControl_Paint()
    ucSupport.RequestIDERepaint UserControl.hDC
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        Caption = .ReadProperty("Caption", "")
        FontBold = .ReadProperty("FontBold", False)
        FontSize = .ReadProperty("FontSize", 10)
        Value = .ReadProperty("Value", True)
    End With
End Sub

Private Sub UserControl_Resize()
    If Not g_IsProgramRunning Then ucSupport.RequestRepaint True
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "Caption", ucSupport.GetCaptionText, ""
        .WriteProperty "FontBold", ucSupport.GetCaptionFontBold, False
        .WriteProperty "FontSize", ucSupport.GetCaptionFontSize, 10
        .WriteProperty "Value", m_TitleState, True
    End With
End Sub

'Because this control automatically forces all internal buttons to identical sizes, we have to recalculate a number
' of internal sizing metrics whenever the control size changes.
Private Sub UpdateControlLayout()

    'Retrieve DPI-aware control dimensions from the support class
    Dim bWidth As Long, bHeight As Long
    bWidth = ucSupport.GetBackBufferWidth
    bHeight = ucSupport.GetBackBufferHeight
    
    Const hTextPadding As Long = 2&, vTextPadding As Long = 2&
    
    'Next, determine the positioning of the caption, if present.  (ucSupport.GetCaptionBottom tells us where the
    ' caption text ends vertically.)
    If ucSupport.IsCaptionActive Then
        
        If ucSupport.GetCaptionHeight + FixDPI(vTextPadding) * 2 <> bHeight Then
            ucSupport.RequestNewSize bWidth, ucSupport.GetCaptionHeight + FixDPI(vTextPadding) * 2, False
        End If
        
        'The control and backbuffer are now guaranteed to be the proper size.
        
        'For caption rendering purposes, we need to determine a target rectangle for the caption itself.  The ucSupport class
        ' will automatically fit the caption within this area, regardless of the currently selected font size.
        With m_CaptionRect
            .Left = FixDPI(hTextPadding)
            .Top = FixDPI(vTextPadding)
            .Bottom = bHeight - FixDPI(vTextPadding)
            
            'The right measurement is the only complicated one, as it requires padding so we have room to render the drop-down arrow.
            .Right = bWidth - FixDPI(hTextPadding) * 2 - bHeight
            
            'Notify the caption renderer of this new caption position, which it will use to automatically adjust its font, as necessary
            ucSupport.SetCaptionCustomPosition .Left, .Top, .Right - .Left, .Bottom - .Top
        End With
        
    End If
        
    'No other special preparation is required for this control, so proceed with recreating the back buffer
    RedrawBackBuffer
            
End Sub

'Before this control does any painting, we need to retrieve relevant colors from PD's primary theming class.  Note that this
' step must also be called if/when PD's visual theme settings change.
Private Sub UpdateColorList()
    With m_Colors
        .LoadThemeColor PDT_Background, "Background", IDE_WHITE
        .LoadThemeColor PDT_Caption, "Caption", IDE_GRAY
        .LoadThemeColor PDT_Arrow, "Arrow", IDE_BLUE
        .LoadThemeColor PDT_Border, "Border", IDE_BLUE
    End With
End Sub

'External functions can call this to request a redraw.  This is helpful for live-updating theme settings, as in the Preferences dialog.
Public Sub UpdateAgainstCurrentTheme()
    UpdateColorList
    If g_IsProgramRunning Then ucSupport.UpdateAgainstThemeAndLanguage
End Sub

'Use this function to completely redraw the back buffer from scratch.  Note that this is computationally expensive compared to just flipping the
' existing buffer to the screen, so only redraw the backbuffer if the control state has somehow changed.
Private Sub RedrawBackBuffer()
    
    Dim ctlFillColor As Long
    ctlFillColor = m_Colors.RetrieveColor(PDT_Background, Me.Enabled, , ucSupport.IsMouseInside)
    
    'Request the back buffer DC, and ask the support module to erase any existing rendering for us.
    Dim bufferDC As Long
    bufferDC = ucSupport.GetBackBufferDC(True, ctlFillColor)
    
    Dim bWidth As Long, bHeight As Long
    bWidth = ucSupport.GetBackBufferWidth
    bHeight = ucSupport.GetBackBufferHeight
    
    Dim textColor As Long, arrowColor As Long, ctlTopLineColor As Long
    arrowColor = m_Colors.RetrieveColor(PDT_Arrow, Me.Enabled, , ucSupport.IsMouseInside)
    ctlTopLineColor = m_Colors.RetrieveColor(PDT_Border, Me.Enabled, ucSupport.DoIHaveFocus, ucSupport.IsMouseInside)
    textColor = m_Colors.RetrieveColor(PDT_Caption, Me.Enabled, , ucSupport.IsMouseInside)
    
    If ucSupport.IsCaptionActive Then
        ucSupport.SetCaptionCustomColor textColor
        ucSupport.PaintCaptionManually
    End If
    
    If g_IsProgramRunning Then
    
        'Next, paint the drop-down arrow.  To simplify calculations, we first calculate a boundary rect.
        Dim arrowRect As RECTF
        arrowRect.Left = bWidth - bHeight - FixDPI(2)
        arrowRect.Top = 1
        arrowRect.Height = bHeight - 2
        arrowRect.Width = bHeight - 2
        
        Dim arrowPt1 As POINTFLOAT, arrowPt2 As POINTFLOAT, arrowPt3 As POINTFLOAT
                    
        'The orientation of the arrow varies depending on open/close state.
        
        'Corresponding panel is open, so arrow points down
        If m_TitleState Then
        
            arrowPt1.x = arrowRect.Left + FixDPIFloat(4)
            arrowPt1.y = arrowRect.Top + (arrowRect.Height / 2) - FixDPIFloat(2)
            
            arrowPt3.x = (arrowRect.Left + arrowRect.Width) - FixDPIFloat(4)
            arrowPt3.y = arrowPt1.y
            
            arrowPt2.x = arrowPt1.x + (arrowPt3.x - arrowPt1.x) / 2
            arrowPt2.y = arrowPt1.y + FixDPIFloat(3)
            
        'Corresponding panel is closed, so arrow points left
        Else
        
            arrowPt1.x = arrowRect.Left + (arrowRect.Width / 2) + FixDPIFloat(2)
            arrowPt1.y = arrowRect.Top + FixDPIFloat(4)
        
            arrowPt3.x = arrowPt1.x
            arrowPt3.y = (arrowRect.Top + arrowRect.Height) - FixDPIFloat(4)
        
            arrowPt2.x = arrowPt1.x - FixDPIFloat(3)
            arrowPt2.y = arrowPt1.y + (arrowPt3.y - arrowPt1.y) / 2
        
        End If
        
        Dim arrowWidth As Single
        If ucSupport.IsMouseInside Then arrowWidth = 2 Else arrowWidth = 1
        GDI_Plus.GDIPlusDrawLineToDC bufferDC, arrowPt1.x, arrowPt1.y, arrowPt2.x, arrowPt2.y, arrowColor, 255, 2, True, GP_LC_Round
        GDI_Plus.GDIPlusDrawLineToDC bufferDC, arrowPt2.x, arrowPt2.y, arrowPt3.x, arrowPt3.y, arrowColor, 255, 2, True, GP_LC_Round
        
        'Finally, frame the control.  At present, this consists of two gradient lines - one across the top, the other down the right side.
        GDI_Plus.GDIPlusDrawGradientLineToDC bufferDC, 0#, 0#, bWidth - 1, 0#, ctlFillColor, ctlTopLineColor, 255, 255, 1, True, GP_LC_Round
        GDI_Plus.GDIPlusDrawGradientLineToDC bufferDC, bWidth - 1, 0#, bWidth - 1, bHeight, ctlTopLineColor, ctlFillColor, 255, 255, 1, True, GP_LC_Round
        
    End If
    
    'Paint the final result to the screen, as relevant
    ucSupport.RequestRepaint
    
End Sub

'By design, PD prefers to not use design-time tooltips.  Apply tooltips at run-time, using this function.
' (IMPORTANT NOTE: translations are handled automatically.  Always pass the original English text!)
Public Sub AssignTooltip(ByVal newTooltip As String, Optional ByVal newTooltipTitle As String, Optional ByVal newTooltipIcon As TT_ICON_TYPE = TTI_NONE)
    ucSupport.AssignTooltip UserControl.ContainerHwnd, newTooltip, newTooltipTitle, newTooltipIcon
End Sub
