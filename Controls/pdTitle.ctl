VERSION 5.00
Begin VB.UserControl pdTitle 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
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
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "pdTitle.ctx":0000
End
Attribute VB_Name = "pdTitle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Collapsible Title Label/Button control
'Copyright 2014-2015 by Tanner Helland
'Created: 19/October/14
'Last updated: 29/October/15
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

'Rect where the caption is rendered.  This is calculated by updateControlLayout, and it needs to be revisited if the
' caption changes, or the control size changes.
Private m_CaptionRect As RECT

'Current back color
Private m_BackColor As OLE_COLOR

'Current title state (TRUE when arrow is pointing down, e.g. the associated container is "open")
Private m_TitleState As Boolean

'User control support class.  Historically, many classes (and associated subclassers) were required by each user control,
' but I've since attempted to wrap these into a single master control support class.
Private WithEvents ucSupport As pdUCSupport
Attribute ucSupport.VB_VarHelpID = -1

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
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal newValue As Boolean)
    
    UserControl.Enabled = newValue
    PropertyChanged "Enabled"
    
    'Redraw the control
    redrawBackBuffer
    
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
    hWnd = UserControl.hWnd
End Property

'Container hWnd must be exposed for external tooltip handling
Public Property Get containerHwnd() As Long
    containerHwnd = UserControl.containerHwnd
End Property

'State is toggled on each click.  TRUE means the accompanying panel should be OPEN.
Public Property Get Value() As Boolean
    Value = m_TitleState
End Property

Public Property Let Value(ByVal newState As Boolean)
    If newState <> m_TitleState Then
        m_TitleState = newState
        PropertyChanged "Value"
        redrawBackBuffer
    End If
End Property

'A few key events are also handled
Private Sub ucSupport_KeyUpCustom(ByVal Shift As ShiftConstants, ByVal vkCode As Long, markEventHandled As Boolean)

    'When space is released, redraw the button to match
    If (vkCode = VK_SPACE) Or (vkCode = VK_RETURN) Then

        If Me.Enabled Then
            m_TitleState = Not m_TitleState
            redrawBackBuffer
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
        redrawBackBuffer
        
    End If
    
End Sub

Private Sub ucSupport_MouseEnter(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    ucSupport.RequestCursor IDC_HAND
    redrawBackBuffer
End Sub

'When the mouse leaves the UC, we must repaint the button (as it's no longer hovered)
Private Sub ucSupport_MouseLeave(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    ucSupport.RequestCursor IDC_DEFAULT
    redrawBackBuffer
End Sub

Private Sub ucSupport_GotFocusAPI()
    RaiseEvent GotFocusAPI
End Sub

Private Sub ucSupport_LostFocusAPI()
    RaiseEvent LostFocusAPI
End Sub

Private Sub ucSupport_RepaintRequired(ByVal updateLayoutToo As Boolean)
    If updateLayoutToo Then updateControlLayout
    redrawBackBuffer
End Sub

Private Sub ucSupport_WindowResize(ByVal newWidth As Long, ByVal newHeight As Long)
    updateControlLayout
End Sub

Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
    m_TitleState = Not m_TitleState
    RaiseEvent Click(m_TitleState)
End Sub

'INITIALIZE control
Private Sub UserControl_Initialize()
    
    'Initialize a master user control support class
    Set ucSupport = New pdUCSupport
    ucSupport.RegisterControl UserControl.hWnd
    
    'Request any control-specific functionality
    ucSupport.RequestExtraFunctionality True, True
    ucSupport.SpecifyRequiredKeys VK_SPACE, VK_RETURN
    ucSupport.RequestCaptionSupport
    ucSupport.SetCaptionAutomaticPainting False
    
    'In design mode, initialize a base theming class, so our paint functions don't fail
    If g_Themer Is Nothing Then Set g_Themer = New pdVisualThemes
    
    'Update the control size parameters at least once
    updateControlLayout
                
End Sub

'Set default properties
Private Sub UserControl_InitProperties()
    BackColor = vbWhite
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
        m_BackColor = .ReadProperty("BackColor", vbWhite)
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
        .WriteProperty "BackColor", m_BackColor, vbWhite
        .WriteProperty "Caption", ucSupport.GetCaptionText, ""
        .WriteProperty "FontBold", ucSupport.GetCaptionFontBold, False
        .WriteProperty "FontSize", ucSupport.GetCaptionFontSize, 10
        .WriteProperty "Value", m_TitleState, True
    End With
End Sub

'Because this control automatically forces all internal buttons to identical sizes, we have to recalculate a number
' of internal sizing metrics whenever the control size changes.
Private Sub updateControlLayout()

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
    redrawBackBuffer
            
End Sub

'External functions can call this to request a redraw.  This is helpful for live-updating theme settings, as in the Preferences dialog.
Public Sub UpdateAgainstCurrentTheme()
    
    'The support class handles most of this for us
    If g_IsProgramRunning Then ucSupport.UpdateAgainstThemeAndLanguage
    
    'If theme changes require us to redraw our control, the support class will raise additional paint events for us.
    
End Sub

'Use this function to completely redraw the back buffer from scratch.  Note that this is computationally expensive compared to just flipping the
' existing buffer to the screen, so only redraw the backbuffer if the control state has somehow changed.
Private Sub redrawBackBuffer()
    
    If g_IsProgramRunning Then
        
        'Colors used throughout this paint function are determined by several factors:
        ' 1) Control enablement (disabled controls are grayed)
        ' 2) Hover state (hovered controls glow)
        ' 3) Value (controls arrow direction)
        ' 4) The central themer (which contains default color values for all these scenarios)
        Dim textColor As Long, arrowColor As Long
        Dim ctlFillColor As Long, ctlTopLineColor As Long
        
        'For this particular control, fill color is always consistent
        ctlFillColor = g_Themer.GetThemeColor(PDTC_BACKGROUND_DEFAULT)
        
        If Me.Enabled Then
            
            'Is the mouse inside the UC?
            If ucSupport.IsMouseInside Then
                textColor = g_Themer.GetThemeColor(PDTC_ACCENT_SHADOW)
                arrowColor = g_Themer.GetThemeColor(PDTC_ACCENT_DEFAULT)
                ctlTopLineColor = g_Themer.GetThemeColor(PDTC_ACCENT_DEFAULT)
                
            'The mouse is not inside the UC
            Else
                
                'If focus was received via keyboard, change the border to reflect it
                If ucSupport.DoIHaveFocus Then
                    ctlTopLineColor = g_Themer.GetThemeColor(PDTC_ACCENT_DEFAULT)
                Else
                    ctlTopLineColor = g_Themer.GetThemeColor(PDTC_GRAY_HIGHLIGHT)
                End If
                
                'Text and arrow color is identical regardless of focus
                textColor = g_Themer.GetThemeColor(PDTC_TEXT_TITLE)
                arrowColor = g_Themer.GetThemeColor(PDTC_GRAY_DEFAULT)
            
            End If
            
        'The button is disabled
        Else
            ctlTopLineColor = g_Themer.GetThemeColor(PDTC_DISABLED)
            textColor = g_Themer.GetThemeColor(PDTC_DISABLED)
            arrowColor = g_Themer.GetThemeColor(PDTC_DISABLED)
        End If
        
        'Request the back buffer DC, and ask the support module to erase any existing rendering for us.
        Dim bufferDC As Long
        bufferDC = ucSupport.GetBackBufferDC(True)
        
        Dim bWidth As Long, bHeight As Long
        bWidth = ucSupport.GetBackBufferWidth
        bHeight = ucSupport.GetBackBufferHeight
        
        'First, we fill the button interior with the established fill color
        GDI_Plus.GDIPlusFillRectToDC bufferDC, 0, 0, bWidth - 1, bHeight - 1, ctlFillColor, 255
                    
        'Paint the caption, if any
        If ucSupport.IsCaptionActive Then
            ucSupport.SetCaptionCustomColor textColor
            ucSupport.PaintCaptionManually
        End If
            
        'Next, paint the drop-down arrow.  To simplify calculations, we first calculate the boundary rect where the arrow will be drawn.
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
        GDI_Plus.GDIPlusDrawLineToDC bufferDC, arrowPt1.x, arrowPt1.y, arrowPt2.x, arrowPt2.y, arrowColor, 255, 2, True, LineCapRound
        GDI_Plus.GDIPlusDrawLineToDC bufferDC, arrowPt2.x, arrowPt2.y, arrowPt3.x, arrowPt3.y, arrowColor, 255, 2, True, LineCapRound
        
        'Finally, frame the control.  At present, this consists of two gradient lines - one across the top, the other down the right side.
        GDI_Plus.GDIPlusDrawGradientLineToDC bufferDC, 0#, 0#, bWidth - 1, 0#, ctlFillColor, ctlTopLineColor, 255, 255, 1, True, LineCapRound
        GDI_Plus.GDIPlusDrawGradientLineToDC bufferDC, bWidth - 1, 0#, bWidth - 1, bHeight, ctlTopLineColor, ctlFillColor, 255, 255, 1, True, LineCapRound
        
    End If
    
    'Paint the final result to the screen, as relevant
    ucSupport.RequestRepaint
    
End Sub

'By design, PD prefers to not use design-time tooltips.  Apply tooltips at run-time, using this function.
' (IMPORTANT NOTE: translations are handled automatically.  Always pass the original English text!)
Public Sub AssignTooltip(ByVal newTooltip As String, Optional ByVal newTooltipTitle As String, Optional ByVal newTooltipIcon As TT_ICON_TYPE = TTI_NONE)
    ucSupport.AssignTooltip UserControl.containerHwnd, newTooltip, newTooltipTitle, newTooltipIcon
End Sub
