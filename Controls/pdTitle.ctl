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
'Last updated: 27/September/15
'Last update: split off from pdButton.  The two controls are similar code-wise, but their visible UI is quite different.
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

'API technique for drawing a focus rectangle; used only for designer mode (see the Paint method for details)
Private Declare Function DrawFocusRect Lib "user32" (ByVal hDC As Long, ByRef lpRect As RECT) As Long

'Mouse and keyboard input handlers
Private WithEvents cMouseEvents As pdInputMouse
Attribute cMouseEvents.VB_VarHelpID = -1
Private WithEvents cKeyEvents As pdInputKeyboard
Attribute cKeyEvents.VB_VarHelpID = -1

'Flicker-free window painter
Private WithEvents cPainter As pdWindowPainter
Attribute cPainter.VB_VarHelpID = -1

'Reliable focus detection requires a specialized subclasser
Private WithEvents cFocusDetector As pdFocusDetector
Attribute cFocusDetector.VB_VarHelpID = -1
Public Event GotFocusAPI()
Public Event LostFocusAPI()

'pdCaption manages all caption-related settings, so we don't have to.  (Note that this includes not just the caption,
' but related settings like caption font size.)
Private m_Caption As pdCaption

'Rect where the caption is rendered.  This is calculated by updateControlLayout, and it needs to be revisited if the
' caption changes, or the control size changes.
Private m_CaptionRect As RECT

'Persistent back buffer, which we manage internally.  This allows for color management (yes, even on UI elements!)
Private m_BackBuffer As pdDIB

'If the user resizes the control, the control's back buffer needs to be redrawn.  If we resize the control as part of an internal
' AutoSize calculation, however, we will already be in the midst of resizing the backbuffer - so we override the behavior
' of the UserControl_Resize event, using this variable.
Private m_InternalResizeState As Boolean

'If the mouse is currently INSIDE the control, this will be set to TRUE
Private m_MouseInsideUC As Boolean

'When the control receives focus via keyboard (e.g. NOT by mouse events), we draw a focus rect to help orient the user.
Private m_FocusRectActive As Boolean

'Current back color
Private m_BackColor As OLE_COLOR

'Current title state (TRUE when arrow is pointing down, e.g. the associated container is "open")
Private m_TitleState As Boolean

'Additional helper for rendering themed and multiline tooltips
Private toolTipManager As pdToolTip

'If the control is currently visible, this will be set to TRUE.  This can be used to suppress redraw requests for hidden controls.
Private m_ControlIsVisible As Boolean

'Caption is handled just like the common control label's caption property.  It is valid at design-time, and any translation,
' if present, will not be processed until run-time.
' IMPORTANT NOTE: only the ENGLISH caption is returned.  I don't have a reason for returning a translated caption (if any),
'                  but I can revisit in the future if it ever becomes relevant.
Public Property Get Caption() As String
Attribute Caption.VB_UserMemId = -518
    Caption = m_Caption.getCaptionEn
End Property

Public Property Let Caption(ByRef newCaption As String)
    
    If m_Caption.setCaption(newCaption) And (m_ControlIsVisible Or (Not g_IsProgramRunning)) Then updateControlLayout
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
    FontBold = m_Caption.getFontBold
End Property

Public Property Let FontBold(ByVal newValue As Boolean)
    If m_Caption.setFontBold(newValue) And (m_ControlIsVisible Or (Not g_IsProgramRunning)) Then updateControlLayout
    PropertyChanged "FontBold"
End Property

Public Property Get FontSize() As Single
    FontSize = m_Caption.getFontSize
End Property

Public Property Let FontSize(ByVal newSize As Single)
    If m_Caption.setFontSize(newSize) And (m_ControlIsVisible Or (Not g_IsProgramRunning)) Then updateControlLayout
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

'When the control receives focus, if the focus isn't received via mouse click, display a focus rect
Private Sub cFocusDetector_GotFocusReliable()
    
    'If the mouse is *not* over the user control, assume focus was set via keyboard
    If Not m_MouseInsideUC Then
        m_FocusRectActive = True
        redrawBackBuffer
    End If
    
    RaiseEvent GotFocusAPI
    
End Sub

'When the control loses focus, erase any focus rects it may have active
Private Sub cFocusDetector_LostFocusReliable()
    makeLostFocusUIChanges
    RaiseEvent LostFocusAPI
End Sub

Private Sub makeLostFocusUIChanges()
    
    'If a focus rect has been drawn, remove it now
    If m_FocusRectActive Or m_MouseInsideUC Then
        m_FocusRectActive = False
        m_MouseInsideUC = False
        redrawBackBuffer
    End If
    
End Sub

'A few key events are also handled
Private Sub cKeyEvents_KeyUpCustom(ByVal Shift As ShiftConstants, ByVal vkCode As Long, markEventHandled As Boolean)

    'When space is released, redraw the button to match
    If (vkCode = VK_SPACE) Or (vkCode = VK_RETURN) Then

        If m_FocusRectActive And Me.Enabled Then
            m_TitleState = Not m_TitleState
            redrawBackBuffer
            RaiseEvent Click(m_TitleState)
        End If
        
    End If

End Sub

'Only left clicks raise Click() events
Private Sub cMouseEvents_MouseDownCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    
    If Me.Enabled And ((Button And pdLeftButton) <> 0) Then
    
        'Ensure that a focus event has been raised, if it wasn't already
        If Not cFocusDetector.HasFocus Then cFocusDetector.setFocusManually
        
        'Set button state and redraw
        m_TitleState = Not m_TitleState
        
        'Note that drawing flags are handled by MouseDown/Up.  Click() is only used for raising a matching Click() event.
        RaiseEvent Click(m_TitleState)
        redrawBackBuffer
        
    End If
    
End Sub

Private Sub cMouseEvents_MouseEnter(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    m_MouseInsideUC = True
    cMouseEvents.setSystemCursor IDC_HAND
    redrawBackBuffer
End Sub

'When the mouse leaves the UC, we must repaint the button (as it's no longer hovered)
Private Sub cMouseEvents_MouseLeave(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    
    If m_MouseInsideUC Then
        m_MouseInsideUC = False
        redrawBackBuffer
    End If
    
    'Reset the cursor
    cMouseEvents.setSystemCursor IDC_ARROW
    
End Sub

'When the mouse enters the button, we must initiate a repaint (to reflect its hovered state)
Private Sub cMouseEvents_MouseMoveCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    
    'Repaint the control as necessary
    If Not m_MouseInsideUC Then
        m_MouseInsideUC = True
        redrawBackBuffer
    End If
    
End Sub

'The pdWindowPaint class raises this event when the control needs to be redrawn.  The passed coordinates contain the
' rect returned by GetUpdateRect (but with right/bottom measurements pre-converted to width/height).
Private Sub cPainter_PaintWindow(ByVal winLeft As Long, ByVal winTop As Long, ByVal winWidth As Long, ByVal winHeight As Long)

    'Flip the relevant chunk of the buffer to the screen
    BitBlt UserControl.hDC, winLeft, winTop, winWidth, winHeight, m_BackBuffer.getDIBDC, winLeft, winTop, vbSrcCopy
    
End Sub

Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
    m_TitleState = Not m_TitleState
    RaiseEvent Click(m_TitleState)
End Sub

Private Sub UserControl_Hide()
    m_ControlIsVisible = False
End Sub

'INITIALIZE control
Private Sub UserControl_Initialize()
    
    'When not in design mode, initialize trackers for input events
    If g_IsProgramRunning Then
    
        Set cMouseEvents = New pdInputMouse
        cMouseEvents.addInputTracker Me.hWnd, True, True, , True
        cMouseEvents.setSystemCursor IDC_HAND
        
        Set cKeyEvents = New pdInputKeyboard
        cKeyEvents.createKeyboardTracker "pdTitle", Me.hWnd, VK_SPACE
        
        'Also start a flicker-free window painter
        Set cPainter = New pdWindowPainter
        cPainter.startPainter Me.hWnd
        
        'Also start a focus detector
        Set cFocusDetector = New pdFocusDetector
        cFocusDetector.startFocusTracking Me.hWnd
        
        'Create a tooltip engine
        Set toolTipManager = New pdToolTip
        
    'In design mode, initialize a base theming class, so our paint function doesn't fail
    Else
        If g_Themer Is Nothing Then Set g_Themer = New pdVisualThemes
    End If
    
    m_MouseInsideUC = False
    m_FocusRectActive = False
    
    'Prep the caption object
    Set m_Caption = New pdCaption
    m_Caption.setWordWrapSupport False
    
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

'Because VB is very dumb about focus handling, it is sometimes necessary for external functions to notify of focus loss.
' (TODO: revisit in light of the new Got/LostFocusAPI functions)
Public Sub notifyFocusLost()

    'If a focus rect has been drawn, remove it now
    If m_FocusRectActive Or m_MouseInsideUC Then
        m_FocusRectActive = False
        m_MouseInsideUC = False
        redrawBackBuffer
    End If

End Sub

Private Sub UserControl_LostFocus()
    makeLostFocusUIChanges
End Sub

'At run-time, painting is handled by PD's pdWindowPainter class.  In the IDE, however, we must rely on VB's internal paint event.
Private Sub UserControl_Paint()
    
    'Provide minimal painting within the designer
    If Not g_IsProgramRunning Then redrawBackBuffer
    
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

'The control dynamically resizes each button to match the dimensions of their relative captions.
Private Sub UserControl_Resize()
    If (Not m_InternalResizeState) Then updateControlLayout
End Sub

'Because this control automatically forces all internal buttons to identical sizes, we have to recalculate a number
' of internal sizing metrics whenever the control size changes.
Private Sub updateControlLayout()
    
    'First, make sure the back buffer exists and mirrors the current control size
    If m_BackBuffer Is Nothing Then Set m_BackBuffer = New pdDIB
    If (m_BackBuffer.getDIBWidth <> UserControl.ScaleWidth) Or (m_BackBuffer.getDIBHeight <> UserControl.ScaleHeight) Then
        m_BackBuffer.createBlank UserControl.ScaleWidth, UserControl.ScaleHeight, 24, m_BackColor
    Else
        GDI_Plus.GDIPlusFillDIBRect m_BackBuffer, 0, 0, m_BackBuffer.getDIBWidth, m_BackBuffer.getDIBHeight, m_BackColor
    End If
    
    Const hTextPadding As Long = 2&, vTextPadding As Long = 2&
    
    'Next, we need to determine the size of the caption.  The caption height determines control height, so if the current control
    ' size does not match that value, we want to immediately resize the control to match.
    If m_Caption.isCaptionActive Then
        
        If m_Caption.getCaptionHeight + FixDPI(vTextPadding) * 2 <> UserControl.ScaleHeight Then
            
            m_InternalResizeState = True
                
            'Resize the user control.  For inexplicable reasons, setting the .Width and .Height properties works for .Width,
            ' but not for .Height (aaarrrggghhh).  Fortunately, we can work around this rather easily by using MoveWindow and
            ' forcing a repaint at run-time, and reverting to the problematic internal methods only in the IDE.
            If g_IsProgramRunning Then
                MoveWindow Me.hWnd, UserControl.Extender.Left, UserControl.Extender.Top, UserControl.ScaleWidth, m_Caption.getCaptionHeight + FixDPI(vTextPadding) * 2, 1
            Else
                UserControl.Size PXToTwipsX(UserControl.ScaleWidth), PXToTwipsY(m_Caption.getCaptionHeight + 2)
            End If
            
            'Recreate the backbuffer to match
            m_BackBuffer.createBlank UserControl.ScaleWidth, UserControl.ScaleHeight, 24
            
            'Restore normal resize behavior
            m_InternalResizeState = False
            
        End If
        
        'The control and backbuffer are now guaranteed to be the proper size.
        
        'For caption rendering purposes, we need to determine a target rectangle for the caption itself.  The pdCaption class will
        ' automatically fit the caption within this area, regardless of the currently selected font size.
        With m_CaptionRect
            .Left = FixDPI(hTextPadding)
            .Top = FixDPI(vTextPadding)
            .Bottom = m_BackBuffer.getDIBHeight - FixDPI(vTextPadding)
            
            'The right measurement is the only complicated one, as it requires padding so we have room to render the drop-down arrow.
            .Right = m_BackBuffer.getDIBWidth - FixDPI(hTextPadding) * 2 - m_BackBuffer.getDIBHeight
        End With
        
        'Notify the caption renderer of this new caption position, which it will use to automatically adjust its font, as necessary
        m_Caption.setControlSize m_CaptionRect.Right - m_CaptionRect.Left, m_CaptionRect.Bottom - m_CaptionRect.Top
        
    End If
        
    'No other special preparation is required for this control, so proceed with recreating the back buffer
    redrawBackBuffer
            
End Sub

Private Sub UserControl_Show()
    m_ControlIsVisible = True
    updateControlLayout
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    'Store all associated properties
    With PropBag
        .WriteProperty "BackColor", m_BackColor, vbWhite
        .WriteProperty "Caption", m_Caption.getCaptionEn, ""
        .WriteProperty "FontBold", m_Caption.getFontBold, False
        .WriteProperty "FontSize", m_Caption.getFontSize, 10
        .WriteProperty "Value", m_TitleState, True
    End With
    
End Sub

'External functions can call this to request a redraw.  This is helpful for live-updating theme settings, as in the Preferences dialog.
Public Sub UpdateAgainstCurrentTheme()
    
    'Make sure captions and tooltips are valid
    m_Caption.UpdateAgainstCurrentTheme
    
    'Redraw the control, which will also cause a resync against any theme changes
    updateControlLayout
    
End Sub

'Use this function to completely redraw the back buffer from scratch.  Note that this is computationally expensive compared to just flipping the
' existing buffer to the screen, so only redraw the backbuffer if the control state has somehow changed.
Private Sub redrawBackBuffer()
    
    'Start by erasing the back buffer
    If g_IsProgramRunning Then
        GDI_Plus.GDIPlusFillDIBRect m_BackBuffer, 0, 0, m_BackBuffer.getDIBWidth, m_BackBuffer.getDIBHeight, m_BackColor, 255
    Else
        m_BackBuffer.createBlank UserControl.ScaleWidth, UserControl.ScaleHeight, 24, RGB(255, 255, 255)
    End If
    
    'Colors used throughout this paint function are determined by several factors:
    ' 1) Control enablement (disabled controls are grayed)
    ' 2) Hover state (hovered controls glow)
    ' 3) Value (controls arrow direction)
    ' 4) The central themer (which contains default color values for all these scenarios)
    Dim textColor As Long, arrowColor As Long
    Dim ctlBorderColor As Long, ctlFillColor As Long, ctlTopLineColor As Long
    
    'For this particular control, fill color is always consistent
    ctlFillColor = g_Themer.getThemeColor(PDTC_BACKGROUND_DEFAULT)
    
    If Me.Enabled Then
        
        'Is the mouse inside the UC?
        If m_MouseInsideUC Then
            ctlBorderColor = g_Themer.getThemeColor(PDTC_BACKGROUND_DEFAULT)
            textColor = g_Themer.getThemeColor(PDTC_ACCENT_SHADOW)
            arrowColor = g_Themer.getThemeColor(PDTC_ACCENT_DEFAULT)
            ctlTopLineColor = g_Themer.getThemeColor(PDTC_ACCENT_DEFAULT)
            
        'The mouse is not inside the UC
        Else
            
            'If focus was received via keyboard, change the border to reflect it
            If m_FocusRectActive Then
                ctlBorderColor = g_Themer.getThemeColor(PDTC_ACCENT_HIGHLIGHT)
                ctlTopLineColor = g_Themer.getThemeColor(PDTC_ACCENT_DEFAULT)
            Else
                ctlBorderColor = g_Themer.getThemeColor(PDTC_BACKGROUND_DEFAULT)
                ctlTopLineColor = g_Themer.getThemeColor(PDTC_GRAY_HIGHLIGHT)
            End If
            
            'Text and arrow color is identical regardless of focus
            textColor = g_Themer.getThemeColor(PDTC_TEXT_TITLE)
            arrowColor = g_Themer.getThemeColor(PDTC_GRAY_DEFAULT)
        
        End If
        
    'The button is disabled
    Else
    
        ctlBorderColor = g_Themer.getThemeColor(PDTC_BACKGROUND_DEFAULT)
        ctlTopLineColor = g_Themer.getThemeColor(PDTC_DISABLED)
        textColor = g_Themer.getThemeColor(PDTC_DISABLED)
        arrowColor = g_Themer.getThemeColor(PDTC_DISABLED)
        
    End If
    
    'First, we fill the button interior with the established fill color
    GDI_Plus.GDIPlusFillDIBRect m_BackBuffer, 0, 0, m_BackBuffer.getDIBWidth - 1, m_BackBuffer.getDIBHeight - 1, ctlFillColor, 255
    
    'A border is always drawn around the control, but at present, its color always matches the background color unless focus was
    ' specifically received via keyboard.
    GDI_Plus.GDIPlusDrawRectOutlineToDC m_BackBuffer.getDIBDC, 0, 0, m_BackBuffer.getDIBWidth - 1, m_BackBuffer.getDIBHeight - 1, ctlBorderColor, 255, 1#
        
    'Paint the caption, if any
    If m_Caption.isCaptionActive Then
        m_Caption.setCaptionColor textColor
        m_Caption.drawCaption m_BackBuffer.getDIBDC, m_CaptionRect.Left, m_CaptionRect.Top
    End If
        
    'Next, paint the drop-down arrow.  To simplify calculations, we first calculate the boundary rect where the arrow will be drawn.
    Dim arrowRect As RECTF
    arrowRect.Left = m_BackBuffer.getDIBWidth - m_BackBuffer.getDIBHeight - FixDPI(2)
    arrowRect.Top = 1
    arrowRect.Height = m_BackBuffer.getDIBHeight - 2
    arrowRect.Width = m_BackBuffer.getDIBHeight - 2
    
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
    If m_MouseInsideUC Then arrowWidth = 2 Else arrowWidth = 1
    GDI_Plus.GDIPlusDrawLineToDC m_BackBuffer.getDIBDC, arrowPt1.x, arrowPt1.y, arrowPt2.x, arrowPt2.y, arrowColor, 255, 2, True, LineCapRound
    GDI_Plus.GDIPlusDrawLineToDC m_BackBuffer.getDIBDC, arrowPt2.x, arrowPt2.y, arrowPt3.x, arrowPt3.y, arrowColor, 255, 2, True, LineCapRound
    
    'Finally, frame the control.  At present, this consists of two gradient lines - one across the top, the other down the right side.
    GDI_Plus.GDIPlusDrawGradientLineToDC m_BackBuffer.getDIBDC, 0#, 0#, m_BackBuffer.getDIBWidth - 1, 0#, ctlFillColor, ctlTopLineColor, 255, 255, 1, True, LineCapRound
    GDI_Plus.GDIPlusDrawGradientLineToDC m_BackBuffer.getDIBDC, m_BackBuffer.getDIBWidth - 1, 0#, m_BackBuffer.getDIBWidth - 1, m_BackBuffer.getDIBHeight, ctlTopLineColor, ctlFillColor, 255, 255, 1, True, LineCapRound
    
    'In the designer, draw a focus rect around the control; this is minimal feedback required for positioning
    If Not g_IsProgramRunning Then
        
        Dim tmpRect As RECT
        With tmpRect
            .Left = 0
            .Top = 0
            .Right = m_BackBuffer.getDIBWidth
            .Bottom = m_BackBuffer.getDIBHeight
        End With
        
        DrawFocusRect m_BackBuffer.getDIBDC, tmpRect

    End If
    
    'Paint the buffer to the screen
    If g_IsProgramRunning Then cPainter.requestRepaint Else BitBlt UserControl.hDC, 0, 0, m_BackBuffer.getDIBWidth, m_BackBuffer.getDIBHeight, m_BackBuffer.getDIBDC, 0, 0, vbSrcCopy

End Sub

'Due to complex interactions between user controls and PD's translation engine, tooltips require this dedicated function.
' (IMPORTANT NOTE: the tooltip class will handle translations automatically.  Always pass the original English text!)
Public Sub AssignTooltip(ByVal newTooltip As String, Optional ByVal newTooltipTitle As String, Optional ByVal newTooltipIcon As TT_ICON_TYPE = TTI_NONE)
    toolTipManager.setTooltip Me.hWnd, Me.containerHwnd, newTooltip, newTooltipTitle, newTooltipIcon
End Sub

