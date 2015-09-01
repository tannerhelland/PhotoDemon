VERSION 5.00
Begin VB.UserControl pdButton 
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
   ToolboxBitmap   =   "pdButton.ctx":0000
End
Attribute VB_Name = "pdButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Generic Button control
'Copyright 2014-2015 by Tanner Helland
'Created: 19/October/14
'Last updated: 31/August/15
'Last update: split off from pdButtonToolbox.  The two controls are similar, but this one needs to manage a caption.
'
'In a surprise to precisely no one, PhotoDemon has some unique needs when it comes to user controls - needs that
' the intrinsic VB controls can't handle.  These range from the obnoxious (lack of an "autosize" property for
' anything but labels) to the critical (no Unicode support).
'
'As such, I've created many of my own UCs for the program.  All are owner-drawn, with the goal of maintaining
' visual fidelity across the program, while also enabling key features like Unicode support.
'
'A few notes on this generic button control, specifically:
'
' 1) Captioning is (mostly) handled by the pdCaption class, so autosizing of overlong text is supported.
' 2) High DPI settings are handled automatically.
' 3) A hand cursor is automatically applied, and clicks are returned via the Click event.
' 4) Coloration is automatically handled by PD's internal theming engine.
' 5) This button cannot be used as a checkbox, e.g. it does not set a "Value" property when clicked.  It simply raises
'     a Click() event.  This is by design.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'This control really only needs one event raised - Click
Public Event Click()

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

'pdCaption manages all caption-related settings, so we don't have to.  (Note that this includes not just the caption, but related
' settings like caption font size.)
Private m_Caption As pdCaption

'Rect where the caption is rendered.  This is calculated by updateControlLayout, and it needs to be revisited if either the caption
' or button images change.
Private m_CaptionRect As RECT

'Button images.  All are optional.
Private btImage As pdDIB                'You must specify this image manually, at run-time.
Private btImageDisabled As pdDIB        'Auto-created disabled version of the image.
Private btImageHover As pdDIB           'Auto-created hover (glow) version of the image.

'(x, y) position of the button image.  This is auto-calculated by the control.
Private btImageCoords As POINTAPI

'Persistent back buffer, which we manage internally.  This allows for color management (yes, even on UI elements!)
Private m_BackBuffer As pdDIB

'If the mouse is currently INSIDE the control, this will be set to TRUE
Private m_MouseInsideUC As Boolean

'When the control receives focus via keyboard (e.g. NOT by mouse events), we draw a focus rect to help orient the user.
Private m_FocusRectActive As Boolean

'Current back color
Private m_BackColor As OLE_COLOR

'Current button state (TRUE when button is down)
Private m_ButtonStateDown As Boolean

'Additional helper for rendering themed and multiline tooltips
Private toolTipManager As pdToolTip

'If the control is currently visible, this will be set to TRUE.  This can be used to suppress redraw requests for hidden controls.
Private m_ControlIsVisible As Boolean

'BackColor is an important property for this control, as it may sit on other controls whose backcolor is not guaranteed in advance.
' So we can't rely on theming alone to determine this value.
Public Property Get BackColor() As OLE_COLOR
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal newColor As OLE_COLOR)
    m_BackColor = newColor
    redrawBackBuffer
End Property

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
    
    'If a focus rect has been drawn, remove it now
    If m_FocusRectActive Then
        m_FocusRectActive = False
        redrawBackBuffer
    End If
    
    RaiseEvent LostFocusAPI

End Sub

'A few key events are also handled
Private Sub cKeyEvents_KeyDownCustom(ByVal Shift As ShiftConstants, ByVal vkCode As Long, markEventHandled As Boolean)
        
    'When space is pressed, raise a click event.
    If (vkCode = VK_SPACE) Then

        If m_FocusRectActive And Me.Enabled Then
            m_ButtonStateDown = True
            redrawBackBuffer
            RaiseEvent Click
        End If
        
    End If

End Sub

Private Sub cKeyEvents_KeyUpCustom(ByVal Shift As ShiftConstants, ByVal vkCode As Long, markEventHandled As Boolean)

    'When space is released, redraw the button to match
    If (vkCode = VK_SPACE) Then

        If Me.Enabled Then
            m_ButtonStateDown = False
            redrawBackBuffer
        End If
        
    End If

End Sub

'Only left clicks raise Click() events
Private Sub cMouseEvents_ClickCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    
    If Me.Enabled And (Button = pdLeftButton) Then
        
        'Note that drawing flags are handled by MouseDown/Up.  Click() is only used for raising a matching Click() event.
        RaiseEvent Click
        
    End If
    
End Sub

Private Sub cMouseEvents_MouseDownCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    
    If Me.Enabled And ((Button And pdLeftButton) <> 0) Then
    
        'Ensure that a focus event has been raised, if it wasn't already
        If Not cFocusDetector.HasFocus Then cFocusDetector.setFocusManually
        
        'Set button state and redraw
        m_ButtonStateDown = True
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

Private Sub cMouseEvents_MouseUpCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal ClickEventAlsoFiring As Boolean)
    
    'If toggle mode is active, remove the button's TRUE state and redraw it
    If m_ButtonStateDown Then
        m_ButtonStateDown = False
        redrawBackBuffer
    End If
    
End Sub

'Assign a DIB to this button.  Matching disabled and hover state DIBs are automatically generated.
' Note that you can supply an existing DIB, or a resource name.  You must supply one or the other (obviously).
' No preprocessing is currently applied to DIBs loaded as a resource.
Public Sub AssignImage(Optional ByVal resName As String = "", Optional ByRef srcDIB As pdDIB, Optional ByVal scalePixelsWhenDisabled As Long = 0, Optional ByVal customGlowWhenHovered As Long = 0)
    
    'Load the requested resource DIB, as necessary
    If Len(resName) <> 0 Then loadResourceToDIB resName, srcDIB
        
    'Start by making a copy of the source DIB
    Set btImage = New pdDIB
    btImage.createFromExistingDIB srcDIB
        
    'Next, create a grayscale copy of the image for the disabled state
    Set btImageDisabled = New pdDIB
    btImageDisabled.createFromExistingDIB btImage
    GrayscaleDIB btImageDisabled, True
    If scalePixelsWhenDisabled <> 0 Then ScaleDIBRGBValues btImageDisabled, scalePixelsWhenDisabled, True
    
    'Finally, create a "glowy" hovered version of the DIB for hover state
    Set btImageHover = New pdDIB
    btImageHover.createFromExistingDIB btImage
    If customGlowWhenHovered = 0 Then
        ScaleDIBRGBValues btImageHover, UC_HOVER_BRIGHTNESS, True
    Else
        ScaleDIBRGBValues btImageHover, customGlowWhenHovered, True
    End If
    
    'Request a control size update, which will also calculate a centered position for the new image
    updateControlLayout

End Sub

'The pdWindowPaint class raises this event when the control needs to be redrawn.  The passed coordinates contain the
' rect returned by GetUpdateRect (but with right/bottom measurements pre-converted to width/height).
Private Sub cPainter_PaintWindow(ByVal winLeft As Long, ByVal winTop As Long, ByVal winWidth As Long, ByVal winHeight As Long)

    'Flip the relevant chunk of the buffer to the screen
    BitBlt UserControl.hDC, winLeft, winTop, winWidth, winHeight, m_BackBuffer.getDIBDC, winLeft, winTop, vbSrcCopy
    
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
        cKeyEvents.createKeyboardTracker "Toolbox button UC", Me.hWnd, VK_SPACE
        
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
    
    'Update the control size parameters at least once
    updateControlLayout
                
End Sub

'Set default properties
Private Sub UserControl_InitProperties()
    BackColor = vbWhite
    Caption = ""
    FontSize = 10
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

'At run-time, painting is handled by PD's pdWindowPainter class.  In the IDE, however, we must rely on VB's internal paint event.
Private Sub UserControl_Paint()
    
    'Provide minimal painting within the designer
    If Not g_IsProgramRunning Then redrawBackBuffer
    
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    With PropBag
        m_BackColor = .ReadProperty("BackColor", vbWhite)
        Caption = .ReadProperty("Caption", "")
        FontSize = .ReadProperty("FontSize", 10)
    End With

End Sub

'The control dynamically resizes each button to match the dimensions of their relative captions.
Private Sub UserControl_Resize()
    updateControlLayout
End Sub

'Because this control automatically forces all internal buttons to identical sizes, we have to recalculate a number
' of internal sizing metrics whenever the control size changes.
Private Sub updateControlLayout()
    
    'Reset the back buffer
    Set m_BackBuffer = New pdDIB
    m_BackBuffer.createBlank UserControl.ScaleWidth, UserControl.ScaleHeight, m_BackColor
    
    'Next, we need to determine the positioning of the caption and/or image.  Both (or neither) of these may be missing, so handling
    ' can get a little complicated.
    
    'Start with the caption
    If m_Caption.isCaptionActive Then
        
        'We need to find the available area for the caption.  The caption class will automatically fit any translated text inside
        ' this rect.
        Const hTextPadding As Long = 8&, vTextPadding As Long = 4&
        
        'The y-positioning of the caption is always constant
        m_CaptionRect.Top = vTextPadding
        m_CaptionRect.Bottom = m_BackBuffer.getDIBHeight - vTextPadding
        
        'Similarly, the right bound doesn't change either
        m_CaptionRect.Right = m_BackBuffer.getDIBWidth - hTextPadding
        
        'If a button image is active, forcibly calculate its position first.  Its position is hard-coded.
        If Not (btImage Is Nothing) Then
        
            Const leftButtonPadding As Long = 12&
            btImageCoords.x = fixDPI(leftButtonPadding)
            btImageCoords.y = (m_BackBuffer.getDIBHeight - btImage.getDIBHeight) \ 2
            
            'Use the button's position to calculate an x-coord for the caption, too
            m_CaptionRect.Left = btImageCoords.x + btImage.getDIBWidth + hTextPadding
                    
        Else
            m_CaptionRect.Left = hTextPadding
        End If
        
        'Notify the caption renderer of this new caption position, which it will use to automatically adjust its font, as necessary
        m_Caption.setControlWidth m_CaptionRect.Right - m_CaptionRect.Left
    
    'If there's no caption, center the button image on the control
    Else
        
        'Determine positioning of the button image, if any
        If Not (btImage Is Nothing) Then
            btImageCoords.x = (m_BackBuffer.getDIBWidth - btImage.getDIBWidth) \ 2
            btImageCoords.y = (m_BackBuffer.getDIBHeight - btImage.getDIBHeight) \ 2
        End If
        
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
        .WriteProperty "FontSizeCaption", m_Caption.getFontSize, 10
    End With
    
End Sub

'External functions can call this to request a redraw.  This is helpful for live-updating theme settings, as in the Preferences dialog.
Public Sub updateAgainstCurrentTheme()
    
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
    ' 1) Control enablement (disabled buttons are grayed)
    ' 2) Hover state (hovered buttons glow)
    ' 3) Value (pressed buttons have a different appearance, obviously)
    ' 4) The central themer (which contains default color values for all these scenarios)
    Dim btnColorBorder As Long, btnColorFill As Long
    Dim textColor As Long
        
    If Me.Enabled Then
    
        'Is the button pressed?
        If m_ButtonStateDown Then
            btnColorFill = g_Themer.getThemeColor(PDTC_ACCENT_HIGHLIGHT)
            btnColorBorder = g_Themer.getThemeColor(PDTC_ACCENT_SHADOW)
            textColor = g_Themer.getThemeColor(PDTC_TEXT_INVERT)
            
        'The button is not pressed
        Else
        
            'Is the mouse inside the UC?
            If m_MouseInsideUC Then
                btnColorFill = g_Themer.getThemeColor(PDTC_ACCENT_ULTRALIGHT)
                btnColorBorder = g_Themer.getThemeColor(PDTC_ACCENT_DEFAULT)
                textColor = g_Themer.getThemeColor(PDTC_TEXT_TITLE)
            
            'The mouse is not inside the UC
            Else
                btnColorFill = m_BackColor
                btnColorBorder = g_Themer.getThemeColor(PDTC_GRAY_DEFAULT)
                textColor = g_Themer.getThemeColor(PDTC_TEXT_TITLE)
            
            End If
            
        End If
        
    'The button is disabled
    Else
    
        btnColorFill = m_BackColor
        btnColorBorder = g_Themer.getThemeColor(PDTC_DISABLED)
        textColor = g_Themer.getThemeColor(PDTC_DISABLED)
        
    End If
    
    'First, we fill the button interior with the established fill color
    GDI_Plus.GDIPlusFillDIBRect m_BackBuffer, 0, 0, m_BackBuffer.getDIBWidth - 1, m_BackBuffer.getDIBHeight - 1, btnColorFill, 255
    
    'A border is always drawn around the control; its size varies by hover state.  (This is standard Win 10 behavior.)
    Dim borderWidth As Single
    If m_MouseInsideUC Then borderWidth = 3 Else borderWidth = 1
    GDI_Plus.GDIPlusDrawRectOutlineToDC m_BackBuffer.getDIBDC, 0, 0, m_BackBuffer.getDIBWidth - 1, m_BackBuffer.getDIBHeight - 1, btnColorBorder, 255, borderWidth
    
    'Paint the image, if any
    If Not (btImage Is Nothing) Then
        
        If Me.Enabled Then
            
            If m_MouseInsideUC Then
                btImageHover.alphaBlendToDC m_BackBuffer.getDIBDC, 255, btImageCoords.x, btImageCoords.y
            Else
                btImage.alphaBlendToDC m_BackBuffer.getDIBDC, 255, btImageCoords.x, btImageCoords.y
            End If
            
        Else
            btImageDisabled.alphaBlendToDC m_BackBuffer.getDIBDC, 255, btImageCoords.x, btImageCoords.y
        End If
        
    End If
    
    'Paint the caption, if any
    If m_Caption.isCaptionActive Then
        m_Caption.setCaptionColor textColor
        m_Caption.drawCaptionCentered m_BackBuffer.getDIBDC, m_CaptionRect
    End If
            
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

'The color selector dialog has the unique need of capturing colors from anywhere on the screen, using a custom hook solution.  For it to work,
' the pdInputMouse class inside this button control must forcibly release its capture.
Public Sub overrideMouseCapture(ByVal NewState As Boolean)
    cMouseEvents.setCaptureOverride NewState
    cMouseEvents.setCursorOverrideState NewState
End Sub

'Due to complex interactions between user controls and PD's translation engine, tooltips require this dedicated function.
' (IMPORTANT NOTE: the tooltip class will handle translations automatically.  Always pass the original English text!)
Public Sub assignTooltip(ByVal newTooltip As String, Optional ByVal newTooltipTitle As String, Optional ByVal newTooltipIcon As TT_ICON_TYPE = TTI_NONE)
    toolTipManager.setTooltip Me.hWnd, Me.containerHwnd, newTooltip, newTooltipTitle, newTooltipIcon
End Sub

