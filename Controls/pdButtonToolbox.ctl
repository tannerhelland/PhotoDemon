VERSION 5.00
Begin VB.UserControl pdButtonToolbox 
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
   ToolboxBitmap   =   "pdButtonToolbox.ctx":0000
End
Attribute VB_Name = "pdButtonToolbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Toolbox Button control
'Copyright 2014-2015 by Tanner Helland
'Created: 19/October/14
'Last updated: 19/October/14
'Last update: initial build, using the button strip's code as a starting point
'
'In a surprise to precisely no one, PhotoDemon has some unique needs when it comes to user controls - needs that
' the intrinsic VB controls can't handle.  These range from the obnoxious (lack of an "autosize" property for
' anything but labels) to the critical (no Unicode support).
'
'As such, I've created many of my own UCs for the program.  All are owner-drawn, with the goal of maintaining
' visual fidelity across the program, while also enabling key features like Unicode support.
'
'A few notes on this toolbox button control, specifically:
'
' 1) Why make a separate control for toolbox buttons?  I could add a style property to the regular PD button, but I don't
'     like the complications that introduces.  "Do one thing and do it well" is the idea with PD user controls.
' 2) High DPI settings are handled automatically.
' 3) A hand cursor is automatically applied, and clicks are returned via the Click event.
' 4) Coloration is automatically handled by PD's internal theming engine.
' 5) This button does not support text, by design.  It is image-only.
' 6) This button does not automatically set its Value property when clicked.  It simply raises a Click() event.  This is
'     by design to make it easier to toggle state in the toolbox maintenance code.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'This control really only needs one event raised - Click
Public Event Click()

'API technique for drawing a focus rectangle; used only for designer mode (see the Paint method for details)
Private Declare Function DrawFocusRect Lib "user32" (ByVal hDC As Long, lpRect As RECT) As Long

'Mouse and keyboard input handlers
Private WithEvents cMouseEvents As pdInputMouse
Attribute cMouseEvents.VB_VarHelpID = -1
Private WithEvents cKeyEvents As pdInputKeyboard
Attribute cKeyEvents.VB_VarHelpID = -1

'Flicker-free window painter
Private WithEvents cPainter As pdWindowPainter
Attribute cPainter.VB_VarHelpID = -1

'Current button state
Private m_ButtonState As Boolean

'Button images.  (Since this control doesn't support text, you'd better make use of these!)
Private btImage As pdDIB                'You must specify this image manually, at run-time.
Private btImageDisabled As pdDIB        'Auto-created disabled version of the image.
Private btImageHover As pdDIB           'Auto-created hover (glow) version of the image.

'(x, y) position of the button image.  This is auto-calculated by the control.
Private btImageCoords As POINTAPI

'Persistent back buffer, which we manage internally
Private m_BackBuffer As pdDIB

'If the mouse is currently INSIDE the control, this will be set to TRUE
Private m_MouseInsideUC As Boolean

'When the option button receives focus via keyboard (e.g. NOT by mouse events), we draw a focus rect to help orient the user.
Private m_FocusRectActive As Boolean

'Current back color
Private m_BackColor As OLE_COLOR

'AutoToggle mode allows the button to operate as a normal button (e.g. no persistent value)
Private m_AutoToggle As Boolean

'Additional helpers for rendering themed and multiline tooltips
Private m_ToolTip As clsToolTip
Private m_ToolString As String

'This toolbox button control is designed to be used in a "radio button"-like system, where buttons exist in a group, and the
' pressing of one results in the unpressing of any others.  For the rare circumstances where this behavior is undesirable
' (e.g. the pdCanvas status bar, where some instances of this control serve as actual buttons), the AutoToggle property can
' be set to TRUE.  This will cause the button to operate as a normal command button.
Public Property Get AutoToggle() As Boolean
    AutoToggle = m_AutoToggle
End Property

Public Property Let AutoToggle(ByVal newToggle As Boolean)
    m_AutoToggle = newToggle
End Property

'BackColor is an important property for this control, as it may sit on other controls whose backcolor is not guaranteed in advance.
' So we can't rely on theming alone to determine this value.
Public Property Get BackColor() As OLE_COLOR
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal newColor As OLE_COLOR)
    m_BackColor = newColor
    redrawBackBuffer
End Property

'The Enabled property is a bit unique; see http://msdn.microsoft.com/en-us/library/aa261357%28v=vs.60%29.aspx
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_UserMemId = -514
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal NewValue As Boolean)
    
    UserControl.Enabled = NewValue
    PropertyChanged "Enabled"
    
    'Redraw the control
    redrawBackBuffer
    
End Property

'A few key events are also handled
Private Sub cKeyEvents_KeyDownCustom(ByVal Shift As ShiftConstants, ByVal vkCode As Long, markEventHandled As Boolean)
        
    'If space is pressed, and our value is not true, raise a click event.
    If (vkCode = VK_SPACE) Then

        If m_FocusRectActive And Me.Enabled And (Not m_ButtonState) Then
            m_ButtonState = True
            redrawBackBuffer
            RaiseEvent Click
        End If
        
    End If

End Sub

Private Sub cKeyEvents_KeyUpCustom(ByVal Shift As ShiftConstants, ByVal vkCode As Long, markEventHandled As Boolean)

    'If space was pressed, and AutoToggle is active, remove the button state and redraw it
    If (vkCode = VK_SPACE) Then

        If Me.Enabled And m_ButtonState Then
            m_ButtonState = False
            redrawBackBuffer
        End If
        
    End If

End Sub

'To improve responsiveness, MouseDown is used instead of Click
Private Sub cMouseEvents_MouseDownCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)

    If Me.Enabled And (Not m_ButtonState) Then
        m_ButtonState = True
        redrawBackBuffer
        RaiseEvent Click
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
    If m_AutoToggle And m_ButtonState Then
        m_ButtonState = False
        redrawBackBuffer
    End If
    
End Sub

'hWnds aren't exposed by default
Public Property Get hWnd() As Long
Attribute hWnd.VB_UserMemId = -515
    hWnd = UserControl.hWnd
End Property

'Container hWnd must be exposed for external tooltip handling
Public Property Get containerHwnd() As Long
    containerHwnd = UserControl.containerHwnd
End Property

'The most relevant part of this control is this Value property, which is important since this button operates as a toggle.
Public Property Get Value() As Boolean
    Value = m_ButtonState
End Property

Public Property Let Value(ByVal NewValue As Boolean)
    
    'Update our internal value tracker, but only if autotoggle is not active.  (Autotoggle causes the button to behave like
    ' a normal button, so there's no concept of a persistent "value".)
    If (m_ButtonState <> NewValue) And (Not m_AutoToggle) Then
    
        m_ButtonState = NewValue
        
        'Redraw the control to match the new state
        redrawBackBuffer
        
        'Note that we don't raise a Click event here.  This is by design.  The toolbox handles all toggle code for these buttons,
        ' and it's more efficient to let it handle this, as it already has a detailed notion of things like program state, which
        ' affects whether buttons are clickable, etc.
        
        'As such, the Click event is not raised for Value changes alone - only for actions initiated by a user.
        
    End If
    
End Property

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
    updateControlSize

End Sub

'The pdWindowPaint class raises this event when the control needs to be redrawn.  The passed coordinates contain the
' rect returned by GetUpdateRect (but with right/bottom measurements pre-converted to width/height).
Private Sub cPainter_PaintWindow(ByVal winLeft As Long, ByVal winTop As Long, ByVal winWidth As Long, ByVal winHeight As Long)

    'Flip the relevant chunk of the buffer to the screen
    BitBlt UserControl.hDC, winLeft, winTop, winWidth, winHeight, m_BackBuffer.getDIBDC, winLeft, winTop, vbSrcCopy
    
End Sub

'When the control receives focus, if the focus isn't received via mouse click, display a focus rect
Private Sub UserControl_GotFocus()

    'If the mouse is *not* over the user control, assume focus was set via keyboard
    If Not m_MouseInsideUC Then
        m_FocusRectActive = True
        redrawBackBuffer
    End If

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
        
    'In design mode, initialize a base theming class, so our paint function doesn't fail
    Else
        Set g_Themer = New pdVisualThemes
    End If
    
    m_MouseInsideUC = False
    m_FocusRectActive = False
        
    'Update the control size parameters at least once
    updateControlSize
                
End Sub

'Set default properties
Private Sub UserControl_InitProperties()
    
    Value = False
    BackColor = vbWhite
    AutoToggle = False
    
End Sub

'When the control loses focus, erase any focus rects it may have active
Private Sub UserControl_LostFocus()

    'If a focus rect has been drawn, remove it now
    If m_FocusRectActive Then
        m_FocusRectActive = False
        redrawBackBuffer
    End If

End Sub

'Because VB is very dumb about focus handling, it is sometimes necessary for external functions to notify of focus loss.
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
        m_AutoToggle = .ReadProperty("AutoToggle", False)
    End With

End Sub

'The control dynamically resizes each button to match the dimensions of their relative captions.
Private Sub UserControl_Resize()
    updateControlSize
End Sub

Private Sub UserControl_Show()

    'When the control is first made visible, remove the control's tooltip property and reassign it to the checkbox
    ' using a custom solution (which allows for linebreaks and theming).  Note that this has the ugly side-effect of
    ' permanently erasing the extender's tooltip, so FOR THIS CONTROL, TOOLTIPS MUST BE SET AT RUN-TIME!
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

'Because this control automatically forces all internal buttons to identical sizes, we have to recalculate a number
' of internal sizing metrics whenever the control size changes.
Private Sub updateControlSize()
    
    'Reset the back buffer
    Set m_BackBuffer = New pdDIB
    m_BackBuffer.createBlank UserControl.ScaleWidth, UserControl.ScaleHeight, m_BackColor
    
    'Determine positioning of the button image, if any
    If Not (btImage Is Nothing) Then
        btImageCoords.x = (m_BackBuffer.getDIBWidth - btImage.getDIBWidth) \ 2
        btImageCoords.y = (m_BackBuffer.getDIBHeight - btImage.getDIBHeight) \ 2
    End If
    
    'No other special preparation is required for this control, so proceed with recreating the back buffer
    redrawBackBuffer
            
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    'Store all associated properties
    With PropBag
        .WriteProperty "BackColor", m_BackColor, vbWhite
        .WriteProperty "AutoToggle", m_AutoToggle, False
    End With
    
End Sub

'External functions can call this to request a redraw.  This is helpful for live-updating theme settings, as in the Preferences dialog.
Public Sub updateAgainstCurrentTheme()
    
    'Redraw the control, which will also cause a resync against any theme changes
    updateControlSize
    
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
    ' 4) The central themer (which contains default values for all these scenarios)
    Dim btnColorBorder As Long, btnColorFill As Long
    Dim curColor As Long
        
    If Me.Enabled Then
    
        'Is the button pressed?
        If m_ButtonState Then
            btnColorFill = g_Themer.getThemeColor(PDTC_ACCENT_ULTRALIGHT)
            btnColorBorder = g_Themer.getThemeColor(PDTC_ACCENT_HIGHLIGHT)
            
        'The button is not pressed
        Else
        
            'Is the mouse inside the UC?
            If m_MouseInsideUC Then
                btnColorFill = m_BackColor
                btnColorBorder = g_Themer.getThemeColor(PDTC_ACCENT_DEFAULT)
            
            'The mouse is not inside the UC
            Else
                btnColorFill = m_BackColor
                btnColorBorder = m_BackColor
            
            End If
            
        End If
        
    'The button is disabled
    Else
    
        btnColorFill = m_BackColor
        btnColorBorder = m_BackColor
        
    End If
    
    'First, we fill the button interior with the established fill color
    GDI_Plus.GDIPlusFillDIBRect m_BackBuffer, 0, 0, m_BackBuffer.getDIBWidth - 1, m_BackBuffer.getDIBHeight - 1, btnColorFill, 255
    
    'A single-pixel border is always drawn around the control
    GDI_Plus.GDIPlusDrawRectOutlineToDC m_BackBuffer.getDIBDC, 0, 0, m_BackBuffer.getDIBWidth - 1, m_BackBuffer.getDIBHeight - 1, btnColorBorder, 255, 1
    
    'TODO: if this button has received focus via keyboard, paint it with some special marker
    'If m_FocusRectActive Then
    '    GDI_Plus.GDIPlusDrawRectOutlineToDC m_BackBuffer.getDIBDC, .btBounds.Left + 2, .btBounds.Top + 2, .btBounds.Right - 2, .btBounds.Bottom - 3, btnColorActiveBorder, 255, 1
    'End If
                    
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
