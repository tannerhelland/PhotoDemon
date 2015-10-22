VERSION 5.00
Begin VB.UserControl colorSelector 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   1710
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   ClipControls    =   0   'False
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   MousePointer    =   99  'Custom
   ScaleHeight     =   114
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   ToolboxBitmap   =   "colorSelector.ctx":0000
End
Attribute VB_Name = "colorSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Color Selector custom control
'Copyright 2013-2015 by Tanner Helland
'Created: 17/August/13
'Last updated: 04/September/15
'Last update: added caption support, to simplify layout code in various effect dialogs
'
'This thin user control is basically an empty control that when clicked, displays a color selection window.  If a
' color is selected (e.g. Cancel is not pressed), it updates its back color to match, and raises a "ColorChanged"
' event.
'
'Though simple, this control solves a lot of problems.  It is especially helpful for improving interaction with the
' command bar user control, as it easily supports color reset/randomize/preset events.  It is also nice to be able
' to update a single master function for color selection, then have the change propagate to all tool windows.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************


Option Explicit

'This control doesn't really do anything interesting, besides allow a color to be selected.
Public Event ColorChanged()

'A specialized class handles mouse input for this control
Private WithEvents cMouseEvents As pdInputMouse
Attribute cMouseEvents.VB_VarHelpID = -1

'Reliable focus detection requires a specialized subclasser
Private WithEvents cFocusDetector As pdFocusDetector
Attribute cFocusDetector.VB_VarHelpID = -1
Public Event GotFocusAPI()
Public Event LostFocusAPI()

'Flicker-free window painter
Private WithEvents cPainter As pdWindowPainter
Attribute cPainter.VB_VarHelpID = -1

'Additional helper for rendering themed and multiline tooltips
Private toolTipManager As pdToolTip

'pdCaption manages all caption-related settings, so we don't have to.  (Note that this includes not just the caption, but related
' settings like caption font size.)
Private m_Caption As pdCaption

'This control uses two basic layout rects: one for the control title label (if any), and another for the clickable
' color region.  These rects are calculated by the updateControlLayout function.
Private m_ButtonRect As RECT, m_CaptionRect As RECT

'Persistent back buffer, which we manage internally.  This allows for color management (yes, even on UI elements!)
Private m_BackBuffer As pdDIB

'The control's current color
Private curColor As OLE_COLOR

'When the select color dialog is live, this will be set to TRUE
Private isDialogLive As Boolean

'This value will be TRUE while the mouse is inside the UC, or specifically the clickable button region
Private m_MouseInsideUC As Boolean, m_MouseInsideButton As Boolean

'When the control receives focus via keyboard (e.g. NOT by mouse events), we draw a focus rect to help orient the user.
Private m_FocusRectActive As Boolean

'Last-known mouse coords.  These are used to determine if the mouse is over the clickable region of the control.
Private m_MouseX As Single, m_MouseY As Single

'Used to prevent recursive redraws
Private m_InternalResizeActive As Boolean

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
    If (m_Caption.setCaption(newCaption) And m_ControlIsVisible) Or (Not g_IsProgramRunning) Then updateControlLayout
    PropertyChanged "Caption"
End Property

'At present, all this control does is store a color value
Public Property Get Color() As OLE_COLOR
    Color = curColor
End Property

Public Property Let Color(ByVal newColor As OLE_COLOR)
    
    curColor = newColor
    redrawBackBuffer
    
    PropertyChanged "Color"
    RaiseEvent ColorChanged
    
End Property

'The Enabled property is a bit unique; see http://msdn.microsoft.com/en-us/library/aa261357%28v=vs.60%29.aspx
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_UserMemId = -514
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

Public Property Get containerHwnd() As Long
    containerHwnd = UserControl.containerHwnd
End Property

'Call this to force a display of the color window.  Note that it's *public*, so outside callers can raise dialogs, too.
Public Sub DisplayColorSelection()
    
    isDialogLive = True
    
    'Store the current color
    Dim newColor As Long, oldColor As Long
    oldColor = Color
    
    'Use the default color dialog to select a new color
    If showColorDialog(newColor, CLng(curColor), Me) Then
        Color = newColor
    Else
        Color = oldColor
    End If
    
    isDialogLive = False
    
End Sub

Private Sub cMouseEvents_MouseDownCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    
    m_MouseX = x
    m_MouseY = y
    
    'Ensure that a focus event has been raised, if it wasn't already
    If Not cFocusDetector.HasFocus Then cFocusDetector.setFocusManually
    
    If isMouseInButtonArea() And ((Button Or pdLeftButton) <> 0) Then DisplayColorSelection
    
End Sub

Private Sub cMouseEvents_MouseEnter(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    
    m_MouseX = x
    m_MouseY = y
    m_MouseInsideUC = True
    m_MouseInsideButton = isMouseInButtonArea()
    
    redrawBackBuffer
    updateCursor
    
End Sub

Private Sub cMouseEvents_MouseLeave(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    
    m_MouseX = -1
    m_MouseY = -1
    m_MouseInsideButton = False
    m_MouseInsideUC = False
    
    redrawBackBuffer
    updateCursor
    
End Sub

Private Sub cMouseEvents_MouseMoveCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    
    m_MouseX = x
    m_MouseY = y
    updateCursor
    
    If isMouseInButtonArea() <> m_MouseInsideButton Then
        m_MouseInsideButton = isMouseInButtonArea()
        redrawBackBuffer
    End If
    
End Sub

'When the control receives focus, relay the event externally
Private Sub cFocusDetector_GotFocusReliable()
    m_FocusRectActive = True
    RaiseEvent GotFocusAPI
End Sub

'When the control loses focus, relay the event externally
Private Sub cFocusDetector_LostFocusReliable()
    m_FocusRectActive = False
    RaiseEvent LostFocusAPI
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

Private Sub UserControl_Initialize()
        
    'Prep the caption object
    Set m_Caption = New pdCaption
    m_Caption.setWordWrapSupport False
    
    If g_IsProgramRunning Then
        
        'Initialize mouse handling
        Set cMouseEvents = New pdInputMouse
        cMouseEvents.addInputTracker UserControl.hWnd, True, True, , True
        cMouseEvents.setSystemCursor IDC_HAND
        
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
    
    m_MouseX = -1
    m_MouseY = -1
    
    'Update the control size parameters at least once
    updateControlLayout
    
End Sub

Private Sub UserControl_InitProperties()
    Color = RGB(255, 255, 255)
    FontSize = 12
    Caption = ""
End Sub

'At run-time, painting is handled by PD's pdWindowPainter class.  In the IDE, however, we must rely on VB's internal paint event.
Private Sub UserControl_Paint()
    
    'Provide minimal painting within the designer
    If Not g_IsProgramRunning Then updateControlLayout
    
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        Color = .ReadProperty("curColor", RGB(255, 255, 255))
        Caption = .ReadProperty("Caption", "")
        FontSize = .ReadProperty("FontSize", 12)
    End With
End Sub

Private Sub UserControl_Resize()
    updateControlLayout
End Sub

Private Sub UserControl_Show()
    m_ControlIsVisible = True
    updateControlLayout
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "Caption", m_Caption.getCaptionEn, ""
        .WriteProperty "FontSize", m_Caption.getFontSize, 12
        .WriteProperty "curColor", curColor, RGB(255, 255, 255)
    End With
End Sub

'Whenever a control property changes that affects control size or layout (including internal changes, like caption adjustments),
' call this function to recalculate the control's layout
Private Sub updateControlLayout()
    
    If m_InternalResizeActive Then Exit Sub
    
    'Set a control-level flag to prevent recursive redraws
    m_InternalResizeActive = True
    
    'To allow the control to render correctly in the IDE, determine a background color in advance
    Dim controlBackgroundColor As Long
    If g_IsProgramRunning Then
        controlBackgroundColor = g_Themer.getThemeColor(PDTC_BACKGROUND_DEFAULT)
    Else
        controlBackgroundColor = RGB(255, 255, 255)
    End If
    
    'Start by resetting the back buffer, as necessary
    If m_BackBuffer Is Nothing Then Set m_BackBuffer = New pdDIB
    
    If (m_BackBuffer.getDIBWidth <> UserControl.ScaleWidth) Or (m_BackBuffer.getDIBHeight <> UserControl.ScaleHeight) Then
        m_BackBuffer.createBlank UserControl.ScaleWidth, UserControl.ScaleHeight, 24, controlBackgroundColor
    Else
        If g_IsProgramRunning Then
            GDI_Plus.GDIPlusFillDIBRect m_BackBuffer, 0, 0, m_BackBuffer.getDIBWidth, m_BackBuffer.getDIBHeight, controlBackgroundColor, 255
        Else
            Drawing.fillRectToDC m_BackBuffer.getDIBDC, -1, -1, m_BackBuffer.getDIBWidth + 1, m_BackBuffer.getDIBHeight + 1, controlBackgroundColor
        End If
    End If
    
    'Next, we need to determine the positioning of the caption, if present.
    If m_Caption.isCaptionActive Then
        
        'Notify the caption renderer of our width.  It will auto-fit its font to match.
        ' (Because this control doesn't support wordwrap, container height is irrelevant; pass 0)
        m_Caption.setControlSize m_BackBuffer.getDIBWidth, 0
        
        'We now have all the information necessary to calculate caption positioning.
        With m_CaptionRect
            .Top = 0
            .Left = 0
            .Right = m_BackBuffer.getDIBWidth
            .Bottom = m_Caption.getCaptionHeight() + FixDPI(6)
        End With
        
        With m_ButtonRect
            .Top = m_CaptionRect.Bottom + 1
            .Left = FixDPI(8)
            .Right = m_BackBuffer.getDIBWidth - 1
            .Bottom = m_BackBuffer.getDIBHeight - 1
        End With
        
        'We actually paint the caption now, to spare us having to do it in the interior redraw loop
        m_Caption.drawCaption m_BackBuffer.getDIBDC, 1, 1
        
    'If there's no caption, allow the clickable portion to fill the entire control
    Else
        
        With m_ButtonRect
            .Top = 0
            .Left = 0
            .Right = m_BackBuffer.getDIBWidth - 1
            .Bottom = m_BackBuffer.getDIBHeight - 1
        End With
        
    End If
    
    'Reset the redraw flag, and request a background repaint
    m_InternalResizeActive = False
    redrawBackBuffer
            
End Sub

'When the mouse moves, the cursor should be updated to match
Private Sub updateCursor()
    If isMouseInButtonArea() Then
        cMouseEvents.setSystemCursor IDC_HAND
    Else
        cMouseEvents.setSystemCursor IDC_DEFAULT
    End If
End Sub

'Returns TRUE if the mouse is inside the clickable region of the color selector (e.g. NOT the caption area, if one exists)
Private Function isMouseInButtonArea() As Boolean
    isMouseInButtonArea = Math_Functions.isPointInRect(m_MouseX, m_MouseY, m_ButtonRect)
End Function

'Redraw the entire control, including the caption (if present)
Private Sub redrawBackBuffer()
    
    'If a caption exists, it has already been drawn.  We just need to draw the clickable button portion.
    
    'Use the API to draw borders around the control
    If g_IsProgramRunning Then
    
        'Draw borders around the brush results.
        Dim outlineColor As Long, outlineSize As Single
        
        If Me.Enabled Then
        
            If isMouseInButtonArea() Then
                outlineColor = g_Themer.getThemeColor(PDTC_ACCENT_DEFAULT)
                outlineSize = 3#
            Else
                outlineColor = vbBlack
                outlineSize = 1#
            End If
            
        Else
            outlineColor = g_Themer.getThemeColor(PDTC_DISABLED)
            outlineSize = 1#
        End If
        
        'Render the button
        With m_ButtonRect
            GDI_Plus.GDIPlusFillDIBRect m_BackBuffer, .Left, .Top, .Right, .Bottom, Me.Color, 255
            
            If outlineSize > 1 Then
                Dim outlineModifier As Single
                outlineModifier = outlineSize / 2
                GDI_Plus.GDIPlusDrawRectOutlineToDC m_BackBuffer.getDIBDC, .Left + outlineModifier, .Top + outlineModifier, .Right - outlineModifier + 1, .Bottom - outlineModifier + 1, outlineColor, 255, outlineSize, False, LineJoinMiter
            Else
                GDI_Plus.GDIPlusDrawRectOutlineToDC m_BackBuffer.getDIBDC, .Left, .Top, .Right, .Bottom, outlineColor, 255, outlineSize, False, LineJoinMiter
            End If
        End With
        
    'In the designer, use GDI to do the same thing
    Else
        With m_ButtonRect
            Drawing.fillRectToDC m_BackBuffer.getDIBDC, .Left, .Top, .Right + 1, .Bottom + 1, Me.Color
            Drawing.outlineRectToDC m_BackBuffer.getDIBDC, .Left, .Top, .Right + 1, .Bottom + 1, vbBlack
        End With
    End If
    
    'Paint the final result to the screen, as relevant
    If g_IsProgramRunning Then
        cPainter.requestRepaint
    Else
        BitBlt UserControl.hDC, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, m_BackBuffer.getDIBDC, 0, 0, vbSrcCopy
    End If
    
End Sub

'If a color selection dialog is active, it will pass color updates backward to this function, so that we can let
' our parent form display live updates *while the user is playing with colors* - very cool!
Public Sub notifyOfLiveColorChange(ByVal newColor As Long)
    Color = newColor
End Sub

'External functions can call this to request a redraw.  This is helpful for live-updating theme settings, as in the Preferences dialog.
Public Sub UpdateAgainstCurrentTheme()
    
    If g_IsProgramRunning Then
            
        'Our tooltip object must also be refreshed (in case the language has changed)
        toolTipManager.UpdateAgainstCurrentTheme
        
        'The caption manager will also refresh itself
        m_Caption.UpdateAgainstCurrentTheme
        
        'Re-enable color management for the underlying UC
        Color_Management.TurnOnDefaultColorManagement UserControl.hDC, UserControl.hWnd
        
    End If
    
    'Update the control's layout to account for new translations and/or theme changes
    updateControlLayout
    
    'Redraw the control to match any updated settings
    redrawBackBuffer
    
End Sub

'Due to complex interactions between user controls and PD's translation engine, tooltips require this dedicated function.
' (IMPORTANT NOTE: the tooltip class will handle translations automatically.  Always pass the original English text!)
Public Sub AssignTooltip(ByVal newTooltip As String, Optional ByVal newTooltipTitle As String, Optional ByVal newTooltipIcon As TT_ICON_TYPE = TTI_NONE)
    toolTipManager.setTooltip Me.hWnd, UserControl.containerHwnd, newTooltip, newTooltipTitle, newTooltipIcon
End Sub
