VERSION 5.00
Begin VB.UserControl pdRadioButton 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   555
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3735
   ClipBehavior    =   0  'None
   ClipControls    =   0   'False
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
   ForeColor       =   &H00404040&
   HasDC           =   0   'False
   MousePointer    =   99  'Custom
   ScaleHeight     =   37
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   249
   ToolboxBitmap   =   "pdRadioButton.ctx":0000
End
Attribute VB_Name = "pdRadioButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Radio Button control
'Copyright 2013-2026 by Tanner Helland
'Created: 28/January/13
'Last updated: 15/February/16
'Last update: finalize theming support
'
'In a surprise to precisely no one, PhotoDemon has some unique needs when it comes to user controls - needs that
' the intrinsic VB controls can't handle.  These range from the obnoxious (lack of an "autosize" property for
' anything but labels) to the critical (no Unicode support).
'
'As such, I've created many of my own UCs for the program.  All are owner-drawn, with the goal of maintaining
' visual fidelity across the program, while also enabling key features like Unicode support.
'
'A few notes on this radio button replacement, specifically:
'
' 1) The control is no longer autosized based on the current font and caption.  If a caption exceeds the size of the
'     (manually set) width, the font size will be repeatedly reduced until the caption fits.
' 2) High DPI settings are handled automatically, so do not attempt to handle this manually.
' 3) A hand cursor is automatically applied, and clicks on both the button and label are registered properly.
' 4) Coloration is automatically handled by PD's internal theming engine.
' 5) When the control receives focus via keyboard, a special focus rect is drawn.  Focus via mouse is conveyed via text glow.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'This control really only needs one event raised - Click
Public Event Click()
Attribute Click.VB_UserMemId = -600

'Because VB focus events are wonky, especially when we use CreateWindow within a UC, this control raises its own
' specialized focus events.  If you need to track focus, use these instead of the default VB functions.
Public Event GotFocusAPI()
Public Event LostFocusAPI()

'If we cannot physically fit a translated caption into the user control's area (because we run out of allowable font sizes),
' this failure state will be set to TRUE.  When that happens, ellipses will be forcibly appended to the control caption.
Private m_FitFailure As Boolean

'Current control value
Private m_Value As Boolean

'Rect where the caption is rendered.  This is calculated by UpdateControlLayout, and it needs to be revisited if the
' caption changes, or the control size changes.
Private m_CaptionRect As RectF

'Similar rect for the radio button itself
Private m_RadioButtonRect As RectF

'Whenever the caption changes or the control is resized, the "clickable" rect must be updated.  (This control allows the user
' to click on either the radio button, or the caption itself.)  It's tracked separately, because there's some fairly messy
' padding calculations involved in positioning the radio button and caption relative to the control as a whole.
Private m_ClickableRect As RectF, m_MouseInsideClickableRect As Boolean

'User control support class.  Historically, many classes (and associated subclassers) were required by each user control,
' but I've since wrapped these into a single central support class.
Private WithEvents ucSupport As pdUCSupport
Attribute ucSupport.VB_VarHelpID = -1

'Local list of themable colors.  This list includes all potential colors used by this class, regardless of state change
' or internal control settings.  The list is updated by calling the UpdateColorList function.
' (Note also that this list does not include variants, e.g. "BorderColor" vs "BorderColor_Hovered".  Variant values are
'  automatically calculated by the color management class, and they are retrieved by passing boolean modifiers to that
'  class, rather than treating every imaginable variant as a separate constant.)
Private Enum PDRADIOBUTTON_COLOR_LIST
    [_First] = 0
    PDRB_Background = 0
    PDRB_Caption = 1
    PDRB_ButtonFill = 2
    PDRB_ButtonBorder = 3
    [_Last] = 3
    [_Count] = 4
End Enum

'Color retrieval and storage is handled by a dedicated class; this allows us to optimize theme interactions,
' without worrying about the details locally.
Private m_Colors As pdThemeColors

Public Function GetControlType() As PD_ControlType
    GetControlType = pdct_RadioButton
End Function

Public Function GetControlName() As String
    GetControlName = UserControl.Extender.Name
End Function

'IMPORTANT NOTE: only the ENGLISH caption is returned.  I don't have a reason for returning a translated caption (if any),
'                 but I can revisit in the future if it ever becomes relevant.
Public Property Get Caption() As String
Attribute Caption.VB_UserMemId = -518
    Caption = ucSupport.GetCaptionText
End Property

Public Property Let Caption(ByRef newCaption As String)
    ucSupport.SetCaptionText newCaption
    PropertyChanged "Caption"
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

Public Property Get FontSize() As Single
    FontSize = ucSupport.GetCaptionFontSize
End Property

Public Property Let FontSize(ByVal newSize As Single)
    ucSupport.SetCaptionFontSize newSize
    PropertyChanged "FontSize"
End Property

Public Property Get hWnd() As Long
Attribute hWnd.VB_UserMemId = -515
    hWnd = UserControl.hWnd
End Property

'Container hWnd is used to make sure radio button groups lie within the same parent control.
Public Property Get ContainerHwnd() As Long
    ContainerHwnd = UserControl.ContainerHwnd
End Property

Public Property Get Value() As Boolean
Attribute Value.VB_UserMemId = 0
    Value = m_Value
End Property

Public Property Let Value(ByVal newValue As Boolean)
    
    If (m_Value <> newValue) Then
    
        m_Value = newValue
        RedrawBackBuffer
        
        'It's important to only raise change events when a radio button is set to TRUE.  Otherwise, clicking one button will cause
        ' Click() events to fire for all other radio buttons (as they're being set to FALSE).
        If newValue Then
            UpdateOtherButtons
            RaiseEvent Click
            PropertyChanged "Value"
        End If
        
    End If
    
End Property

'To support high-DPI settings properly, we expose specialized move+size functions
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

'Call to reset all other radio buttons to match this button's new state.  This button's state must be TRUE.
Private Sub UpdateOtherButtons()

    'If the option button is set to TRUE, turn off all other option buttons on a form
    If m_Value And PDMain.IsProgramRunning() Then

        'Enumerate through each control on the form; if it's another option button whose value is TRUE, set it to FALSE
        Dim eControl As Object
        For Each eControl In Parent.Controls
            If TypeOf eControl Is pdRadioButton Then
                If eControl.Container.hWnd = UserControl.ContainerHwnd Then
                    If Not (eControl.hWnd = UserControl.hWnd) Then
                        If eControl.Value Then eControl.Value = False
                    End If
                End If
            End If
        Next eControl
    
    End If
    
End Sub

Private Sub ucSupport_GotFocusAPI()
    RedrawBackBuffer
    RaiseEvent GotFocusAPI
End Sub

Private Sub ucSupport_LostFocusAPI()
    RedrawBackBuffer
    RaiseEvent LostFocusAPI
End Sub

'Space and Enter keypresses toggle control state
Private Sub ucSupport_KeyDownCustom(ByVal Shift As ShiftConstants, ByVal vkCode As Long, markEventHandled As Boolean)
    
    markEventHandled = False
    
    If Me.Enabled And ((vkCode = VK_SPACE) Or (vkCode = VK_RETURN)) Then
        markEventHandled = True
        Me.Value = True
    End If
    
End Sub

Private Sub ucSupport_KeyDownSystem(ByVal Shift As ShiftConstants, ByVal whichSysKey As PD_NavigationKey, markEventHandled As Boolean)
    
    'Enter/Esc get reported directly to the system key handler.  Note that we track the return, because TRUE
    ' means the key was successfully forwarded to the relevant handler.  (If FALSE is returned, no control
    ' accepted the keypress, meaning we should forward the event down the line.)
    markEventHandled = NavKey.NotifyNavKeypress(Me, whichSysKey, Shift)
    
End Sub

'To improve responsiveness, MouseDown is used instead of Click
Private Sub ucSupport_MouseDownCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal timeStamp As Long)
    If Me.Enabled And IsMouseOverClickArea(x, y) Then
        Me.Value = True
    End If
End Sub

'When the mouse leaves the UC, we must repaint the caption (as it's no longer hovered)
Private Sub ucSupport_MouseLeave(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    m_MouseInsideClickableRect = False
    RedrawBackBuffer
End Sub

'When the mouse enters the clickable portion of the UC, we must repaint the caption (to reflect its hovered state)
Private Sub ucSupport_MouseMoveCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal timeStamp As Long)

    'If the mouse is over the relevant portion of the user control, display the cursor as clickable
    If (m_MouseInsideClickableRect <> IsMouseOverClickArea(x, y)) Then
        m_MouseInsideClickableRect = IsMouseOverClickArea(x, y)
        RedrawBackBuffer
    End If
    
    If m_MouseInsideClickableRect Then
        ucSupport.RequestCursor IDC_HAND
    Else
        ucSupport.RequestCursor IDC_DEFAULT
    End If

End Sub

Private Sub ucSupport_RepaintRequired(ByVal updateLayoutToo As Boolean)
    If updateLayoutToo Then UpdateControlLayout Else RedrawBackBuffer
End Sub

'See if the mouse is over the clickable portion of the control
Private Function IsMouseOverClickArea(ByVal mouseX As Single, ByVal mouseY As Single) As Boolean
    IsMouseOverClickArea = PDMath.IsPointInRectF(mouseX, mouseY, m_ClickableRect)
End Function

Private Sub UserControl_Initialize()
    
    'Initialize a user control support class
    Set ucSupport = New pdUCSupport
    ucSupport.RegisterControl UserControl.hWnd, True
    
    'Request any control-specific functionality
    ucSupport.RequestExtraFunctionality True, True
    ucSupport.SpecifyRequiredKeys VK_SPACE, VK_RETURN
    ucSupport.RequestCaptionSupport
    ucSupport.SetCaptionAutomaticPainting False
    
    'Prep the color manager and load default colors
    Set m_Colors = New pdThemeColors
    Dim colorCount As PDRADIOBUTTON_COLOR_LIST: colorCount = [_Count]
    m_Colors.InitializeColorList "PDRadioButton", colorCount
    If Not PDMain.IsProgramRunning() Then UpdateColorList
                
End Sub

'Set default properties
Private Sub UserControl_InitProperties()
    Me.Caption = "caption"
    Me.FontSize = 10
    m_Value = True
End Sub

'At run-time, painting is handled by PD's pdWindowPainter class.  In the IDE, however, we must rely on VB's internal paint event.
Private Sub UserControl_Paint()
    ucSupport.RequestIDERepaint UserControl.hDC
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        Me.Caption = .ReadProperty("Caption", vbNullString)
        Me.FontSize = .ReadProperty("FontSize", 10)
        Me.Value = .ReadProperty("Value", False)
    End With
End Sub

Private Sub UserControl_Resize()
    If Not PDMain.IsProgramRunning() Then If Not PDMain.IsProgramRunning() Then ucSupport.NotifyIDEResize UserControl.Width, UserControl.Height
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "Caption", ucSupport.GetCaptionText, "caption"
        .WriteProperty "FontSize", ucSupport.GetCaptionFontSize, 10
        .WriteProperty "Value", m_Value, False
    End With
End Sub

'Whenever the size of the control changes, we must recalculate some internal rendering metrics.
Private Sub UpdateControlLayout()
    
    'Retrieve DPI-aware control dimensions from the support class
    Dim bWidth As Long, bHeight As Long
    bWidth = ucSupport.GetBackBufferWidth
    bHeight = ucSupport.GetBackBufferHeight
    
    'A little bit of horizontal and vertical padding is applied around various control parts
    Const vTextPadding As Long = 1&, hBoxCaptionPadding As Long = 8&
    
    'Next, determine the positioning of the caption, if present.  (ucSupport.GetCaptionBottom tells us where the
    ' caption text ends vertically.)
    If ucSupport.IsCaptionActive Then
        
        'Start by making sure the control is tall enough to fit the caption.  (Control height is auto-controlled at present.)
        If (ucSupport.GetCaptionHeight(False) + Interface.FixDPI(vTextPadding) * 2 <> bHeight) Then
            bHeight = ucSupport.GetCaptionHeight(False) + Interface.FixDPI(vTextPadding) * 2
            ucSupport.RequestNewSize bWidth, bHeight, False
        End If
        
    End If
    
    'Because the radio button size and font size are inextricably connected, we now need to retrieve a font object matching
    ' the current control font size.  That font's metrics will determine how everything gets positioned.
    Dim tmpFont As pdFont
    Set tmpFont = Fonts.GetMatchingUIFont(ucSupport.GetCaptionFontSize)
    tmpFont.SetTextAlignment vbLeftJustify
    
    'Retrieve the height of the current caption, or if no caption exists, a placeholder
    Dim captionHeight As Long
    If ucSupport.IsCaptionActive Then
        captionHeight = tmpFont.GetHeightOfString(ucSupport.GetCaptionTextTranslated)
    Else
        captionHeight = tmpFont.GetHeightOfString("ABjy69")
    End If
    
    'Using the font metrics, determine a check box offset and size.  Note that 1px is manually added as part of maintaining a
    ' 1px border around the user control as a whole (which is used for a keyboard focus rect).
    Dim offsetX As Long, offsetY As Long, radioButtonSize As Long
    offsetX = 1 + Interface.FixDPI(2)
    offsetY = offsetX
    radioButtonSize = bHeight - (offsetY * 2)
    
    'Use that to populate the radio button rect; we store it at module-level, and use it for rendering and hit-detection
    With m_RadioButtonRect
        .Left = offsetX
        .Top = offsetY
        .Width = radioButtonSize
        .Height = radioButtonSize
    End With
    
    'Pass the available space to the support class; it needs this information to auto-fit the caption
    Dim captionLeft As Long
    captionLeft = (m_RadioButtonRect.Left + m_RadioButtonRect.Width) + Interface.FixDPI(hBoxCaptionPadding)
    ucSupport.SetCaptionCustomPosition captionLeft, 0, bWidth - captionLeft, bHeight
    
    'While here, calculate a caption rect, taking into account the auto-sized caption text (which may be using a different font size)
    With m_CaptionRect
        .Left = captionLeft
        .Top = (bHeight - ucSupport.GetCaptionHeight(True)) / 2
        .Width = ucSupport.GetCaptionWidth(True) + 1
        If (.Left + .Width > bWidth) Then .Width = (bWidth - .Left)
        .Height = ucSupport.GetCaptionHeight(True) + 1
    End With
    
    'The clickable rect is the union of the radio button and caption rect.  Calculate it now.
    PDMath.UnionRectF m_ClickableRect, m_RadioButtonRect, m_CaptionRect
    
    'If the caption still does not fit within the available area (typically because we reached the minimum allowable font
    ' size, but the caption was *still* too long), set a module-level failure state to TRUE.  This notifies the renderer
    ' that ellipses must be forcibly appended to the caption.
    m_FitFailure = (ucSupport.GetCaptionWidth(True) > bWidth - m_CaptionRect.Left)
    
    RedrawBackBuffer
            
End Sub

'Use this function to completely redraw the back buffer from scratch.  Note that this is computationally expensive compared to just flipping the
' existing buffer to the screen, so only redraw the backbuffer if the control state has somehow changed.
Private Sub RedrawBackBuffer()
    
    'Request the back buffer DC, and ask the support module to erase any existing rendering for us.
    Dim bufferDC As Long
    bufferDC = ucSupport.GetBackBufferDC(True, m_Colors.RetrieveColor(PDRB_Background, Me.Enabled))
    If (bufferDC = 0) Then Exit Sub
    
    Dim bWidth As Long, bHeight As Long
    bWidth = ucSupport.GetBackBufferWidth
    bHeight = ucSupport.GetBackBufferHeight
    
    If PDMain.IsProgramRunning() Then
        
        'Populate colors from the central theme object
        Dim radioColorBorder As Long, radioColorFill As Long, txtColor As Long
        radioColorBorder = m_Colors.RetrieveColor(PDRB_ButtonBorder, Me.Enabled, m_Value, m_MouseInsideClickableRect Or ucSupport.DoIHaveFocus)
        radioColorFill = m_Colors.RetrieveColor(PDRB_ButtonFill, Me.Enabled, m_Value, m_MouseInsideClickableRect)
        txtColor = m_Colors.RetrieveColor(PDRB_Caption, Me.Enabled, m_Value, m_MouseInsideClickableRect Or ucSupport.DoIHaveFocus)
        
        'pd2D is used for painting
        Dim cSurface As pd2DSurface
        Drawing2D.QuickCreateSurfaceFromDC cSurface, bufferDC, True
        
        'Draw the radio button border
        Dim borderWidth As Single, cPen As pd2DPen
        If m_MouseInsideClickableRect Or ucSupport.DoIHaveFocus Then borderWidth = 3! Else borderWidth = 1.5!
        Drawing2D.QuickCreateSolidPen cPen, borderWidth, radioColorBorder
        
        With m_RadioButtonRect
            PD2D.DrawCircleF cSurface, cPen, .Left + .Width * 0.5, .Top + .Height * 0.5, .Width * 0.5
        End With
        
        'If the button state is TRUE, draw a smaller circle inside the border
        Dim cBrush As pd2DBrush
        Drawing2D.QuickCreateSolidBrush cBrush, radioColorFill
        
        If m_Value Then
            With m_RadioButtonRect
                PD2D.FillEllipseF cSurface, cBrush, .Left + Interface.FixDPIFloat(3), .Top + Interface.FixDPIFloat(3), .Width - Interface.FixDPIFloat(6), .Height - Interface.FixDPIFloat(6)
            End With
        End If
        
        Set cPen = Nothing: Set cBrush = Nothing: Set cSurface = Nothing
        
    End If
    
    'Render the text, appending ellipses as necessary
    If m_FitFailure Then
        ucSupport.PaintCaptionManually_Clipped m_CaptionRect.Left, m_CaptionRect.Top, m_CaptionRect.Width, m_CaptionRect.Height, txtColor, True
    Else
        ucSupport.PaintCaptionManually m_CaptionRect.Left, m_CaptionRect.Top, txtColor
    End If
    
    'Paint the final result to the screen, as relevant
    ucSupport.RequestRepaint

End Sub

'Before this control does any painting, we need to retrieve relevant colors from PD's primary theming class.  Note that this
' step must also be called if/when PD's visual theme settings change.
Private Sub UpdateColorList()
    With m_Colors
        .LoadThemeColor PDRB_Background, "Background", IDE_WHITE
        .LoadThemeColor PDRB_Caption, "Caption", IDE_GRAY
        .LoadThemeColor PDRB_ButtonFill, "ButtonFill", IDE_BLUE
        .LoadThemeColor PDRB_ButtonBorder, "ButtonBorder", IDE_BLACK
    End With
End Sub

'External functions can call this to request a redraw.  This is helpful for live-updating theme settings, as in the Preferences dialog.
Public Sub UpdateAgainstCurrentTheme(Optional ByVal hostFormhWnd As Long = 0)
    If ucSupport.ThemeUpdateRequired Then
        UpdateColorList
        If PDMain.IsProgramRunning() Then NavKey.NotifyControlLoad Me, hostFormhWnd
        If PDMain.IsProgramRunning() Then ucSupport.UpdateAgainstThemeAndLanguage
    End If
End Sub

'By design, PD prefers to not use design-time tooltips.  Apply tooltips at run-time, using this function.
' (IMPORTANT NOTE: translations are handled automatically.  Always pass the original English text!)
Public Sub AssignTooltip(ByRef newTooltip As String, Optional ByRef newTooltipTitle As String = vbNullString, Optional ByVal raiseTipsImmediately As Boolean = False)
    ucSupport.AssignTooltip UserControl.ContainerHwnd, newTooltip, newTooltipTitle, raiseTipsImmediately
End Sub
