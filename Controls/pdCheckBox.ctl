VERSION 5.00
Begin VB.UserControl pdCheckBox 
   BackColor       =   &H80000005&
   ClientHeight    =   375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2520
   ClipControls    =   0   'False
   DrawStyle       =   5  'Transparent
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HasDC           =   0   'False
   MousePointer    =   99  'Custom
   ScaleHeight     =   25
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   168
   ToolboxBitmap   =   "pdCheckBox.ctx":0000
End
Attribute VB_Name = "pdCheckBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Checkbox control
'Copyright 2013-2018 by Tanner Helland
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
'A few notes on this checkbox replacement, specifically:
'
' 1) The control is no longer autosized based on the current font and caption.  If a caption exceeds the size of the
'     (manually set) width, the font size will be repeatedly reduced until the caption fits.
' 2) High DPI settings are handled automatically, so do not attempt to handle this manually.
' 3) A hand cursor is automatically applied, and clicks on both the button and label are registered properly.
' 4) Coloration is automatically handled by PD's internal theming engine.
' 5) When the control receives focus via keyboard, a special focus rect is drawn.  Focus via mouse is conveyed via text glow.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'This control really only needs one event raised - Click
Public Event Click()

'Because VB focus events are wonky, especially when we use CreateWindow within a UC, this control raises its own
' specialized focus events.  If you need to track focus, use these instead of the default VB functions.
Public Event GotFocusAPI()
Public Event LostFocusAPI()

'If we cannot physically fit a translated caption into the user control's area (because we run out of allowable font sizes),
' this failure state will be set to TRUE.  When that happens, ellipses will be forcibly appended to the control caption.
Private m_FitFailure As Boolean

'Current control value
Private m_Value As CheckBoxConstants

'Rect where the caption is rendered.  This is calculated by UpdateControlLayout, and it needs to be revisited if the
' caption changes, or the control size changes.
Private m_CaptionRect As RectF

'Similar rect for the checkbox
Private m_CheckboxRect As RectF

'Whenever the caption changes or the control is resized, the "clickable" rect must be updated.  (This control allows the user
' to click on either the checkbox, or the caption itself.)  It's tracked separately, because there's some fairly messy padding
' calculations involved in positioning the checkbox and caption relative to the control as a whole.
Private m_ClickableRect As RectF, m_MouseInsideClickableRect As Boolean

'User control support class.  Historically, many classes (and associated subclassers) were required by each user control,
' but I've since attempted to wrap these into a single master control support class.
Private WithEvents ucSupport As pdUCSupport
Attribute ucSupport.VB_VarHelpID = -1

'Local list of themable colors.  This list includes all potential colors used by this class, regardless of state change
' or internal control settings.  The list is updated by calling the UpdateColorList function.
' (Note also that this list does not include variants, e.g. "BorderColor" vs "BorderColor_Hovered".  Variant values are
'  automatically calculated by the color management class, and they are retrieved by passing boolean modifiers to that
'  class, rather than treating every imaginable variant as a separate constant.)
Private Enum PDCHECKBOX_COLOR_LIST
    [_First] = 0
    PDCB_Background = 0
    PDCB_Caption = 1
    PDCB_ButtonFill = 2
    PDCB_ButtonBorder = 3
    PDCB_Checkmark = 4
    [_Last] = 4
    [_Count] = 5
End Enum

'Color retrieval and storage is handled by a dedicated class; this allows us to optimize theme interactions,
' without worrying about the details locally.
Private m_Colors As pdThemeColors

Public Function GetControlType() As PD_ControlType
    GetControlType = pdct_CheckBox
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

'State is toggled on each click.  For backwards compatibility reasons, we use VB's built-in checkbox constants, although this control
' really only supports true/false states.
Public Property Get Value() As CheckBoxConstants
Attribute Value.VB_UserMemId = 0
    Value = m_Value
End Property

Public Property Let Value(ByVal newValue As CheckBoxConstants)
    If m_Value <> newValue Then
        m_Value = newValue
        RedrawBackBuffer
        RaiseEvent Click
        PropertyChanged "Value"
    End If
End Property

Private Sub ucSupport_GotFocusAPI()
    RedrawBackBuffer
    RaiseEvent GotFocusAPI
End Sub

'Space and Enter keypresses toggle control state
Private Sub ucSupport_KeyDownCustom(ByVal Shift As ShiftConstants, ByVal vkCode As Long, markEventHandled As Boolean)
    
    markEventHandled = False
    
    If (Me.Enabled And (vkCode = VK_SPACE)) Then
        markEventHandled = True
        If CBool(Me.Value) Then Me.Value = vbUnchecked Else Me.Value = vbChecked
    End If
    
End Sub

Private Sub ucSupport_KeyDownSystem(ByVal Shift As ShiftConstants, ByVal whichSysKey As PD_NavigationKey, markEventHandled As Boolean)
    
    'Enter/Esc get reported directly to the system key handler.  Note that we track the return, because TRUE
    ' means the key was successfully forwarded to the relevant handler.  (If FALSE is returned, no control
    ' accepted the keypress, meaning we should forward the event down the line.)
    markEventHandled = NavKey.NotifyNavKeypress(Me, whichSysKey, Shift)
    
End Sub

Private Sub ucSupport_LostFocusAPI()
    RedrawBackBuffer
    RaiseEvent LostFocusAPI
End Sub

'To improve responsiveness, MouseDown is used instead of Click
Private Sub ucSupport_MouseDownCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal timeStamp As Long)
    If Me.Enabled And IsMouseOverClickArea(x, y) Then
        If CBool(Me.Value) Then Me.Value = vbUnchecked Else Me.Value = vbChecked
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
    If m_MouseInsideClickableRect <> IsMouseOverClickArea(x, y) Then
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
    
    'Initialize a master user control support class
    Set ucSupport = New pdUCSupport
    ucSupport.RegisterControl UserControl.hWnd, True
    
    'Request any control-specific functionality
    ucSupport.RequestExtraFunctionality True, True
    ucSupport.SpecifyRequiredKeys VK_SPACE
    ucSupport.RequestCaptionSupport
    ucSupport.SetCaptionAutomaticPainting False
    
    'Prep the color manager and load default colors
    Set m_Colors = New pdThemeColors
    Dim colorCount As PDCHECKBOX_COLOR_LIST: colorCount = [_Count]
    m_Colors.InitializeColorList "PDCheckBox", colorCount
    If Not MainModule.IsProgramRunning() Then UpdateColorList
    
    'Update the control size parameters at least once
    UpdateControlLayout
                
End Sub

'Set default properties
Private Sub UserControl_InitProperties()
    Caption = "caption"
    FontSize = 10
    Value = vbChecked
End Sub

'At run-time, painting is handled by PD's pdWindowPainter class.  In the IDE, however, we must rely on VB's internal paint event.
Private Sub UserControl_Paint()
    If Not MainModule.IsProgramRunning() Then ucSupport.RequestIDERepaint UserControl.hDC
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        Caption = .ReadProperty("Caption", "")
        FontSize = .ReadProperty("FontSize", 10)
        Value = .ReadProperty("Value", vbChecked)
    End With
End Sub

Private Sub UserControl_Resize()
    If Not MainModule.IsProgramRunning() Then ucSupport.NotifyIDEResize UserControl.Width, UserControl.Height
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "Caption", ucSupport.GetCaptionText, "caption"
        .WriteProperty "FontSize", ucSupport.GetCaptionFontSize, 10
        .WriteProperty "Value", Value, vbChecked
    End With
End Sub

'Whenever the size of the control changes, we must recalculate some internal rendering metrics.
Private Sub UpdateControlLayout()
    
    'Retrieve DPI-aware control dimensions from the support class
    Dim bWidth As Long, bHeight As Long
    bWidth = ucSupport.GetBackBufferWidth
    bHeight = ucSupport.GetBackBufferHeight
    
    'A little bit of horizontal and vertical padding is applied around various control parts
    Const hTextPadding As Long = 1&, vTextPadding As Long = 1&, hBoxCaptionPadding As Long = 8&
    
    'Next, determine the positioning of the caption, if present.  (ucSupport.GetCaptionBottom tells us where the
    ' caption text ends vertically.)
    If ucSupport.IsCaptionActive Then
        
        'Start by making sure the control is tall enough to fit the caption.  (Control height is auto-controlled at present.)
        If ucSupport.GetCaptionHeight(False) + FixDPI(vTextPadding) * 2 <> bHeight Then
            bHeight = ucSupport.GetCaptionHeight(False) + FixDPI(vTextPadding) * 2
            ucSupport.RequestNewSize bWidth, bHeight, False
        End If
        
    End If
    
    'Because the checkbox size and font size are inextricably connected, we now need to retrieve a font object matching the
    ' current control font size.  That font's metrics will determine how everything gets positioned.
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
    Dim offsetX As Long, offsetY As Long, chkBoxSize As Long
    offsetX = 1 + FixDPI(2)
    offsetY = 1 + FixDPI(2)
    chkBoxSize = bHeight - (offsetY * 2)
    
    'Use that to populate the checkbox rect
    With m_CheckboxRect
        .Left = offsetX
        .Top = offsetY
        .Width = chkBoxSize
        .Height = chkBoxSize
    End With
    
    'Pass the available space to the support class; it needs this information to auto-fit the caption
    Dim captionLeft As Long
    captionLeft = (m_CheckboxRect.Left + m_CheckboxRect.Width) + FixDPI(hBoxCaptionPadding)
    ucSupport.SetCaptionCustomPosition captionLeft, 0, bWidth - captionLeft, bHeight
    
    'While here, calculate a caption rect, taking into account the auto-sized caption text (which may be using a different font size)
    With m_CaptionRect
        .Left = captionLeft
        .Top = (bHeight - ucSupport.GetCaptionHeight(True)) / 2
        .Width = ucSupport.GetCaptionWidth(True) + 1
        If .Left + .Width > bWidth Then .Width = (bWidth - .Left)
        .Height = ucSupport.GetCaptionHeight(True) + 1
    End With
    
    'The clickable rect is the union of the checkbox and caption rect.  Calculate it now.
    PDMath.UnionRectF m_ClickableRect, m_CheckboxRect, m_CaptionRect
    
    'If the caption still does not fit within the available area (typically because we reached the minimum allowable font
    ' size, but the caption was *still* too long), set a module-level failure state to TRUE.  This notifies the renderer
    ' that ellipses must be forcibly appended to the caption.
    If ucSupport.GetCaptionWidth(True) > bWidth - m_CaptionRect.Left Then
        m_FitFailure = True
    Else
        m_FitFailure = False
    End If
    
    RedrawBackBuffer
            
End Sub

'Use this function to completely redraw the back buffer from scratch.  Note that this is computationally expensive compared to just flipping the
' existing buffer to the screen, so only redraw the backbuffer if the control state has somehow changed.
Private Sub RedrawBackBuffer()
    
    'Request the back buffer DC, and ask the support module to erase any existing rendering for us.
    Dim bufferDC As Long
    bufferDC = ucSupport.GetBackBufferDC(True, m_Colors.RetrieveColor(PDCB_Background, Me.Enabled))
    
    Dim bWidth As Long, bHeight As Long
    bWidth = ucSupport.GetBackBufferWidth
    bHeight = ucSupport.GetBackBufferHeight
    
    'Populate colors from the master theme object
    Dim chkBoxColorBorder As Long, chkBoxColorFill As Long, chkColor As Long, txtColor As Long
    chkBoxColorBorder = m_Colors.RetrieveColor(PDCB_ButtonBorder, Me.Enabled, m_Value, m_MouseInsideClickableRect)
    chkBoxColorFill = m_Colors.RetrieveColor(PDCB_ButtonFill, Me.Enabled, m_Value, m_MouseInsideClickableRect)
    chkColor = m_Colors.RetrieveColor(PDCB_Checkmark, Me.Enabled, m_Value, m_MouseInsideClickableRect)
    txtColor = m_Colors.RetrieveColor(PDCB_Caption, Me.Enabled, m_Value, m_MouseInsideClickableRect)
    
    If MainModule.IsProgramRunning() Then
        
        'Fill the checkbox area
        With m_CheckboxRect
            GDI_Plus.GDIPlusFillRectToDC bufferDC, Int(.Left), Int(.Top), Int(.Width + 1.999), Int(.Height + 1.999), chkBoxColorFill
        End With
        
        'If the check box button is checked, draw a checkmark inside the border.  The checkmark is defined by three points,
        ' defined in LTR order
        If CBool(m_Value) Then
            Dim pt1 As PointFloat, pt2 As PointFloat, pt3 As PointFloat
            pt1.x = m_CheckboxRect.Left + FixDPIFloat(3)
            pt1.y = m_CheckboxRect.Top + (m_CheckboxRect.Height / 2)
            pt2.x = m_CheckboxRect.Left + (m_CheckboxRect.Width / 2) - FixDPIFloat(1.25)
            pt2.y = m_CheckboxRect.Top + m_CheckboxRect.Height - FixDPIFloat(3)
            pt3.x = (m_CheckboxRect.Left + m_CheckboxRect.Width) - FixDPIFloat(2)
            pt3.y = m_CheckboxRect.Top + FixDPIFloat(3)
            GDI_Plus.GDIPlusDrawLineToDC bufferDC, pt1.x, pt1.y, pt2.x, pt2.y, chkColor, 255, FixDPI(2), True, GP_LC_Round, True
            GDI_Plus.GDIPlusDrawLineToDC bufferDC, pt2.x, pt2.y, pt3.x, pt3.y, chkColor, 255, FixDPI(2), True, GP_LC_Round, True
        End If
        
        'Draw the checkbox border.  (Note that it has variable width, contingent on MouseOver status.)
        Dim borderWidth As Single
        If m_MouseInsideClickableRect Then borderWidth = 3 Else borderWidth = 1
        GDI_Plus.GDIPlusDrawRectFOutlineToDC bufferDC, m_CheckboxRect, chkBoxColorBorder, , borderWidth, , GP_LJ_Miter
        
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
        .LoadThemeColor PDCB_Background, "Background", IDE_WHITE
        .LoadThemeColor PDCB_Caption, "Caption", IDE_GRAY
        .LoadThemeColor PDCB_ButtonFill, "ButtonFill", IDE_WHITE
        .LoadThemeColor PDCB_ButtonBorder, "ButtonBorder", IDE_BLACK
        .LoadThemeColor PDCB_Checkmark, "Checkmark", IDE_BLUE
    End With
End Sub

'External functions can call this to request a redraw.  This is helpful for live-updating theme settings, as in the Preferences dialog.
Public Sub UpdateAgainstCurrentTheme(Optional ByVal hostFormhWnd As Long = 0)
    If ucSupport.ThemeUpdateRequired Then
        UpdateColorList
        If MainModule.IsProgramRunning() Then NavKey.NotifyControlLoad Me, hostFormhWnd
        If MainModule.IsProgramRunning() Then ucSupport.UpdateAgainstThemeAndLanguage
    End If
End Sub

'By design, PD prefers to not use design-time tooltips.  Apply tooltips at run-time, using this function.
' (IMPORTANT NOTE: translations are handled automatically.  Always pass the original English text!)
Public Sub AssignTooltip(ByVal newTooltip As String, Optional ByVal newTooltipTitle As String, Optional ByVal newTooltipIcon As TT_ICON_TYPE = TTI_NONE)
    ucSupport.AssignTooltip UserControl.ContainerHwnd, newTooltip, newTooltipTitle, newTooltipIcon
End Sub
