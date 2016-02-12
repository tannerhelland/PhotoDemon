VERSION 5.00
Begin VB.UserControl pdSliderStandalone 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5055
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
   ScaleHeight     =   20
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   337
End
Attribute VB_Name = "pdSliderStandalone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon standalone Slider custom control
'Copyright 2013-2016 by Tanner Helland
'Created: 19/April/13
'Last updated: 11/February/16
'Last update: split slider functionality into its own control, which is kinda the best separate we can get in VB6 :/
'
'In PD, this control is never used on its own.  It's only a component of the pdSlider control (which also attaches
' a spinner), and it's split out here in an attempt to simplify its rendering code and input handling, which are
' fairly specialized.
'
'For implementation details, refer to pdSlider.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'This object provides a single raised event:
' - Change (which triggers when the slider moves in any direction)
Public Event Change()

'Because VB focus events are wonky, especially when we use CreateWindow within a UC, this control raises its own
' specialized focus events.  If you need to track focus, use these instead of the default VB functions.
Public Event GotFocusAPI()
Public Event LostFocusAPI()

'Used to internally track value, min, and max values as floating-points
Private m_Value As Double, m_Min As Double, m_Max As Double

'The number of significant digits for this control.  0 means integer values.
Private m_significantDigits As Long

'When the mouse is down on the slider, these values will be updated accordingly
Private m_MouseDown As Boolean
Private m_InitX As Single, m_InitY As Single

'Track and slider diameter, at 96 DPI.  Note that the actual render and hit detection functions will adjust these constants for
' the current screen DPI.
Private Const TRACK_DIAMETER As Long = 6
Private Const SLIDER_DIAMETER As Long = 16

'Track and slider diameter, at current DPI.  This is set when the control is first loaded.  In a perfect world, we would catch screen
' DPI changes and update these values accordingly, but I'm postponing that project until a later date.
Private m_TrackDiameter As Single, m_SliderDiameter As Single

'Width/height of the full slider area.  These are set at control intialization, and will only be updated if the control size changes.
' As ScaleWidth and ScaleHeight properties can be slow to read, we cache these values manually.
Private m_SliderAreaWidth As Long, m_SliderAreaHeight As Long

'Background track style.  This can be changed at run-time or design-time, and it will (obviously) affect the way the background
' track is rendered.  For the custom-drawn method, the owner must supply their own DIB for the background area.  Note that the control
' will automatically crop the supplied DIB to the rounded-rect shape required by the track, so the owner need only supply a stock
' rectangular DIB.
Public Enum SLIDER_TRACK_STYLE
    DefaultStyle = 0
    NoFrills = 1
    GradientTwoPoint = 2
    GradientThreePoint = 3
    HueSpectrum360 = 4
    CustomOwnerDrawn = 5
End Enum

#If False Then
    Const DefaultStyle = 0, NoFrills = 1, GradientTwoPoint = 2, GradientThreePoint = 3, HueSpectrum360 = 4, CustomOwnerDrawn = 5
#End If

Private m_SliderStyle As SLIDER_TRACK_STYLE

'Gradient colors.  For the two-color gradient style, only colors Left and Right are relevant.  Color Middle is used for the
' 3-color style only, and note that it *must* be accompanied by an owner-supplied middle position value.
Private m_GradientColorLeft As Long, m_GradientColorRight As Long, m_GradientColorMiddle As Long
Private m_GradientMiddleValue As Double

'Notch positioning.  This can be changed at run-time or design-time, and it will (obviously) affect where the "zero-position" notch
' appears.  When "Automatic" is selected, PD will automatically set the notch to one of two places: 0 (if 0 is a selectable position),
' or the control's minimum value.  For some controls, no notch may be wanted - in this case, use the "none" style.  Finally, a custom
' position may be required for some tools, like Gamma, where the default value isn't obvious (1.0 in that case), or the Opacity slider,
' where the default is 100, not 0.
Public Enum SLIDER_NOTCH_POSITION
    AutomaticPosition = 0
    DoNotDisplayNotch = 1
    CustomPosition = 2
End Enum

#If False Then
    Const AutomaticPosition = 0, DoNotDisplayNotch = 1, CustomPosition = 2
#End If

'Current notch positioning.  If CustomPosition is set, the corresponding NotchCustomValue will be used.
Private m_NotchPosition As SLIDER_NOTCH_POSITION
Private m_CustomNotchValue As Double

'When the slider track is drawn, this rect will be filled with its relevant coordinates.  We use this to track Mouse_Over behavior,
' so we can change the cursor to a hand.
Private m_SliderTrackRect As RECTF

'Internal gradient DIB.  This is recreated as necessary to reflect the gradient colors and positions.
Private m_GradientDIB As pdDIB

'Full slider background DIB, with gradient, outline, notch (if any).  The only thing missing is the slider knob and the highlighted
' portion of the track, both of which are paintedin a separate step (as they are the ones most likely to require changes!)
Private m_SliderBackgroundDIB As pdDIB

'For UI purposes, we track whether the mouse is over the slider, or the background track.  Note that these are
' mutually exclusive; if the mouse is over the slider, it will *not* be marked as over the background track.
Private m_MouseOverSlider As Boolean, m_MouseOverSliderTrack As Boolean, m_MouseTrackX As Single

'To optimize run-time performance, this class doesn't raise Change() events unless the control's value actually changes.
' However, we need to ensure that at least one Change() event is raised, so that initialization steps in a parent object
' can fire if necessary.  We use this tracker to guarantee that at least one Change() event is fired at initialization.
Private m_FirstChangeEvent As Boolean

'User control support class.  Historically, many classes (and associated subclassers) were required by each user control,
' but I've since attempted to wrap these into a single master control support class.
Private WithEvents ucSupport As pdUCSupport
Attribute ucSupport.VB_VarHelpID = -1

'Local list of themable colors.  This list includes all potential colors used by this class, regardless of state change
' or internal control settings.  The list is updated by calling the UpdateColorList function.
' (Note also that this list does not include variants, e.g. "BorderColor" vs "BorderColor_Hovered".  Variant values are
'  automatically calculated by the color management class, and they are retrieved by passing boolean modifiers to that
'  class, rather than treating every imaginable variant as a separate constant.)
Private Enum PDSLIDERSTANDALONE_COLOR_LIST
    [_First] = 0
    PDSS_Background = 0
    PDSS_TrackBack = 1
    PDSS_TrackFill = 2
    PDSS_TrackJumpIndicator = 3
    PDSS_ThumbFill = 4
    PDSS_ThumbBorder = 5
    PDSS_Notch = 6
    [_Last] = 6
    [_Count] = 7
End Enum

'Color retrieval and storage is handled by a dedicated class; this allows us to optimize theme interactions,
' without worrying about the details locally.
Private m_Colors As pdThemeColors

'The Enabled property is a bit unique; see http://msdn.microsoft.com/en-us/library/aa261357%28v=vs.60%29.aspx
Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal newValue As Boolean)
    UserControl.Enabled = newValue
    RedrawBackBuffer
    PropertyChanged "Enabled"
End Property

'Gradient colors.  For the two-color gradient style, only colors Left and Right are relevant.  Color Middle is used for the
' 3-color style only, and note that it *must* be accompanied by an owner-supplied middle position value.
Public Property Get GradientColorLeft() As OLE_COLOR
    GradientColorLeft = m_GradientColorLeft
End Property

Public Property Let GradientColorLeft(ByVal newColor As OLE_COLOR)
    newColor = ConvertSystemColor(newColor)
    If newColor <> m_GradientColorLeft Then
        m_GradientColorLeft = newColor
        CreateGradientTrack
        RedrawSlider
        PropertyChanged "GradientColorLeft"
    End If
End Property

Public Property Get GradientColorMiddle() As OLE_COLOR
    GradientColorMiddle = m_GradientColorMiddle
End Property

Public Property Let GradientColorMiddle(ByVal newColor As OLE_COLOR)
    newColor = ConvertSystemColor(newColor)
    If newColor <> m_GradientColorMiddle Then
        m_GradientColorMiddle = newColor
        CreateGradientTrack
        RedrawSlider
        PropertyChanged "GradientColorMiddle"
    End If
End Property

Public Property Get GradientColorRight() As OLE_COLOR
    GradientColorRight = m_GradientColorRight
End Property

Public Property Let GradientColorRight(ByVal newColor As OLE_COLOR)
    newColor = ConvertSystemColor(newColor)
    If newColor <> m_GradientColorRight Then
        m_GradientColorRight = newColor
        CreateGradientTrack
        RedrawSlider
        PropertyChanged "GradientColorRight"
    End If
End Property

'Custom middle value for the 3-color gradient style.  This value is ignored for all other styles.
Public Property Get GradientMiddleValue() As Double
    GradientMiddleValue = m_GradientMiddleValue
End Property

Public Property Let GradientMiddleValue(ByVal newValue As Double)
    If newValue <> m_GradientMiddleValue Then
        m_GradientMiddleValue = newValue
        CreateGradientTrack
        RedrawSlider
        PropertyChanged "GradientMiddleValue"
    End If
End Property

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

Public Property Get Max() As Double
    Max = m_Max
End Property

Public Property Let Max(ByVal newValue As Double)
    
    m_Max = newValue
    If m_Value > m_Max Then Value = m_Max
    
    'If the background track style has a custom appearance (like a gradient), changing the maximum value potentially
    ' alters its appearance.  We have no choice but to recreate that background track image now.
    If (m_SliderStyle = GradientTwoPoint) Or (m_SliderStyle = GradientThreePoint) Then CreateGradientTrack
    
    RedrawSlider
    PropertyChanged "Max"
    
End Property

Public Property Get Min() As Double
    Min = m_Min
End Property

Public Property Let Min(ByVal newValue As Double)
    
    m_Min = newValue
    If m_Value < m_Min Then Value = m_Min
    
    'If the background track style has a custom appearance (like a gradient), changing the maximum value potentially
    ' alters its appearance.  We have no choice but to recreate that background track image now.
    If (m_SliderStyle = GradientTwoPoint) Or (m_SliderStyle = GradientThreePoint) Then CreateGradientTrack
    
    RedrawSlider
    PropertyChanged "Min"
    
End Property

'Notch positioning technique.  If CUSTOM is set, make sure to supply a custom value to match!
Public Property Get NotchPosition() As SLIDER_NOTCH_POSITION
    NotchPosition = m_NotchPosition
End Property

Public Property Let NotchPosition(ByVal newPosition As SLIDER_NOTCH_POSITION)
    m_NotchPosition = newPosition
    RedrawSlider
    PropertyChanged "NotchPosition"
End Property

'Custom notch value.  This value is only used if NotchPosition = CustomPosition.
Public Property Get NotchValueCustom() As Double
    NotchValueCustom = m_CustomNotchValue
End Property

Public Property Let NotchValueCustom(ByVal newValue As Double)
    m_CustomNotchValue = newValue
    RedrawSlider
    PropertyChanged "NotchValueCustom"
End Property

'Significant digits determines whether the control allows float values or int values (and with how much precision).
' Because the slider's position is locked to allowable values, this setting requires a redraw, so try to limit how frequently
' you modify it.
Public Property Get SigDigits() As Long
    SigDigits = m_significantDigits
End Property

Public Property Let SigDigits(ByVal newValue As Long)
    m_significantDigits = newValue
    PropertyChanged "SigDigits"
End Property

'Track style has no mechanical bearing on the control - it only affects visual appearance.  As such, the correctness of its
' behavior is not guaranteed if you change this setting at run-time.
Public Property Get SliderTrackStyle() As SLIDER_TRACK_STYLE
    SliderTrackStyle = m_SliderStyle
End Property

Public Property Let SliderTrackStyle(ByVal newStyle As SLIDER_TRACK_STYLE)
    m_SliderStyle = newStyle
    RedrawSlider
    PropertyChanged "SliderTrackStyle"
End Property

Public Property Get Value() As Double
    Value = m_Value
End Property

Public Property Let Value(ByVal newValue As Double)
    
    'Don't make any changes unless the new value deviates from the existing one, OR this is the first time a Value has been assigned
    If (newValue <> m_Value) Or m_FirstChangeEvent Then
        
        m_FirstChangeEvent = False
        m_Value = newValue
        
        'This control handles bound-checking differently from most common controls.  Out-of-bound value requests are silently
        ' forced in-bounds.  This is by design, and the behavior cannot be modified by the caller.
        If m_Value < m_Min Then m_Value = m_Min
        If m_Value > m_Max Then m_Value = m_Max
                
        'Because we support subpixel positioning for the slider, value changes always require a redraw, even if the slider's
        ' position only changes by a miniscule amount
        RedrawBackBuffer
        
        If Me.Enabled Then RaiseEvent Change
        PropertyChanged "Value"
        
    End If
    
End Property

'This function serves two purposes: most of the time, we use it for hit-detection against the track slider, but some functions
' also use it to check hit-detection against the underlying track, which allows for "jump to position" behavior.
Private Function IsMouseOverSlider(ByVal mouseX As Single, ByVal mouseY As Single, Optional ByVal alsoCheckBackgroundTrack As Boolean = True) As Boolean

    'Retrieve the current x/y position of the slider's CENTER
    Dim sliderX As Single, sliderY As Single
    GetSliderCoordinates sliderX, sliderY
    
    'See if the mouse is within distance of the slider's center
    If distanceTwoPoints(sliderX, sliderY, mouseX, mouseY) < (FixDPIFloat(SLIDER_DIAMETER) / 2) Then
        IsMouseOverSlider = True
    Else
        
        'If the mouse is not over the slider itself, check the background track as well
        If IsPointInRectF(mouseX, mouseY, m_SliderTrackRect) And alsoCheckBackgroundTrack Then
            IsMouseOverSlider = True
        Else
            IsMouseOverSlider = False
        End If
    End If

End Function

Private Sub UpdateMousePosition(ByVal mouseX As Single, ByVal mouseY As Single)
    'Update mouse tracker:
    'm_MouseInsideBrushRect = Math_Functions.IsPointInRectF(mouseX, mouseY, m_BrushRect)
    
    'Contingent on the tracker value, update cursor values to match
    'If m_MouseInsideBrushRect Then ucSupport.RequestCursor IDC_HAND Else ucSupport.RequestCursor IDC_DEFAULT
End Sub

Private Sub ucSupport_GotFocusAPI()
    RaiseEvent GotFocusAPI
End Sub

Private Sub ucSupport_LostFocusAPI()
    RaiseEvent LostFocusAPI
End Sub

'Up and right arrows are used to increment the slider value, while left and down arrows decrement it
Private Sub ucSupport_KeyDownCustom(ByVal Shift As ShiftConstants, ByVal vkCode As Long, markEventHandled As Boolean)
    If (vkCode = VK_UP) Or (vkCode = VK_RIGHT) Then Value = Value + GetIncrementAmount
    If (vkCode = VK_LEFT) Or (vkCode = VK_DOWN) Then Value = Value - GetIncrementAmount
End Sub

Private Sub ucSupport_MouseDownCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    
    If ((Button And pdLeftButton) <> 0) Then
    
        'Check to see if the mouse is over a) the slider control button, or b) the background track.  Both are valid.
        If IsMouseOverSlider(x, y) Then
            
            m_MouseDown = True
            m_MouseOverSlider = True
            m_MouseOverSliderTrack = False
            
            'Calculate a new control value.  This will cause the slider to "jump" to a slightly modified position,
            ' if positions are restricted by some combination of the total range and significant digit allowance.
            Value = (m_Max - m_Min) * ((x - GetTrackLeft) / (GetTrackRight - GetTrackLeft)) + m_Min
            
            'Retrieve the current slider x/y values, and store the mouse position relative to those values
            Dim sliderX As Single, sliderY As Single
            GetSliderCoordinates sliderX, sliderY
            m_InitX = x - sliderX
            m_InitY = y - sliderY
            
            'Force an immediate redraw (instead of waiting for WM_PAINT to process); this makes the control feel more responsive
            ucSupport.RequestRepaint True
            
        End If
            
    End If
    
End Sub

'Because this control supports quite a few different hover behaviors, we may need to redraw the control upon MouseLeave
Private Sub ucSupport_MouseLeave(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    m_MouseOverSlider = False
    m_MouseOverSliderTrack = False
    RedrawSlider
    ucSupport.RequestCursor IDC_DEFAULT
End Sub

Private Sub ucSupport_MouseMoveCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    
    If m_MouseDown Then
        
        'Because the slider tracks the mouse's motion, we automatically assume the mouse is over it
        ' (and not the background track).
        m_MouseOverSlider = True
        m_MouseOverSliderTrack = False
        
        'Calculate a new control value.  This will cause the slider to "jump" to the current position, if positions
        ' are restricted by some combination of the total range and significant digit allowance.
        Value = (m_Max - m_Min) * (((x + m_InitX) - GetTrackLeft) / (GetTrackRight - GetTrackLeft)) + m_Min
        
        'Force an immediate redraw (instead of waiting for WM_PAINT to process)
        ucSupport.RequestRepaint True
        
    'If the LMB is *not* down, modify the cursor according to its position relative to the slider and/or track
    Else
        
        m_MouseOverSlider = IsMouseOverSlider(x, y, False)
        If m_MouseOverSlider Then
            m_MouseOverSliderTrack = False
        Else
            m_MouseOverSliderTrack = IsMouseOverSlider(x, y, True)
            m_MouseTrackX = x
        End If
        
        If m_MouseOverSlider Or m_MouseOverSliderTrack Then
            ucSupport.RequestCursor IDC_HAND
        Else
            ucSupport.RequestCursor IDC_ARROW
        End If
        
        'Because hover effects may be rendered on multiple parts of the control, we need redraw the button on MouseMove events
        RedrawSlider
    
    End If
    
End Sub

'When the mouse button is released, we always perform a final MouseMove update at the last reported x/y position.
' If intensive processing occurred while the slider was being used, this ensures that the mouse location at its
' exact point of release is correctly rendered.
Private Sub ucSupport_MouseUpCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal ClickEventAlsoFiring As Boolean)
    If ((Button And pdLeftButton) <> 0) And m_MouseDown Then
        Value = (m_Max - m_Min) * (((x + m_InitX) - GetTrackLeft) / (GetTrackRight - GetTrackLeft)) + m_Min
        m_MouseDown = False
    End If
End Sub

Private Sub ucSupport_RepaintRequired(ByVal updateLayoutToo As Boolean)
    If updateLayoutToo Then UpdateControlLayout
    RedrawSlider
End Sub

Private Sub ucSupport_WindowResize(ByVal newWidth As Long, ByVal newHeight As Long)
    UpdateControlLayout
End Sub

Private Sub UserControl_Initialize()

    'Initialize a master user control support class
    Set ucSupport = New pdUCSupport
    ucSupport.RegisterControl UserControl.hWnd
    ucSupport.RequestExtraFunctionality True, True
    ucSupport.SpecifyRequiredKeys VK_UP, VK_RIGHT, VK_DOWN, VK_LEFT
    
    'Prep the color manager and load default colors
    Set m_Colors = New pdThemeColors
    Dim colorCount As PDSLIDERSTANDALONE_COLOR_LIST: colorCount = [_Count]
    m_Colors.InitializeColorList "PDSliderStandalone", colorCount
    If Not g_IsProgramRunning Then UpdateColorList
    
    'Update the control-level track and slider diameters to reflect current screen DPI
    m_TrackDiameter = FixDPI(TRACK_DIAMETER)
    m_SliderDiameter = FixDPI(SLIDER_DIAMETER)
    
    'Set slider area width/height
    m_SliderAreaWidth = ucSupport.GetControlWidth
    m_SliderAreaHeight = ucSupport.GetControlHeight
    
    'Guarantee that at least one Change() event gets fired before duplicates start being tracked
    m_FirstChangeEvent = True
    
    'Initialize various back buffers and background DIBs
    Set m_SliderBackgroundDIB = New pdDIB
    
    'Update the control size parameters at least once
    UpdateControlLayout

End Sub

Private Sub UserControl_InitProperties()
    Value = 0
    Min = 0
    Max = 10
    SigDigits = 0
    SliderTrackStyle = DefaultStyle
    
    'These default gradient values are useless; if you're using a gradient style, MAKE CERTAIN TO SPECIFY ACTUAL COLORS!
    GradientColorLeft = RGB(0, 0, 0)
    GradientColorRight = RGB(255, 255, 25)
    GradientColorMiddle = RGB(121, 131, 135)
    
    'This default gradient middle value is useless; if you use the 3-color gradient style, MAKE CERTAIN TO SPECIFY THIS VALUE!
    GradientMiddleValue = 0
    
    'Default notch position; for most controls, it should be set to AUTOMATIC.  If CUSTOM is set, make sure to supply
    ' the custom value you want in the corresponding property!
    NotchPosition = AutomaticPosition
    NotchValueCustom = 0
End Sub

'At run-time, painting is handled by the support class.  In the IDE, however, we must rely on VB's internal paint event.
Private Sub UserControl_Paint()
    ucSupport.RequestIDERepaint UserControl.hDC
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        SigDigits = .ReadProperty("SigDigits", 0)
        Max = .ReadProperty("Max", 10)
        Min = .ReadProperty("Min", 0)
        Value = .ReadProperty("Value", 0)
        GradientColorLeft = .ReadProperty("GradientColorLeft", RGB(0, 0, 0))
        GradientColorRight = .ReadProperty("GradientColorRight", RGB(255, 255, 255))
        GradientColorMiddle = .ReadProperty("GradientColorMiddle", RGB(121, 131, 135))
        GradientMiddleValue = .ReadProperty("GradientMiddleValue", 0)
        NotchPosition = .ReadProperty("NotchPosition", 0)
        NotchValueCustom = .ReadProperty("NotchValueCustom", 0)
        SliderTrackStyle = .ReadProperty("SliderTrackStyle", DefaultStyle)
    End With
End Sub

Private Sub UserControl_Resize()
    If Not g_IsProgramRunning Then ucSupport.RequestRepaint True
End Sub

'If the track style is some kind of custom gradient, make sure our internal gradient backdrop is valid before the control
' is shown for the first time.
Private Sub UserControl_Show()
    UpdateControlLayout
    RedrawSlider
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "Min", m_Min, 0
        .WriteProperty "Max", m_Max, 10
        .WriteProperty "SigDigits", m_significantDigits, 0
        .WriteProperty "Value", m_Value, 0
        .WriteProperty "GradientColorLeft", m_GradientColorLeft, RGB(0, 0, 0)
        .WriteProperty "GradientColorRight", m_GradientColorRight, RGB(255, 255, 255)
        .WriteProperty "GradientColorMiddle", m_GradientColorMiddle, RGB(121, 131, 135)
        .WriteProperty "GradientMiddleValue", m_GradientMiddleValue, 0
        .WriteProperty "NotchPosition", m_NotchPosition, 0
        .WriteProperty "NotchValueCustom", m_CustomNotchValue, 0
        .WriteProperty "SliderTrackStyle", m_SliderStyle, DefaultStyle
    End With
End Sub

'Whenever a control property changes that affects control size or layout, call this function to recalculate the control's
' internal layout.
Private Sub UpdateControlLayout()
    
    'Retrieve DPI-aware control dimensions from the support class
    m_SliderAreaWidth = ucSupport.GetBackBufferWidth
    m_SliderAreaHeight = ucSupport.GetBackBufferHeight
    
    'Redraw any custom background images, followed by the whole slider
    If ((m_SliderStyle = GradientTwoPoint) Or (m_SliderStyle = GradientThreePoint) Or (m_SliderStyle = HueSpectrum360)) Then CreateGradientTrack
    RedrawSlider
            
End Sub

'Render a custom slider to the slider area picture box.  Note that the background gradient, if any, should already have been created
' in a separate CreateGradientTrack request.
Private Sub RedrawSlider(Optional ByVal refreshImmediately As Boolean = False)

    'Drawing is done in several stages.  The bulk of the slider is rendered to a persistent slider-only DIB, which contains everything
    ' but the knob and "highlighted" portion of the track.  These are rendered in a separate step, as they are the most common update
    ' required, and we can shortcut the process by not redrawing the full slider on every update.
    
    'Start by retrieving the colors necessary to render various display elements
    Dim backgroundColor As Long, trackColor As Long
    backgroundColor = m_Colors.RetrieveColor(PDSS_Background, Me.Enabled, False, m_MouseOverSlider Or m_MouseOverSliderTrack)
    trackColor = m_Colors.RetrieveColor(PDSS_TrackBack, Me.Enabled, False, m_MouseOverSlider Or m_MouseOverSliderTrack)
    
    'Initialize or repaint the background DIB, as necessary
    m_SliderAreaWidth = ucSupport.GetBackBufferWidth
    m_SliderAreaHeight = ucSupport.GetBackBufferHeight
    If (m_SliderBackgroundDIB.getDIBWidth <> m_SliderAreaWidth) Or (m_SliderBackgroundDIB.getDIBHeight <> m_SliderAreaHeight) Then
        m_SliderBackgroundDIB.createBlank m_SliderAreaWidth, m_SliderAreaHeight, 32, backgroundColor, 255
    Else
        If g_IsProgramRunning Then GDI_Plus.GDIPlusFillDIBRect m_SliderBackgroundDIB, 0, 0, m_SliderAreaWidth, m_SliderAreaHeight, backgroundColor, 255
    End If
        
    'There are a few components to the slider:
    ' 1) The track that sits behind the slider.  It has two relevant parameters: a radius, and a color.  Its width is automatically
    '     calculated relevant to the width of the control as a whole.
    ' 2) The slider knob that sits atop the track.  It has three relevant parameters: a radius, a fill color, and an edge color.
    '     Its width is constant from a programmatic standpoint, though it does get updated at run-time to account for screen DPI.
    
    'We are going to assemble part (1) in this step.
    
    'We always start with the default style: a gray track with rounded edges
    GDI_Plus.GDIPlusDrawLineToDC m_SliderBackgroundDIB.getDIBDC, GetTrackLeft, m_SliderAreaHeight \ 2, GetTrackRight, m_SliderAreaHeight \ 2, trackColor, 255, m_TrackDiameter + 1, True, LineCapRound
    
    If Me.Enabled Then
    
        'This control supports a variety of specialty track styles.  Some of these styles require a DIB supplied by the owner -
        ' note that they *will not* render properly until that DIB is provided!
        Select Case m_SliderStyle
            
            'Gradient styles.  There are a variety of these, and once they've been created, they are all rendered identically.
            '(Basically, we draw the gradient effect DIB to the location where we'd normally draw the track line.  Alpha has already
            ' been calculated for the gradient DIB, so it will sit precisely inside the trackline drawn above, giving the track a
            ' sharp 1px border.)
            Case GradientTwoPoint, GradientThreePoint, HueSpectrum360
                If m_GradientDIB Is Nothing Then CreateGradientTrack
                m_GradientDIB.alphaBlendToDC m_SliderBackgroundDIB.getDIBDC, 255, GetTrackLeft - (m_TrackDiameter \ 2), 0
            
            'In the future, we may support fully owner-drawn sliders, but this is not currently implemented.
            Case CustomOwnerDrawn
        
        End Select
        
        'While the control is enabled, we also draw a slight notch above and below the slider track at the "default" position.
        ' (This position can be user-controlled, so rendering is somewhat complicated.)
        DrawNotchToDIB m_SliderBackgroundDIB
        
    End If
    
    'Store the calculated position of the slider background.  Mouse hit detection code can make use of this, so we don't have to
    ' constantly re-calculate it during mouse events.
    With m_SliderTrackRect
        .Left = GetTrackLeft
        .Width = GetTrackRight - .Left
        .Top = (m_SliderAreaHeight / 2) - ((m_TrackDiameter + 1) / 2)
        .Height = m_TrackDiameter + 1
    End With
        
    'The slider background is now ready for action.  As a final step, pass control to the knob renderer function.
    RedrawBackBuffer refreshImmediately
        
End Sub

'Render a slight notch at the specified position on the specified DIB.  Note that this sub WILL automatically convert a custom notch
' value into it's appropriate x-coordinate; the caller is not responsible for that.
Private Sub DrawNotchToDIB(ByRef dstDIB As pdDIB)
    
    'First, see if a notch needs to be drawn.  If the notch mode is "none", exit now.
    If m_NotchPosition = DoNotDisplayNotch Then Exit Sub
    
    Dim notchColor As Long
    notchColor = m_Colors.RetrieveColor(PDSS_Notch, Me.Enabled, False, m_MouseOverSlider Or m_MouseOverSliderTrack)
    
    Dim renderNotchValue As Double
    
    'For controls where the notch would be drawn at the "minimum value" position, I prefer to keep a clean visual style and
    ' not draw a redundant notch (as the filled slider conveys the exact same message).  For such controls, notch display
    ' is automatically overridden.
    Dim overrideNotchDraw As Boolean
    overrideNotchDraw = False
    
    'Next, calculate a notch position as necessary.  If the notch mode is "automatic", this function is responsible for
    ' determining what notch value to use.
    If m_NotchPosition = AutomaticPosition Then
    
        'The automatic position varies according to a few factors.  First, some slider styles have their own notch calculations.
        
        'Three-point gradients always display a notch at the position of the middle gradient point; this is the assumed default
        ' value of the control.
        If m_SliderStyle = GradientThreePoint Then
            renderNotchValue = GradientMiddleValue
        
        'All other slider styles use the same heuristic for automatic notch positioning.  If 0 is available, use it.
        ' Otherwise, use the control's minimum value.
        Else
            
            If (0 > m_Min) And (0 <= m_Max) Then
                renderNotchValue = 0
            Else
                renderNotchValue = m_Min
                
                'To keep sliders visually clean, notches are not drawn unless useful, and notches at the obvious minimum position
                ' serve no purpose - so override the entire notch drawing process.
                overrideNotchDraw = True
            End If
            
        End If
    
    'If the notch position is not "do not display", and also not "automatic", it must be custom.  Custom values are always rendered.
    Else
        renderNotchValue = m_CustomNotchValue
    End If
    
    If Not overrideNotchDraw Then
    
        'Convert our calculated notch *value* into an actual *pixel position* on the track
        Dim customX As Single, customY As Single
        GetCustomValueCoordinates renderNotchValue, customX, customY
        
        'Calculate the height of the notch; this varies by DPI, which is automatically factored into m_trackDiameter
        Dim notchSize As Single
        notchSize = (m_SliderAreaHeight - m_TrackDiameter) \ 2 - 4
        
        'Currently, we draw a detached notch above and below the slider's track
        GDI_Plus.GDIPlusDrawLineToDC dstDIB.getDIBDC, customX, 1, customX, 1 + notchSize, notchColor, 255, 1, True, LineCapFlat
        GDI_Plus.GDIPlusDrawLineToDC dstDIB.getDIBDC, customX, m_SliderAreaHeight - 1, customX, m_SliderAreaHeight - 1 - notchSize, notchColor, 255, 1, True, LineCapFlat
        
    End If
    
End Sub

'Given any arbitrary DIB, resize it to the area of the background track.
' (This is used for custom-drawn sliders, and it should be the first step in assembling the track DIB.)
Public Sub SizeDIBToTrackArea(ByRef targetDIB As pdDIB)
    If targetDIB Is Nothing Then Set targetDIB = New pdDIB
    targetDIB.createBlank (GetTrackRight - GetTrackLeft) + m_TrackDiameter, m_SliderAreaHeight, 32, ConvertSystemColor(vbWindowBackground), 255
End Sub

'When using a two-color or three-color gradient track style, this function can be called to recreate the background track DIB.
' Please note that this process is expensive (as we have to manually calculate per-pixel alpha masking of the final gradient),
' so please only call this function when absolutely necessary.
Private Sub CreateGradientTrack()

    'Recreate the gradient DIB to the size of the background track area
    SizeDIBToTrackArea m_GradientDIB
    
    Dim trackRadius As Single
    trackRadius = (m_TrackDiameter) \ 2
    
    Dim x As Long
    Dim relativeMiddlePosition As Single, tmpY As Single
    
    'Draw the gradient differently depending on the type of gradient
    Select Case m_SliderStyle
    
        'Two-point gradients are the easiest; simply draw a gradient from left color to right color, the full width of the image
        Case GradientTwoPoint
           Drawing.DrawHorizontalGradientToDIB m_GradientDIB, trackRadius, m_GradientDIB.getDIBWidth - trackRadius, m_GradientColorLeft, m_GradientColorRight
        
        'Three-point gradients are more involved; draw a custom blend from left to middle to right, while accounting for the
        ' center point's position (which is variable, and which may change at run-time).
        Case GradientThreePoint
            
            'Calculate a relative pixel position for the supplied gradient middle value
            If (m_GradientMiddleValue >= m_Min) And (m_GradientMiddleValue <= m_Max) Then
                GetCustomValueCoordinates m_GradientMiddleValue, relativeMiddlePosition, tmpY
            Else
                relativeMiddlePosition = GetTrackLeft + ((GetTrackRight - GetTrackLeft) \ 2)
            End If
            
            'Draw two gradients; one each for the left and right of the gradient middle position
            Drawing.DrawHorizontalGradientToDIB m_GradientDIB, trackRadius, relativeMiddlePosition, m_GradientColorLeft, m_GradientColorMiddle
            Drawing.DrawHorizontalGradientToDIB m_GradientDIB, relativeMiddlePosition, m_GradientDIB.getDIBWidth - trackRadius, m_GradientColorMiddle, m_GradientColorRight
            
        'Hue gradients simply draw a full hue spectrum from 0 to 360.
        Case HueSpectrum360
        
            'From left-to-right, draw a full hue range onto the DIB
            Dim hueSpread As Long
            hueSpread = (m_GradientDIB.getDIBWidth - m_TrackDiameter)
            
            Dim tmpR As Double, tmpG As Double, tmpB As Double
            
            For x = 0 To m_GradientDIB.getDIBWidth - 1
                
                If x < trackRadius Then
                    fHSVtoRGB 0, 1, 1, tmpR, tmpG, tmpB
                    GDI_Plus.GDIPlusDrawLineToDC m_GradientDIB.getDIBDC, x, 0, x, m_GradientDIB.getDIBHeight, RGB(tmpR * 255, tmpG * 255, tmpB * 255), 255, 1, False, LineCapFlat
                ElseIf x > (m_GradientDIB.getDIBWidth - trackRadius) Then
                    fHSVtoRGB 1, 1, 1, tmpR, tmpG, tmpB
                    GDI_Plus.GDIPlusDrawLineToDC m_GradientDIB.getDIBDC, x, 0, x, m_GradientDIB.getDIBHeight, RGB(tmpR * 255, tmpG * 255, tmpB * 255), 255, 1, False, LineCapFlat
                Else
                    fHSVtoRGB (x - trackRadius) / hueSpread, 1, 1, tmpR, tmpG, tmpB
                    GDI_Plus.GDIPlusDrawLineToDC m_GradientDIB.getDIBDC, x, 0, x, m_GradientDIB.getDIBHeight, RGB(tmpR * 255, tmpG * 255, tmpB * 255), 255, 1, False, LineCapFlat
                End If
                
            Next x
            
    End Select
    
    'Next, on custom gradients, we need to fill in the DIB just past the gradient on either side; this makes the gradient's
    ' termination point fall directly on the maximum slider position (instead of the edge of the DIB, which would be
    ' incorrect as we need some padding for the rounded edge of the track area).  Note that hue gradients automatically
    ' handle this step during the DIB creation phase.
    If (m_SliderStyle = GradientTwoPoint) Or (m_SliderStyle = GradientThreePoint) Then
    
        For x = 0 To trackRadius
            GDI_Plus.GDIPlusDrawLineToDC m_GradientDIB.getDIBDC, x, 0, x, m_GradientDIB.getDIBHeight, m_GradientColorLeft, 255, 1, False, LineCapFlat
        Next x
        
        For x = m_GradientDIB.getDIBWidth - trackRadius To m_GradientDIB.getDIBWidth
            GDI_Plus.GDIPlusDrawLineToDC m_GradientDIB.getDIBDC, x, 0, x, m_GradientDIB.getDIBHeight, m_GradientColorRight, 255, 1, False, LineCapFlat
        Next x
        
    End If
    
    'Next, we need to crop the track DIB to the shape of the background slider area.  This is a fairly involved operation, as we need to
    ' render a track line slightly smaller than the usual size, then manually apply a per-pixel copy of alpha data from the created line
    ' to the internal DIB.  This will allows us to alpha-blend the custom DIB over a copy of the background line, to retain a small border.
    
    'Start by creating the image we're going to use as our alpha mask.
    Dim alphaMask As pdDIB
    Set alphaMask = New pdDIB
    alphaMask.createBlank m_GradientDIB.getDIBWidth, m_GradientDIB.getDIBHeight, 32, 0, 0
    
    'Next, use GDI+ to render a slightly smaller line than the typical track onto the alpha mask.  GDI+'s antialiasing code will automatically
    ' set the relevant alpha bytes for the region of interest.
    GDI_Plus.GDIPlusDrawLineToDC alphaMask.getDIBDC, trackRadius, m_GradientDIB.getDIBHeight \ 2, m_GradientDIB.getDIBWidth - trackRadius, m_GradientDIB.getDIBHeight \ 2, 0, 255, m_TrackDiameter - 1, True, LineCapRound
    
    'Transfer the alpha from the alpha mask to the gradient DIB itself
    'alphaMask.setAlphaPremultiplication False
    m_GradientDIB.copyAlphaFromExistingDIB alphaMask
    
    'Premultiply the gradient DIB, so we can successfully alpha-blend it later
    m_GradientDIB.SetAlphaPremultiplication True
    
    'The gradient mask is now complete!
    
End Sub

'Retrieve the current hypothetical coordinates of the slider's *center point*.  Note that these are likely to be floating-point,
' unless the control is in integer-only mode - then you will get an integer-only result.
Private Sub GetSliderCoordinates(ByRef sliderX As Single, ByRef sliderY As Single)
    
    'This dumb catch exists for when sliders are first loaded, and their max/min may both be zero.  This causes a divide-by-zero
    ' error in the horizontal slider position calculation, so if that happens, simply set the slider to its minimum position and exit.
    If m_Min <> m_Max Then
        
        'If an integer-only slider is in use, we use a slightly modified formula
        If SigDigits = 0 Then sliderX = (Int(m_Value + 0.5) - m_Min) Else sliderX = (m_Value - m_Min)
        sliderX = GetTrackLeft + (sliderX / (m_Max - m_Min)) * (GetTrackRight - GetTrackLeft)
        
    Else
        sliderX = GetTrackLeft
    End If
    
    If SigDigits = 0 Then sliderY = m_SliderAreaHeight \ 2 Else sliderY = m_SliderAreaHeight / 2
    
End Sub

'Retrieve the current coordinates of any custom value.  Note that the x/y pair returned are the custom value's *center point*.
Private Sub GetCustomValueCoordinates(ByVal customValue As Single, ByRef customX As Single, ByRef customY As Single)
    
    'This dumb catch exists for when sliders are first loaded, and their max/min may both be zero.  This causes a divide-by-zero
    ' error in the horizontal slider position calculation, so if that happens, simply set the slider to its minimum position and exit.
    If m_Min <> m_Max Then
        customX = GetTrackLeft + ((customValue - m_Min) / (m_Max - m_Min)) * (GetTrackRight - GetTrackLeft)
    Else
        customX = GetTrackLeft
    End If
    
    customY = m_SliderAreaHeight \ 2
    
End Sub

'Returns a single increment amount for the current control.  The increment amount varies according to the significant digits setting;
' it can be as high as 1.0, or as low as 0.01.
Private Function GetIncrementAmount() As Double
    GetIncrementAmount = 1 / (10 ^ m_significantDigits)
End Function

'Return the min/max position of the track behind the slider.  This is used for a lot of things: rendering the track, calculating the
' value of the slider during user interactions (by determing the slider position relative to these two values), etc.  The minimum
' position is constant once the control is created, but the max position can change if the control size changes.
Private Function GetTrackLeft() As Long
    GetTrackLeft = m_SliderDiameter \ 2 + 2
End Function

Private Function GetTrackRight() As Long
    GetTrackRight = ucSupport.GetControlWidth - (m_SliderDiameter \ 2) - 2
End Function

'Primary rendering function.  Note that ucSupport handles a number of rendering duties (like maintaining a back buffer for us).
' Also, this step only composites the knob atop the final slider background, before copying the entire assembled image into the
' control's persistent backbuffer.  This means that you *must* have already assembled the basic track components prior to calling
' this function.
Private Sub RedrawBackBuffer(Optional ByVal refreshImmediately As Boolean = False)
    
    'We can improve shutdown performance by ignoring redraw requests
    If g_ProgramShuttingDown Then
        If (g_Themer Is Nothing) Then Exit Sub
    End If
    
    'Request the back buffer DC, and ask the support module to erase any existing rendering for us.
    Dim backgroundColor As Long, bufferDC As Long
    backgroundColor = m_Colors.RetrieveColor(PDSS_Background, Me.Enabled, m_MouseDown, m_MouseOverSlider Or m_MouseOverSliderTrack)
    bufferDC = ucSupport.GetBackBufferDC(True, backgroundColor)
    
    'Copy the previously assembled track onto the back buffer.  (This is faster than AlphaBlending the result, especially because
    ' we don't need any blending.)
    BitBlt bufferDC, 0, 0, m_SliderAreaWidth, m_SliderAreaHeight, m_SliderBackgroundDIB.getDIBDC, 0, 0, vbSrcCopy
    
    If Me.Enabled And g_IsProgramRunning Then
        
        Dim trackHighlightColor As Long, trackJumpIndicatorColor As Long
        Dim thumbFillColor As Long, thumbBorderColor As Long
        trackHighlightColor = m_Colors.RetrieveColor(PDSS_TrackFill, True, m_MouseDown, m_MouseOverSlider)
        trackJumpIndicatorColor = m_Colors.RetrieveColor(PDSS_TrackJumpIndicator, True, m_MouseDown, m_MouseOverSliderTrack)
        thumbFillColor = m_Colors.RetrieveColor(PDSS_ThumbFill, True, m_MouseDown, m_MouseOverSlider)
        thumbBorderColor = m_Colors.RetrieveColor(PDSS_ThumbBorder, True, m_MouseDown, m_MouseOverSlider)
        
        'Retrieve the current slider x/y position.  (Floating-point values are used for sub-pixel positioning.)
        Dim relevantSliderPosX As Single, relevantSliderPosY As Single
        GetSliderCoordinates relevantSliderPosX, relevantSliderPosY
        
        'Additional draw variables are required for the "default" draw style, which fills the slider track to match the current
        ' knob position.
        Dim customX As Single, customY As Single
        Dim relevantMin As Single
        
        If m_SliderStyle = DefaultStyle Then
        
            'Determine a minimum value for the control, using the formula provided:
            ' 1) If 0 is a valid control value, use 0.
            ' 2) If 0 is not a valid control value, use the control minimum.
            If (0 >= m_Min) And (0 <= m_Max) Then relevantMin = 0 Else relevantMin = m_Min
            
            'Convert that value into an actual pixel position on the track, then render a line between it and the current thumb position
            GetCustomValueCoordinates relevantMin, customX, customY
            GDI_Plus.GDIPlusDrawLineToDC bufferDC, customX, customY, relevantSliderPosX, customY, trackHighlightColor, 255, m_TrackDiameter + 1, True, LineCapRound
            
        End If
        
        'If the mouse is *not* over the slider, draw a small dot on the background track to indicate where the slider will "jump"
        ' if the mouse is clicked.
        If m_MouseOverSliderTrack Then
            Dim jumpIndicatorDiameter As Single: jumpIndicatorDiameter = m_TrackDiameter
            GDI_Plus.GDIPlusFillEllipseToDC bufferDC, m_MouseTrackX - (jumpIndicatorDiameter / 2), (m_SliderAreaHeight \ 2) - (jumpIndicatorDiameter / 2), jumpIndicatorDiameter, jumpIndicatorDiameter, trackJumpIndicatorColor, True
        End If
        
        'Finally, draw the thumb
        GDI_Plus.GDIPlusFillEllipseToDC bufferDC, relevantSliderPosX - (m_SliderDiameter \ 2), relevantSliderPosY - (m_SliderDiameter \ 2), m_SliderDiameter, m_SliderDiameter, thumbFillColor, True
        
        'Draw the edge (exterior) circle around the slider.
        Dim sliderWidth As Single
        If m_MouseOverSlider Then sliderWidth = 2 Else sliderWidth = 1.5
        GDI_Plus.GDIPlusDrawCircleToDC bufferDC, relevantSliderPosX, relevantSliderPosY, m_SliderDiameter \ 2, thumbBorderColor, 255, sliderWidth, True
        
    End If
    
    ucSupport.RequestRepaint refreshImmediately
    
End Sub

'Before this control does any painting, we need to retrieve relevant colors from PD's primary theming class.  Note that this
' step must also be called if/when PD's visual theme settings change.
Private Sub UpdateColorList()
    With m_Colors
        .LoadThemeColor PDSS_Background, "Background", IDE_WHITE
        .LoadThemeColor PDSS_TrackBack, "TrackBack", IDE_GRAY
        .LoadThemeColor PDSS_TrackFill, "TrackFill", IDE_BLUE
        .LoadThemeColor PDSS_TrackJumpIndicator, "TrackJumpIndicator", IDE_BLUE
        .LoadThemeColor PDSS_ThumbFill, "ThumbFill", IDE_WHITE
        .LoadThemeColor PDSS_ThumbBorder, "ThumbBorder", IDE_BLUE
        .LoadThemeColor PDSS_Notch, "Notch", IDE_GRAY
    End With
End Sub

'External functions can call this to request a redraw.  This is helpful for live-updating theme settings, as in the Preferences dialog.
Public Sub UpdateAgainstCurrentTheme()
    UpdateColorList
    If g_IsProgramRunning Then ucSupport.UpdateAgainstThemeAndLanguage
    UpdateControlLayout
End Sub

'By design, PD prefers to not use design-time tooltips.  Apply tooltips at run-time, using this function.
' (IMPORTANT NOTE: translations are handled automatically.  Always pass the original English text!)
Public Sub AssignTooltip(ByVal newTooltip As String, Optional ByVal newTooltipTitle As String, Optional ByVal newTooltipIcon As TT_ICON_TYPE = TTI_NONE)
    ucSupport.AssignTooltip UserControl.ContainerHwnd, newTooltip, newTooltipTitle, newTooltipIcon
End Sub
