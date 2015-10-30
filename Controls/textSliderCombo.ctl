VERSION 5.00
Begin VB.UserControl sliderTextCombo 
   BackColor       =   &H80000005&
   ClientHeight    =   495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6000
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
   ScaleHeight     =   33
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   400
   ToolboxBitmap   =   "textSliderCombo.ctx":0000
   Begin PhotoDemon.textUpDown tudPrimary 
      Height          =   345
      Left            =   4800
      TabIndex        =   1
      Top             =   45
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   609
   End
   Begin VB.PictureBox picScroll 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   60
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   313
      TabIndex        =   0
      Top             =   60
      Width           =   4695
   End
End
Attribute VB_Name = "sliderTextCombo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Text / Slider custom control
'Copyright 2013-2015 by Tanner Helland
'Created: 19/April/13
'Last updated: 25/August/15
'Last update: integrate captions directly into the control itself
'
'Software like PhotoDemon requires a lot of UI elements.  Ideally, every setting should be adjustable by at least
' two mechanisms: direct text entry, and some kind of slider or scroll bar, which allows for a quick method to
' make both large and small adjustments to a given parameter.
'
'Historically, I accomplished this by providing a scroll bar and text box for every parameter in the program.
' This got the job done, but it had a number of limitations - such as requiring an enormous amount of time if
' changes ever needed to be made, and custom code being required in every form to handle text / scroll syncing.
'
'In April 2013, it was brought to my attention that some locales (e.g. Italy) use a comma instead of a decimal
' for float values.  Rather than go through and add custom support for this to every damn form, I finally did
' the smart thing and built a custom text/scroll user control.  This effectively replaces all other text/scroll
' combos in the program.
'
'In June 2014, I finally did what I should have done long ago and swapped out the scroll bar for a custom-drawn
' slider.  That update also added support for some new features (like custom images on the background-track),
' while helping prepare PD for full theming support.
'
'Anyway, as of today, this control handles the following things automatically:
' 1) Syncing of text and scroll/slide values
' 2) Validation of text entries, including a function for external validation requests
' 3) Locale handling (like the aforementioned comma/decimal replacement in some countries)
' 4) A single "Change" event that fires for either scroll or text changes, and only if a text change is valid
' 5) Support for integer or floating-point values via the "SigDigits" property
' 6) Several different drawing modes, including support for 2- or 3-point gradients
' 7) Self-captioning, to remove the need for a redundant label control next to this one
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'This object provides a single raised event:
' - Change (which triggers when either the scrollbar or text box is modified in any way)
Public Event Change()

'Because we have multiple components on this user control, including an API text box, we report our own Got/Lost focus events.
Public Event GotFocusAPI()
Public Event LostFocusAPI()

'Reliable focus detection on the slider picture box requires a specialized subclasser
Private WithEvents cFocusDetector As pdFocusDetector
Attribute cFocusDetector.VB_VarHelpID = -1

'Flicker-free window painters for both the slider, and the background region
Private WithEvents cSliderPainter As pdWindowPainter
Attribute cSliderPainter.VB_VarHelpID = -1
Private WithEvents cBackgroundPainter As pdWindowPainter
Attribute cBackgroundPainter.VB_VarHelpID = -1

'API technique for drawing a focus rectangle; used only for designer mode (see the Paint method for details)
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function DrawFocusRect Lib "user32" (ByVal hDC As Long, lpRect As RECT) As Long

'Additional helper for rendering themed and multiline tooltips
Private toolTipManager As pdToolTip

'Used to internally track value, min, and max values as floating-points
Private controlVal As Double, controlMin As Double, controlMax As Double

'The number of significant digits for this control.  0 means integer values.
Private significantDigits As Long

'pdCaption manages all caption-related settings, so we don't have to.  (Note that this includes not just the caption, but related
' settings like caption font size.)
Private m_Caption As pdCaption
Attribute m_Caption.VB_VarHelpID = -1

'Two font sizes are currently supported: one for the control caption, and one for the text entry area.  pdCaption manages the
' caption one for us, so only the text up/down (TUD) font size is relevant here.
Private m_FontSizeTUD As Single

'If the text box is initiating a value change, we must track that so as to not overwrite the user's entry mid-typing
Private textBoxInitiated As Boolean

'Mouse and keyboard input handlers
Private WithEvents cMouseEvents As pdInputMouse
Attribute cMouseEvents.VB_VarHelpID = -1
Private WithEvents cKeyEvents As pdInputKeyboard
Attribute cKeyEvents.VB_VarHelpID = -1

'When the mouse is down on the slider, these values will be updated accordingly
Private m_MouseDown As Boolean
Private m_InitX As Single, m_InitY As Single

'Track and slider diameter, at 96 DPI.  Note that the actual render and hit detection functions will adjust these constants for
' the current screen DPI.
Private Const TRACK_DIAMETER As Long = 6
Private Const SLIDER_DIAMETER As Long = 16

'Track and slider diameter, at current DPI.  This is set when the control is first loaded.  In a perfect world, we would catch screen
' DPI changes and update these values accordingly, but I'm postponing that project until a later date.
Private m_trackDiameter As Single, m_sliderDiameter As Single

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

Private curSliderStyle As SLIDER_TRACK_STYLE

'Gradient colors.  For the two-color gradient style, only colors Left and Right are relevant.  Color Middle is used for the
' 3-color style only, and note that it *must* be accompanied by an owner-supplied middle position value.
Private gradColorLeft As Long, gradColorRight As Long, gradColorMiddle As Long
Private gradMiddleValue As Double

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
Private curNotchPosition As SLIDER_NOTCH_POSITION
Private customNotchValue As Double

'When the slider track is drawn, this rect will be filled with its relevant coordinates.  We use this to track Mouse_Over behavior,
' so we can change the cursor to a hand.
Private m_SliderTrackRect As RECTF

'Internal gradient DIB.  This is recreated as necessary to reflect the gradient colors and positions.
Private m_GradientDIB As pdDIB

'Full slider background DIB, with gradient, outline, notch (if any).  The only thing missing is the slider knob, which is added
' to the final buffer in a separate step (as it is the most likely to require changes!)
Private m_SliderBackgroundDIB As pdDIB

'This control manages two buffers: one for the control itself (with text painted atop it), and one for the slider.  This improves
' performance during redraws, as the smaller slider region is most likely to require repaints, so we can simply repaint it when
' necessary instead of redrawing the entire control.
Private m_BackBufferControl As pdDIB
Private m_BackBufferSlider As pdDIB

'Tracks whether the control (any component) has focus.  This is helpful as we must synchronize between VB's focus events and API
' focus events.  Every time an individual component gains focus, we increment this counter by 1.  Every time an individual component
' loses focus, we decrement the counter by 1.  When the counter hits 0, we report a control-wide Got/LostFocusAPI event.
Private m_ControlFocusCount As Long

'Used to prevent recursive redraws
Private m_InternalResizeActive As Boolean

'If the control is currently visible, this will be set to TRUE.  This can be used to suppress redraw requests for hidden controls.
Private m_ControlIsVisible As Boolean

'For UI purposes, we track whether the mouse is over the slider, or the background track.  Note that these are
' mutually exclusive; if the mouse is over the slider, it will *not* be marked as over the background track.
Private m_MouseOverSlider As Boolean, m_MouseOverSliderTrack As Boolean, m_MouseTrackX As Single

'Caption is handled just like the common control label's caption property.  It is valid at design-time, and any translation,
' if present, will not be processed until run-time.
' IMPORTANT NOTE: only the ENGLISH caption is returned.  I don't have a reason for returning a translated caption (if any),
'                  but I can revisit in the future if it ever becomes relevant.
Public Property Get Caption() As String
Attribute Caption.VB_UserMemId = -518
    Caption = m_Caption.GetCaptionEn
End Property

Public Property Let Caption(ByRef newCaption As String)
    If m_Caption.SetCaption(newCaption) And (m_ControlIsVisible Or (Not g_IsProgramRunning)) Then updateControlLayout
    PropertyChanged "Caption"
End Property

'The Enabled property is a bit unique; see http://msdn.microsoft.com/en-us/library/aa261357%28v=vs.60%29.aspx
Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal newValue As Boolean)
    
    UserControl.Enabled = newValue
    
    'Disable text entry
    tudPrimary.Enabled = newValue
    
    'Redraw the slider; when disabled, the slider itself is not drawn (only the track behind it is)
    redrawSlider
    
    PropertyChanged "Enabled"
    
End Property

Public Property Get FontSizeCaption() As Single
    FontSizeCaption = m_Caption.GetFontSize
End Property

Public Property Let FontSizeCaption(ByVal newSize As Single)
    If m_Caption.SetFontSize(newSize) And (m_ControlIsVisible Or (Not g_IsProgramRunning)) Then updateControlLayout
    PropertyChanged "FontSizeCaption"
End Property

Public Property Get FontSizeTUD() As Single
    FontSizeTUD = m_FontSizeTUD
End Property

Public Property Let FontSizeTUD(ByVal newSize As Single)
    If m_FontSizeTUD <> newSize Then
        m_FontSizeTUD = newSize
        tudPrimary.FontSize = m_FontSizeTUD
        PropertyChanged "FontSizeTUD"
    End If
End Property

'Gradient colors.  For the two-color gradient style, only colors Left and Right are relevant.  Color Middle is used for the
' 3-color style only, and note that it *must* be accompanied by an owner-supplied middle position value.
Public Property Get GradientColorLeft() As OLE_COLOR
    GradientColorLeft = gradColorLeft
End Property

Public Property Get GradientColorMiddle() As OLE_COLOR
    GradientColorMiddle = gradColorMiddle
End Property

Public Property Get GradientColorRight() As OLE_COLOR
    GradientColorRight = gradColorRight
End Property

Public Property Let GradientColorLeft(ByVal newColor As OLE_COLOR)

    'Store the new color, then redraw the slider to match
    If newColor <> gradColorLeft Then
        gradColorLeft = ConvertSystemColor(newColor)
        redrawInternalGradientDIB
        redrawSlider
        PropertyChanged "GradientColorLeft"
    End If

End Property

Public Property Let GradientColorMiddle(ByVal newColor As OLE_COLOR)

    'Store the new color, then redraw the slider to match
    If newColor <> gradColorMiddle Then
        gradColorMiddle = ConvertSystemColor(newColor)
        redrawInternalGradientDIB
        redrawSlider
        PropertyChanged "GradientColorMiddle"
    End If

End Property

Public Property Let GradientColorRight(ByVal newColor As OLE_COLOR)

    'Store the new color, then redraw the slider to match
    If newColor <> gradColorRight Then
        gradColorRight = ConvertSystemColor(newColor)
        redrawInternalGradientDIB
        redrawSlider
        PropertyChanged "GradientColorRight"
    End If

End Property

'Custom middle value for the 3-color gradient style.  This value is ignored for all other styles.
Public Property Get GradientMiddleValue() As Double
    GradientMiddleValue = gradMiddleValue
End Property

Public Property Let GradientMiddleValue(ByVal newValue As Double)
    
    'Store the new value, then redraw the slider to match
    If newValue <> gradMiddleValue Then
        gradMiddleValue = newValue
        redrawSlider
        PropertyChanged "GradientMiddleValue"
        redrawInternalGradientDIB
    End If
    
End Property

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

'If the current text value is NOT valid, this will return FALSE.  Note that this property is read-only.
Public Property Get IsValid(Optional ByVal showError As Boolean = True) As Boolean
    IsValid = tudPrimary.IsValid
End Property

'Note: the control's maximum value is settable at run-time
Public Property Get Max() As Double
    Max = controlMax
End Property

Public Property Let Max(ByVal newValue As Double)
    
    controlMax = newValue
    tudPrimary.Max = controlMax
    
    'If the track style is some kind of custom gradient, recreate our internal gradient DIB now
    If (curSliderStyle = GradientTwoPoint) Or (curSliderStyle = GradientThreePoint) Then redrawInternalGradientDIB
    
    'If the current control value is greater than the new max, update it to match (and raise a corresponding _Change event)
    If controlVal > controlMax Then Value = controlMax
    
    'Redraw the control
    redrawSlider
    
    PropertyChanged "Max"
    
End Property

'Note: the control's minimum value is settable at run-time
Public Property Get Min() As Double
    Min = controlMin
End Property

Public Property Let Min(ByVal newValue As Double)
    
    controlMin = newValue
    tudPrimary.Min = controlMin
    
    'If the track style is some kind of custom gradient, recreate our internal gradient DIB now
    If (curSliderStyle = GradientTwoPoint) Or (curSliderStyle = GradientThreePoint) Then redrawInternalGradientDIB
    
    'If the current control value is less than the new minimum, update it to match (and raise a corresponding _Change event)
    If controlVal < controlMin Then Value = controlMin
    
    'Redraw the control
    redrawSlider
    
    PropertyChanged "Min"
    
End Property

'Notch positioning technique.  If CUSTOM is set, make sure to supply a custom value to match!
Public Property Get NotchPosition() As SLIDER_NOTCH_POSITION
    NotchPosition = curNotchPosition
End Property

Public Property Let NotchPosition(ByVal newPosition As SLIDER_NOTCH_POSITION)
    
    'Store the new position
    curNotchPosition = newPosition
    
    'Redraw the control
    redrawSlider
    
    'Raise the property changed event
    PropertyChanged "NotchPosition"
    
End Property

'Custom notch value.  This value is only used if NotchPosition = CustomPosition.
Public Property Get NotchValueCustom() As Double
    NotchValueCustom = customNotchValue
End Property

Public Property Let NotchValueCustom(ByVal newValue As Double)
    
    'Store the new position
    customNotchValue = newValue
    
    'Redraw the control
    redrawSlider
    
    'Raise the property changed event
    PropertyChanged "NotchValueCustom"
    
End Property

'Significant digits determines whether the control allows float values or int values (and with how much precision)
Public Property Get SigDigits() As Long
    SigDigits = significantDigits
End Property

Public Property Let SigDigits(ByVal newValue As Long)
    significantDigits = newValue
    tudPrimary.SigDigits = significantDigits
    PropertyChanged "SigDigits"
End Property

Public Property Get SliderTrackStyle() As SLIDER_TRACK_STYLE
    SliderTrackStyle = curSliderStyle
End Property

Public Property Let SliderTrackStyle(ByVal newStyle As SLIDER_TRACK_STYLE)
    
    'Store the new style
    curSliderStyle = newStyle
    
    'Redraw the control
    redrawSlider
    
    'Raise the property changed event
    PropertyChanged "SliderTrackStyle"
    
End Property

'The control's value is simply a reflection of the embedded scroll bar and text box
Public Property Get Value() As Double
Attribute Value.VB_UserMemId = 0
    Value = controlVal
End Property

Public Property Let Value(ByVal newValue As Double)
    
    'Don't make any changes unless the new value deviates from the existing one
    If newValue <> controlVal Then
    
        'Internally track the value of the control
        controlVal = newValue
        
        'Check bounds
        If controlVal < controlMin Then controlVal = controlMin
        If controlVal > controlMax Then controlVal = controlMax
        
        'Mirror the value to the text box
        If Not textBoxInitiated Then
            
            'Normally, we want to make sure that the control's value has changed; otherwise, updating the text box causes unnecessary
            ' recursive refreshing.  However, we can't compare the text box value to the control value if the user has entered invalid
            ' input, so first make sure that the text box contains meaningful data.
            If tudPrimary.IsValid(False) Then
                
                'The text box contains valid numerical data.  If it matches the current control value, skip the refresh step.
                If StrComp(getFormattedStringValue(tudPrimary), CStr(controlVal), vbBinaryCompare) <> 0 Then
                    tudPrimary.Value = CStr(controlVal)
                End If
            
            'The text box is currently in an error state.  Copy the new text into place without a duplication check.
            Else
            
                tudPrimary.Value = CStr(controlVal)
            
            End If
            
        End If
                
        'Redraw the slider to reflect the new value
        drawSliderKnob
        
        'Mark the value property as being changed, and raise the corresponding event.
        PropertyChanged "Value"
        RaiseEvent Change
        
    End If
    
End Property

Private Sub cFocusDetector_GotFocusReliable()
    m_ControlFocusCount = m_ControlFocusCount + 1
    evaluateFocusCount True
End Sub

Private Sub cFocusDetector_LostFocusReliable()
    m_ControlFocusCount = m_ControlFocusCount - 1
    evaluateFocusCount False
End Sub

'Arrow keys can be used to "nudge" the control value in single-unit increments.
Private Sub cKeyEvents_KeyDownCustom(ByVal Shift As ShiftConstants, ByVal vkCode As Long, markEventHandled As Boolean)

    'Up and right arrows are used to increment the slider value
    If (vkCode = VK_UP) Or (vkCode = VK_RIGHT) Then
        Value = Value + getIncrementAmount
    End If
    
    'Left and down arrows decrement it
    If (vkCode = VK_LEFT) Or (vkCode = VK_DOWN) Then
        Value = Value - getIncrementAmount
    End If

End Sub

Private Sub cMouseEvents_MouseDownCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    
    If ((Button And pdLeftButton) <> 0) Then
    
        'Check to see if the mouse is over a) the slider control button, or b) the background track
        If isMouseOverSlider(x, y) Then
        
            'Track various states to make UI rendering easier
            m_MouseDown = True
            m_MouseOverSlider = True
            m_MouseOverSliderTrack = False
            
            'Calculate a new control value.  This will cause the slider to "jump" to the current position.
            Value = (controlMax - controlMin) * ((x - getTrackMinPos) / (getTrackMaxPos - getTrackMinPos)) + controlMin
            
            'Retrieve the current slider x/y values, and store the mouse position relative to those values
            Dim sliderX As Single, sliderY As Single
            getSliderCoordinates sliderX, sliderY
            m_InitX = x - sliderX
            m_InitY = y - sliderY
            
            'Force an immediate redraw (instead of waiting for WM_PAINT to process)
            cSliderPainter.RequestRepaint True
            
        End If
    
    End If
    
End Sub

Private Sub cMouseEvents_MouseLeave(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    
    'Reset all hover indicators
    m_MouseOverSlider = False
    m_MouseOverSliderTrack = False
    redrawSlider
    
    'Reset the mouse pointer as well
    cMouseEvents.setSystemCursor IDC_ARROW
    
End Sub

Private Sub cMouseEvents_MouseMoveCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)

    'If the mouse is down, adjust the current control value accordingly.
    If m_MouseDown Then
        
        'Because the slider will be tracking the mouse's motion, we automatically assume the mouse is over it
        ' (and not the background track).
        m_MouseOverSlider = True
        m_MouseOverSliderTrack = False
        
        'Calculate a new control value relative to the current mouse position.  (This will automatically force a button redraw.)
        Value = (controlMax - controlMin) * (((x + m_InitX) - getTrackMinPos) / (getTrackMaxPos - getTrackMinPos)) + controlMin
        
        'Force an immediate redraw (instead of waiting for WM_PAINT to process)
        cSliderPainter.RequestRepaint True
        
    'If the LMB is not down, modify the cursor according to its position relative to the slider
    Else
        
        m_MouseOverSlider = isMouseOverSlider(x, y, False)
        If m_MouseOverSlider Then
            m_MouseOverSliderTrack = False
        Else
            m_MouseOverSliderTrack = isMouseOverSlider(x, y, True)
            m_MouseTrackX = x
        End If
        
        If m_MouseOverSlider Or m_MouseOverSliderTrack Then
            cMouseEvents.setSystemCursor IDC_HAND
        Else
            cMouseEvents.setSystemCursor IDC_ARROW
        End If
        
        'Redraw the button to match the new hover state, if any
        redrawSlider
    
    End If

End Sub

Private Sub cMouseEvents_MouseUpCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal ClickEventAlsoFiring As Boolean)
    
    If ((Button And pdLeftButton) <> 0) And m_MouseDown Then
        
        'Perform a final mouse move update at the reported x/y position.  If intensive processing occurred while the slider was being
        ' interacted with, this will ensure that the mouse location at its exact point of release is used.
        Value = (controlMax - controlMin) * (((x + m_InitX) - getTrackMinPos) / (getTrackMaxPos - getTrackMinPos)) + controlMin
        
        m_MouseDown = False
        
    End If
    
End Sub

Private Function isMouseOverSlider(ByVal mouseX As Single, ByVal mouseY As Single, Optional ByVal alsoCheckBackgroundTrack As Boolean = True) As Boolean

    'Retrieve the current x/y position of the slider's CENTER
    Dim sliderX As Single, sliderY As Single
    getSliderCoordinates sliderX, sliderY
    
    'See if the mouse is within distance of the slider's center
    If distanceTwoPoints(sliderX, sliderY, mouseX, mouseY) < FixDPI(SLIDER_DIAMETER) \ 2 Then
        isMouseOverSlider = True
    Else
        
        'If the mouse is not over the slider itself, check the background track as well
        If isPointInRectF(mouseX, mouseY, m_SliderTrackRect) And alsoCheckBackgroundTrack Then
            isMouseOverSlider = True
        Else
            isMouseOverSlider = False
        End If
    End If

End Function

'The pdWindowPaint class raises this event when the control needs to be redrawn.  The passed coordinates contain the
' rect returned by GetUpdateRect (but with right/bottom measurements pre-converted to width/height).
Private Sub cSliderPainter_PaintWindow(ByVal winLeft As Long, ByVal winTop As Long, ByVal winWidth As Long, ByVal winHeight As Long)
    BitBlt picScroll.hDC, winLeft, winTop, winWidth, winHeight, m_BackBufferSlider.getDIBDC, winLeft, winTop, vbSrcCopy
End Sub

Private Sub cBackgroundPainter_PaintWindow(ByVal winLeft As Long, ByVal winTop As Long, ByVal winWidth As Long, ByVal winHeight As Long)
    
    If Not m_InternalResizeActive Then
        BitBlt UserControl.hDC, winLeft, winTop, winWidth, winHeight, m_BackBufferControl.getDIBDC, winLeft, winTop, vbSrcCopy
    End If
    
End Sub

Private Sub tudPrimary_Change()
    If tudPrimary.IsValid(False) Then
        textBoxInitiated = True
        Me.Value = tudPrimary.Value
        textBoxInitiated = False
    End If
End Sub

Private Sub tudPrimary_GotFocusAPI()
    m_ControlFocusCount = m_ControlFocusCount + 1
    evaluateFocusCount True
End Sub

Private Sub tudPrimary_LostFocusAPI()
    m_ControlFocusCount = m_ControlFocusCount - 1
    evaluateFocusCount False
End Sub

Private Sub UserControl_GotFocus()
    m_ControlFocusCount = m_ControlFocusCount + 1
    evaluateFocusCount True
End Sub

Private Sub UserControl_Hide()
    m_ControlIsVisible = False
End Sub

Private Sub UserControl_Initialize()
    
    'When not in design mode, initialize a tracker for mouse and keyboard events
    If g_IsProgramRunning Then
        
        'Start our flicker-free window painters
        Set cSliderPainter = New pdWindowPainter
        cSliderPainter.StartPainter picScroll.hWnd
        
        Set cBackgroundPainter = New pdWindowPainter
        cBackgroundPainter.StartPainter UserControl.hWnd
        
        'Set up mouse events
        Set cMouseEvents = New pdInputMouse
        cMouseEvents.addInputTracker picScroll.hWnd, True, True, , True
        cMouseEvents.setSystemCursor IDC_HAND
        
        'Set up keyboard events
        Set cKeyEvents = New pdInputKeyboard
        cKeyEvents.createKeyboardTracker "Slider/Text UC", picScroll.hWnd, VK_LEFT, VK_UP, VK_RIGHT, VK_DOWN
        
        'Also start a focus detector for the slider picture box
        Set cFocusDetector = New pdFocusDetector
        cFocusDetector.startFocusTracking picScroll.hWnd
        
        'Create a tooltip engine
        Set toolTipManager = New pdToolTip
        
    End If
    
    'Update the control-level track and slider diameters to reflect current screen DPI
    m_trackDiameter = FixDPI(TRACK_DIAMETER)
    m_sliderDiameter = FixDPI(SLIDER_DIAMETER)
    
    'Set slider area width/height
    m_SliderAreaWidth = picScroll.ScaleWidth
    m_SliderAreaHeight = picScroll.ScaleHeight
    
    'Initialize various back buffers and background DIBs
    Set m_SliderBackgroundDIB = New pdDIB
    Set m_BackBufferSlider = New pdDIB
    Set m_BackBufferControl = New pdDIB
    
    'Prep the caption object
    Set m_Caption = New pdCaption
    m_Caption.SetWordWrapSupport False
        
End Sub

'Initialize control properties for the first time
Private Sub UserControl_InitProperties()

    'Reset all controls to their default state.  For each public property, matching internal tracker variables are also updated;
    ' this is not necessary, but it's helpful for reminding me of the names of the internal tracker variables relevant to their
    ' connected property.
    FontSizeTUD = 10
    FontSizeCaption = 12
    Caption = ""
    
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
    
    'Default notch position; for most controls, it should be set to AUTOMATIC.  If CUSTOM is set, make sure to supply whatever
    ' custom value you want in the corresponding property!
    NotchPosition = AutomaticPosition
    NotchValueCustom = 0
    
End Sub

Private Sub UserControl_LostFocus()
    m_ControlFocusCount = m_ControlFocusCount - 1
    evaluateFocusCount False
End Sub

Private Sub UserControl_Paint()
    
    'Provide some visual feedback in the IDE
    If Not g_IsProgramRunning Then
        BitBlt UserControl.hDC, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, m_BackBufferControl.getDIBDC, 0, 0, vbSrcCopy
        redrawSlider
    End If
    
End Sub

'Read control properties from file
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    With PropBag
        Caption = .ReadProperty("Caption", "")
        FontSizeCaption = .ReadProperty("FontSizeCaption", 12)
        FontSizeTUD = .ReadProperty("FontSizeTUD", 10)
        SigDigits = .ReadProperty("SigDigits", 0)
        Max = .ReadProperty("Max", 10)
        Min = .ReadProperty("Min", 0)
        SliderTrackStyle = .ReadProperty("SliderTrackStyle", DefaultStyle)
        Value = .ReadProperty("Value", 0)
        GradientColorLeft = .ReadProperty("GradientColorLeft", RGB(0, 0, 0))
        GradientColorRight = .ReadProperty("GradientColorRight", RGB(255, 255, 255))
        GradientColorMiddle = .ReadProperty("GradientColorMiddle", RGB(121, 131, 135))
        GradientMiddleValue = .ReadProperty("GradientMiddleValue", 0)
        NotchPosition = .ReadProperty("NotchPosition", 0)
        NotchValueCustom = .ReadProperty("NotchValueCustom", 0)
    End With
    
End Sub

Private Sub UserControl_Resize()
    If Not m_InternalResizeActive Then updateControlLayout
End Sub

Private Sub UserControl_Show()
        
    m_ControlIsVisible = True
        
    'When the control is first made visible, remove the control's tooltip property and reassign it to the checkbox
    ' using a custom solution (which allows for linebreaks and theming).
    If Len(Extender.ToolTipText) <> 0 Then AssignTooltip Extender.ToolTipText
    
    'If the track style is some kind of custom gradient, recreate our internal gradient DIB now
    If (curSliderStyle = GradientTwoPoint) Or (curSliderStyle = GradientThreePoint) Or (curSliderStyle = HueSpectrum360) Then redrawInternalGradientDIB
    
    updateControlLayout
    redrawSlider
        
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    'Store all associated properties
    With PropBag
        .WriteProperty "Caption", m_Caption.GetCaptionEn, ""
        .WriteProperty "FontSizeCaption", m_Caption.GetFontSize, 12
        .WriteProperty "FontSizeTUD", m_FontSizeTUD, 10
        .WriteProperty "Min", controlMin, 0
        .WriteProperty "Max", controlMax, 10
        .WriteProperty "SigDigits", significantDigits, 0
        .WriteProperty "SliderTrackStyle", curSliderStyle, DefaultStyle
        .WriteProperty "Value", controlVal, 0
        .WriteProperty "GradientColorLeft", gradColorLeft, RGB(0, 0, 0)
        .WriteProperty "GradientColorRight", gradColorRight, RGB(255, 255, 255)
        .WriteProperty "GradientColorMiddle", gradColorMiddle, RGB(121, 131, 135)
        .WriteProperty "GradientMiddleValue", gradMiddleValue, 0
        .WriteProperty "NotchPosition", curNotchPosition, 0
        .WriteProperty "NotchValueCustom", customNotchValue, 0
    End With
    
End Sub

'When the control is resized, the caption is changed, or font sizes for either the caption or text up/down are modified,
' this function should be called.  It controls the physical positioning of various control sub-elements
' (specifically, the caption area, the slider area, and the text up/down area).
Private Sub updateControlLayout()
    
    If m_InternalResizeActive Then Exit Sub
    
    'Set a control-level flag to prevent recursive redraws
    m_InternalResizeActive = True
    
    'To avoid long-running redraw operations, we only want to apply new positions and sizes as necessary (e.g. if they don't match
    ' existing values)
    Dim newLeft_TUD As Long, newLeft_Slider As Long
    Dim newTop_TUD As Long, newTop_Slider As Long
    Dim newWidth_Slider As Long, newHeight As Long
    Dim newControlHeight As Long
    
    'NB: order of operations is important in this function.  We first calculate all new size/position values.  When all new values
    '    are known, we apply them in a single fell swoop to avoid the need for costly intermediary redraws.
    
    'The first (and most complicated) size consideration is accounting for the presence of a control caption.  If no caption exists,
    ' we can bypass much of this function.
    If m_Caption.IsCaptionActive Then
        
        'Notify the caption renderer of our width.  It will auto-fit its font to match.
        ' (Because this control doesn't support wordwrap, container height is irrelevant; pass 0)
        m_Caption.SetControlSize UserControl.ScaleWidth, 0
        
        'We now have all the information necessary to calculate caption positioning (and by extension, slider and
        ' text up/down positioning, too!)
        
        'Calculate a new height for the usercontrol as a whole.  This is simple formula:
        ' (height of text up/down) + (2 px padding around text up/down) + (height of caption) + (1 px padding around caption)
        Dim textHeight As Long
        textHeight = m_Caption.GetCaptionHeight()
        newControlHeight = tudPrimary.Height + FixDPI(4) + textHeight + FixDPI(2)
        
        'Calculate a new top position for the slider box (which will be vertically centered in the space below the caption)
        newTop_Slider = ((newControlHeight - (textHeight + FixDPI(4))) - tudPrimary.Height) \ 2
        newTop_Slider = textHeight + FixDPI(4) + newTop_Slider
        
    'When a slider lacks a caption, we hard-code its height to preset values
    Else
        
        'Start by setting the control height
        newControlHeight = tudPrimary.Height + FixDPI(4)
        
        'Center the slider box inside the newly calculated height
        newTop_Slider = (newControlHeight - picScroll.Height) \ 2
                
    End If
    
    'Apply the new height
    If UserControl.Extender.Height <> newControlHeight Then UserControl.Extender.Height = newControlHeight
    
    'With the control correctly sized, prep the back buffer to match
    Dim controlBackgroundColor As Long
    If g_IsProgramRunning Then
        controlBackgroundColor = g_Themer.GetThemeColor(PDTC_BACKGROUND_DEFAULT)
    Else
        controlBackgroundColor = vbWhite
    End If
    
    If (m_BackBufferControl.getDIBWidth <> UserControl.ScaleWidth) Or (m_BackBufferControl.getDIBHeight <> UserControl.ScaleHeight) Or (Not g_IsProgramRunning) Then
        m_BackBufferControl.createBlank UserControl.ScaleWidth, UserControl.ScaleHeight, 24, controlBackgroundColor
    Else
        GDI_Plus.GDIPlusFillDIBRect m_BackBufferControl, 0, 0, m_BackBufferControl.getDIBWidth, m_BackBufferControl.getDIBHeight, controlBackgroundColor
    End If
    
    'If text exists, paint it onto the newly created back buffer
    If m_Caption.IsCaptionActive Then m_Caption.DrawCaption m_BackBufferControl.getDIBDC, 1, 1
    
    'With height correctly set, we next want to left-align the TUD against the slider region
    newLeft_TUD = UserControl.ScaleWidth - (tudPrimary.Width + FixDPI(2))
    
    'If the slider width changes, we need to redraw the entire custom control, so we track this value specifically.
    Dim widthChanged As Boolean
    newWidth_Slider = newLeft_TUD - FixDPI(10)
    If (newWidth_Slider > 0) And (picScroll.Width <> newWidth_Slider) Then widthChanged = True Else widthChanged = False
    
    'We now know enough to reposition the slider picture box
    If ((picScroll.Top <> newTop_Slider) Or (picScroll.Width <> newWidth_Slider)) And (newWidth_Slider > 0) Then picScroll.Move picScroll.Left, newTop_Slider, newWidth_Slider
    
    'Vertically center the text up/down relative to the slider
    Dim sliderVerticalCenter As Single
    sliderVerticalCenter = picScroll.Top + (CSng(picScroll.ScaleHeight) / 2)
    newTop_TUD = sliderVerticalCenter - Int(CSng(tudPrimary.Height) / 2)
    
    'Now that we've calculated new text up/down positioning, we can apply it as necessary
    If tudPrimary.Top <> newTop_TUD Or tudPrimary.Left <> newLeft_TUD Then tudPrimary.Move newLeft_TUD, newTop_TUD
    
    'Update slider area width/height to match the new picScroll size
    m_SliderAreaWidth = picScroll.ScaleWidth
    m_SliderAreaHeight = picScroll.ScaleHeight
    
    'If the slider area changed, redraw it now
    If widthChanged Then
    
        'If the track style is some kind of custom gradient, start by redrawing the background gradient DIB
        If ((curSliderStyle = GradientTwoPoint) Or (curSliderStyle = GradientThreePoint) Or (curSliderStyle = HueSpectrum360)) Then redrawInternalGradientDIB
        
        'Redraw the slider as well
        redrawSlider
        
    End If
    
    'Paint the background text buffer to the screen, as relevant
    If g_IsProgramRunning Then
        cBackgroundPainter.RequestRepaint
    Else
        BitBlt UserControl.hDC, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, m_BackBufferControl.getDIBDC, 0, 0, vbSrcCopy
    End If
    
    'Reset the redraw flag, and request a background repaint
    m_InternalResizeActive = False
    
End Sub

'Render a custom slider to the slider area picture box.  Note that the background gradient, if any, should already have been created
' in a separate redrawInternalGradientDIB request.
Private Sub redrawSlider(Optional ByVal refreshImmediately As Boolean = False)

    'Drawing is done in several stages.  The bulk of the slider is rendered to a persistent slider-only DIB, which contains everything
    ' but the knob.  The knob is rendered in a separate step, as it is the most common update required, and we can shortcut by not
    ' redrawing the entire slider on every update.
    
    'Initialize the background DIB, as necessary
    If (m_SliderBackgroundDIB.getDIBWidth <> m_SliderAreaWidth) Or (m_SliderBackgroundDIB.getDIBHeight <> m_SliderAreaHeight) Then
        m_SliderBackgroundDIB.createBlank m_SliderAreaWidth, m_SliderAreaHeight, 24, RGB(255, 255, 255)
    End If
    
    If g_IsProgramRunning Then
        GDI_Plus.GDIPlusFillDIBRect m_SliderBackgroundDIB, 0, 0, m_SliderBackgroundDIB.getDIBWidth, m_SliderBackgroundDIB.getDIBHeight, g_Themer.GetThemeColor(PDTC_BACKGROUND_DEFAULT), 255
    End If
    
    'Initialize the back buffer as well
    If (m_BackBufferSlider.getDIBWidth <> m_SliderAreaWidth) Or (m_BackBufferSlider.getDIBHeight <> m_SliderAreaHeight) Then
        m_BackBufferSlider.createBlank m_SliderAreaWidth, m_SliderAreaHeight, 24, 0
    End If
        
    'There are a few components to the slider:
    ' 1) The track that sits behind the slider.  It has two relevant parameters: a radius, and a color.  Its width is automatically
    '     calculated relevant to the width of the control as a whole.
    ' 2) The slider knob that sits atop the track.  It has three relevant parameters: a radius, a fill color, and an edge color.
    '     Its width is constant from a programmatic standpoint, though it does get updated at run-time to account for screen DPI.
    
    'Pull relevant colors from the global themer object
    Dim trackColor As Long
    If g_IsProgramRunning Then
        trackColor = g_Themer.GetThemeColor(PDTC_GRAY_HIGHLIGHT)
    Else
        trackColor = RGB(127, 127, 127)
    End If
    
    'Retrieve the current slider x/y position.  Floating-point values are used so we can support sub-pixel positioning!
    Dim relevantSliderPosX As Single, relevantSliderPosY As Single
    getSliderCoordinates relevantSliderPosX, relevantSliderPosY
    
    'Draw the background track according to the current SliderTrackStyle property.
    If Me.Enabled Then
    
        'This control supports a variety of different track styles.  Some of these styles require a DIB supplied by the owner, and
        ' they *will not* render properly until that DIB is provided!
        Select Case curSliderStyle
        
            'Default style: fill the "active" part of track with the control highlight color.  The "active part" is the chunk relative
            ' to zero, if the control supports 0 as a value; otherwise, it is relative to the control minimum.
            Case DefaultStyle
            
                'Start by drawing the default background track
                GDI_Plus.GDIPlusDrawLineToDC m_SliderBackgroundDIB.getDIBDC, getTrackMinPos, m_SliderAreaHeight \ 2, getTrackMaxPos, m_SliderAreaHeight \ 2, trackColor, 255, m_trackDiameter + 1, True, LineCapRound
                
                'Filling the track to the notch position happens in the drawSliderKnob function.
                
            'No-frills slider: plain gray background (boooring - use only if absolutely necessary)
            Case NoFrills
                GDI_Plus.GDIPlusDrawLineToDC m_SliderBackgroundDIB.getDIBDC, getTrackMinPos, m_SliderAreaHeight \ 2, getTrackMaxPos, m_SliderAreaHeight \ 2, trackColor, 255, m_trackDiameter + 1, True, LineCapRound
            
            Case GradientTwoPoint, GradientThreePoint, HueSpectrum360
            
                'As a failsafe, make sure our internal gradient DIB exists
                If m_GradientDIB Is Nothing Then redrawInternalGradientDIB
                
                'Draw a stock trackline onto the target DIB.  This will serve as the border of the gradient track area.
                GDI_Plus.GDIPlusDrawLineToDC m_SliderBackgroundDIB.getDIBDC, getTrackMinPos, m_SliderAreaHeight \ 2, getTrackMaxPos, m_SliderAreaHeight \ 2, trackColor, 255, m_trackDiameter + 1, True, LineCapRound
                
                'Next, draw the gradient effect DIB to the location where we'd normally draw the track line.  Alpha has already been
                ' calculated for the gradient DIB, so it will sit precisely inside the trackline drawn above, giving the track a
                ' sharp 1px border.
                m_GradientDIB.alphaBlendToDC m_SliderBackgroundDIB.getDIBDC, 255, getTrackMinPos - (m_trackDiameter \ 2), 0
                
            Case CustomOwnerDrawn
        
        End Select
        
        'Before carrying on, draw a slight notch above and below the slider track, using the value specified by the associated property
        drawNotchToDIB m_SliderBackgroundDIB, trackColor
        
    'Control is disabled; draw a plain track in the background, but no notch or other frills
    Else
        GDI_Plus.GDIPlusDrawLineToDC m_SliderBackgroundDIB.getDIBDC, getTrackMinPos, m_SliderAreaHeight \ 2, getTrackMaxPos, m_SliderAreaHeight \ 2, trackColor, 255, m_trackDiameter + 1, True, LineCapRound
    End If
    
    'Store the calculated position of the slider background.  Mouse hit detection code can make use of this, so we don't have to
    ' constantly re-calculate it during mouse events.
    With m_SliderTrackRect
        .Left = getTrackMinPos
        .Width = getTrackMaxPos - .Left
        .Top = (m_SliderAreaHeight / 2) - ((m_trackDiameter + 1) / 2)
        .Height = m_trackDiameter + 1
    End With
        
    'The slider background is now ready for action.  As a final step, pass control to the knob renderer function.
    drawSliderKnob refreshImmediately
        
End Sub

'Composite the knob atop the final slider background, and keep the entire thing inside a persitent back buffer.
Private Sub drawSliderKnob(Optional ByVal refreshImmediately As Boolean = False)

    'Copy the background DIB into the back buffer
    BitBlt m_BackBufferSlider.getDIBDC, 0, 0, m_BackBufferSlider.getDIBWidth, m_BackBufferSlider.getDIBHeight, m_SliderBackgroundDIB.getDIBDC, 0, 0, vbSrcCopy
    
    'The slider itself is only drawn if the control is enabled; otherwise, we do not display it at all.
    If Me.Enabled Then
        
        'Retrieve colors from the global themer object
        Dim trackEffectColor As Long, trackJumpIndicatorColor As Long
        Dim sliderBackgroundColor As Long, sliderEdgeColor As Long
        
        If g_IsProgramRunning Then
        
            trackEffectColor = g_Themer.GetThemeColor(PDTC_ACCENT_HIGHLIGHT)
            trackJumpIndicatorColor = g_Themer.GetThemeColor(PDTC_ACCENT_SHADOW)
            
            If m_MouseOverSlider Then
                sliderEdgeColor = g_Themer.GetThemeColor(PDTC_ACCENT_DEFAULT)
                sliderBackgroundColor = g_Themer.GetThemeColor(PDTC_BACKGROUND_DEFAULT)
            Else
                sliderEdgeColor = g_Themer.GetThemeColor(PDTC_ACCENT_HIGHLIGHT)
                sliderBackgroundColor = g_Themer.GetThemeColor(PDTC_BACKGROUND_DEFAULT)
            End If
            
        Else
            sliderBackgroundColor = vbWhite
            sliderEdgeColor = vbBlue
            trackEffectColor = vbBlue
            trackJumpIndicatorColor = vbBlue
        End If
        
        'Retrieve the current slider x/y position.  Floating-point values are used so we can support sub-pixel positioning!
        Dim relevantSliderPosX As Single, relevantSliderPosY As Single
        getSliderCoordinates relevantSliderPosX, relevantSliderPosY
        
        'Additional draw variables are required for the "default" draw style, which fills the slider track to match the current
        ' knob position.
        Dim customX As Single, customY As Single
        Dim relevantMin As Single
        
        If curSliderStyle = DefaultStyle Then
        
            'Determine a minimum value for the control, using the formula provided:
            ' 1) If 0 is a valid control value, use 0.
            ' 2) If 0 is not a valid control value, use the control minimum.
            If (0 >= controlMin) And (0 <= controlMax) Then
                relevantMin = 0
            Else
                relevantMin = controlMin
            End If
            
            'Convert our newly calculated relevant min value into an actual pixel position on the track
            getCustomValueCoordinates relevantMin, customX, customY
            
            'Draw a highlighted line between the slider position and our calculated relevant minimum
            GDI_Plus.GDIPlusDrawLineToDC m_BackBufferSlider.getDIBDC, customX, customY, relevantSliderPosX, customY, trackEffectColor, 255, m_trackDiameter + 1, True, LineCapRound
            
        End If
        
        'If the mouse is *not* over the slider, draw a small dot on the background track to indicate where the slider will "jump"
        ' if the mouse is clicked.
        If m_MouseOverSliderTrack Then
            
            Dim jumpIndicatorDiameter As Single
            jumpIndicatorDiameter = m_trackDiameter
            
            GDI_Plus.GDIPlusFillEllipseToDC m_BackBufferSlider.getDIBDC, m_MouseTrackX - (jumpIndicatorDiameter / 2), (m_SliderAreaHeight \ 2) - (jumpIndicatorDiameter / 2), jumpIndicatorDiameter, jumpIndicatorDiameter, trackJumpIndicatorColor, True
            
        End If
        
        'Draw the background (interior fill) circle of the slider
        GDI_Plus.GDIPlusFillEllipseToDC m_BackBufferSlider.getDIBDC, relevantSliderPosX - (m_sliderDiameter \ 2), relevantSliderPosY - (m_sliderDiameter \ 2), m_sliderDiameter, m_sliderDiameter, sliderBackgroundColor, True
        
        'Draw the edge (exterior) circle around the slider
        If m_MouseOverSlider Then
            GDI_Plus.GDIPlusDrawCircleToDC m_BackBufferSlider.getDIBDC, relevantSliderPosX, relevantSliderPosY, m_sliderDiameter \ 2, sliderEdgeColor, 255, 2, True
        Else
            GDI_Plus.GDIPlusDrawCircleToDC m_BackBufferSlider.getDIBDC, relevantSliderPosX, relevantSliderPosY, m_sliderDiameter \ 2, sliderEdgeColor, 255, 1.5, True
        End If
        
    End If
    
    'Paint the buffer to the screen
    If g_IsProgramRunning Then cSliderPainter.RequestRepaint refreshImmediately Else BitBlt picScroll.hDC, 0, 0, picScroll.ScaleWidth, picScroll.ScaleHeight, m_BackBufferSlider.getDIBDC, 0, 0, vbSrcCopy
    
End Sub

'Post-translation, we can request an immediate refresh
Public Sub requestRefresh()
    cBackgroundPainter.RequestRepaint
    cSliderPainter.RequestRepaint
End Sub

'Render a slight notch at the specified position on the specified DIB.  Note that this sub WILL automatically convert a custom notch
' value into it's appropriate x-coordinate; the caller is not responsible for that.
Private Sub drawNotchToDIB(ByRef dstDIB As pdDIB, ByVal trackColor As Long)
    
    'First, see if a notch needs to be drawn.  If the notch mode is "none", exit now.
    If curNotchPosition = DoNotDisplayNotch Then Exit Sub
    
    Dim renderNotchValue As Double
    
    'For controls where the notch would be drawn at the "minimum value" position, I prefer to keep a clean visual style and
    ' not draw a redundant notch (as the filled slider conveys the exact same message).  For such controls, notch display
    ' is automatically overridden.
    Dim overrideNotchDraw As Boolean
    overrideNotchDraw = False
    
    'Next, calculate a notch position as necessary.  If the notch mode is "automatic", this function is responsible for
    ' determining what notch value to use.
    If curNotchPosition = AutomaticPosition Then
    
        'The automatic position varies according to a few factors.  First, some slider styles have their own notch calculations.
        If curSliderStyle = GradientThreePoint Then
        
            'Three-point gradients always display a notch at the position of the middle gradient point; this is the assumed default
            ' value of the control.
            renderNotchValue = GradientMiddleValue
        
        'All other slider styles use the same heuristic for automatic notch positioning.  If 0 is available, use it.
        ' Otherwise, use the control's minimum value.
        Else
            
            If (0 > controlMin) And (0 <= controlMax) Then
                renderNotchValue = 0
            Else
                renderNotchValue = controlMin
                
                'To keep sliders visually clean, notches are not drawn unless useful, and notches at the obvious minimum position
                ' serve no purpose - so override the entire notch drawing process.
                overrideNotchDraw = True
            End If
            
        End If
    
    'If the notch position is not "do not display", and also not "automatic", it must be custom.  Retrieve that value now.
    Else
        renderNotchValue = customNotchValue
    End If
    
    If Not overrideNotchDraw Then
    
        'Convert our calculated notch *value* into an actual *pixel position* on the track
        Dim customX As Single, customY As Single
        getCustomValueCoordinates renderNotchValue, customX, customY
        
        'Calculate the height of the notch; this varies by DPI, which is automatically factored into m_trackDiameter
        Dim notchSize As Single
        notchSize = (m_SliderAreaHeight - m_trackDiameter) \ 2 - 4
        
        'Draw a notch above and below the slider's track, then exit
        GDI_Plus.GDIPlusDrawLineToDC dstDIB.getDIBDC, customX, 1, customX, 1 + notchSize, trackColor, 255, 1, True, LineCapFlat
        GDI_Plus.GDIPlusDrawLineToDC dstDIB.getDIBDC, customX, m_SliderAreaHeight - 1, customX, m_SliderAreaHeight - 1 - notchSize, trackColor, 255, 1, True, LineCapFlat
        
    End If
    
End Sub

'When using a two-color or three-color gradient track style, this function can be called to recreate the background track DIB.
' Please note that this process is expensive (as we have to do per-pixel alpha masking of the final gradient), so please only
' call this function when absolutely necessary.
Private Sub redrawInternalGradientDIB()

    'Recreate the gradient DIB to the size of the background track area
    sizeDIBToTrackArea m_GradientDIB
    
    Dim trackRadius As Single
    trackRadius = (m_trackDiameter) \ 2
    
    Dim x As Long
    Dim relativeMiddlePosition As Single, tmpY As Single
    
    'Draw the gradient differently depending on the type of gradient
    Select Case curSliderStyle
    
        'Two-point gradients are the easiest; simply draw a gradient from left color to right color, the full width of the image
        Case GradientTwoPoint
           Drawing.DrawHorizontalGradientToDIB m_GradientDIB, trackRadius, m_GradientDIB.getDIBWidth - trackRadius, gradColorLeft, gradColorRight
        
        'Three-point gradients are more involved; draw a custom blend from left to middle to right, while accounting for the
        ' center point's position (which is variable, and which may change at run-time).
        Case GradientThreePoint
            
            'Calculate a relative pixel position for the supplied gradient middle value
            If (gradMiddleValue >= controlMin) And (gradMiddleValue <= controlMax) Then
                getCustomValueCoordinates gradMiddleValue, relativeMiddlePosition, tmpY
            Else
                relativeMiddlePosition = getTrackMinPos + ((getTrackMaxPos - getTrackMinPos) \ 2)
            End If
            
            'Draw two gradients; one each for the left and right of the gradient middle position
            Drawing.DrawHorizontalGradientToDIB m_GradientDIB, trackRadius, relativeMiddlePosition, gradColorLeft, gradColorMiddle
            Drawing.DrawHorizontalGradientToDIB m_GradientDIB, relativeMiddlePosition, m_GradientDIB.getDIBWidth - trackRadius, gradColorMiddle, gradColorRight
            
        'Hue gradients simply draw a full hue spectrum from 0 to 360.
        Case HueSpectrum360
        
            'From left-to-right, draw a full hue range onto the DIB
            Dim hueSpread As Long
            hueSpread = (m_GradientDIB.getDIBWidth - m_trackDiameter)
            
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
    ' termination point fall directly on the  maximum slider position (instead of the edge of the DIB, which would be
    ' incorrect as we need some padding for the rounded edge of the track area).  Note that hue gradients automatically
    ' handle this step during the DIB creation phase.
    If (curSliderStyle = GradientTwoPoint) Or (curSliderStyle = GradientThreePoint) Then
    
        For x = 0 To trackRadius
            GDI_Plus.GDIPlusDrawLineToDC m_GradientDIB.getDIBDC, x, 0, x, m_GradientDIB.getDIBHeight, gradColorLeft, 255, 1, False, LineCapFlat
        Next x
        
        For x = m_GradientDIB.getDIBWidth - trackRadius To m_GradientDIB.getDIBWidth
            GDI_Plus.GDIPlusDrawLineToDC m_GradientDIB.getDIBDC, x, 0, x, m_GradientDIB.getDIBHeight, gradColorRight, 255, 1, False, LineCapFlat
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
    GDI_Plus.GDIPlusDrawLineToDC alphaMask.getDIBDC, trackRadius, m_GradientDIB.getDIBHeight \ 2, m_GradientDIB.getDIBWidth - trackRadius, m_GradientDIB.getDIBHeight \ 2, 0, 255, m_trackDiameter - 1, True, LineCapRound
    
    'Transfer the alpha from the alpha mask to the gradient DIB itself
    'alphaMask.setAlphaPremultiplication False
    m_GradientDIB.copyAlphaFromExistingDIB alphaMask
    
    'Release the alpha-mask
    Set alphaMask = Nothing
    
    'Premultiply the gradient DIB, so we can successfully alpha-blend it later
    m_GradientDIB.setAlphaPremultiplication True
    
    'The gradient mask is now complete!
    
End Sub

'Because this control can contain either decimal or float values, we want to make sure any entered strings adhere
' to strict formatting rules.
Private Function getFormattedStringValue(ByVal srcValue As Double) As String

    Select Case significantDigits
    
        Case 0
            getFormattedStringValue = Format(CStr(srcValue), "#0")
        
        Case 1
            getFormattedStringValue = Format(CStr(srcValue), "#0.0")
            
        Case 2
            getFormattedStringValue = Format(CStr(srcValue), "#0.00")
            
        Case Else
            getFormattedStringValue = Format(CStr(srcValue), "#0.000")
    
    End Select
    
End Function

'Check a passed value against a min and max value to see if it is valid.  Additionally, make sure the value is
' numeric, and allow the user to display a warning message if necessary.  (As of v6.6, all validation is off-loaded
' to the embedded text up/down control.)
Private Function IsTextEntryValid(Optional ByVal displayErrorMsg As Boolean = False) As Boolean
    IsTextEntryValid = tudPrimary.IsValid(displayErrorMsg)
End Function

'Retrieve the current coordinates of the slider.  Note that the x/y pair returned references the slider's *center point*.
Private Sub getSliderCoordinates(ByRef sliderX As Single, ByRef sliderY As Single)
    
    'This dumb catch exists for when sliders are first loaded, and their max/min may both be zero.  This causes a divide-by-zero
    ' error in the horizontal slider position calculation, so if that happens, simply set the slider to its minimum position and exit.
    If controlMin <> controlMax Then
        
        'If an integer-only slider is in use, limit positions to whole numbers only
        If SigDigits = 0 Then
            sliderX = getTrackMinPos + ((Int(controlVal + 0.5) - controlMin) / (controlMax - controlMin)) * (getTrackMaxPos - getTrackMinPos)
        Else
            sliderX = getTrackMinPos + ((controlVal - controlMin) / (controlMax - controlMin)) * (getTrackMaxPos - getTrackMinPos)
        End If
        
        
    Else
        sliderX = getTrackMinPos
    End If
    
    sliderY = m_SliderAreaHeight \ 2
    
End Sub

'Retrieve the current coordinates of any custom value.  Note that the x/y pair returned are the custom value's *center point*.
Private Sub getCustomValueCoordinates(ByVal customValue As Single, ByRef customX As Single, ByRef customY As Single)
    
    'This dumb catch exists for when sliders are first loaded, and their max/min may both be zero.  This causes a divide-by-zero
    ' error in the horizontal slider position calculation, so if that happens, simply set the slider to its minimum position and exit.
    If controlMin <> controlMax Then
        customX = getTrackMinPos + ((customValue - controlMin) / (controlMax - controlMin)) * (getTrackMaxPos - getTrackMinPos)
    Else
        customX = getTrackMinPos
    End If
    
    customY = m_SliderAreaHeight \ 2
    
End Sub

'Returns a single increment amount for the current control.  The increment amount varies according to the significant digits setting;
' it can be as high as 1.0, or as low as 0.01.
Private Function getIncrementAmount() As Double
    getIncrementAmount = 1 / (10 ^ significantDigits)
End Function

'Return the min/max position of the track behind the slider.  This is used for a lot of things: rendering the track, calculating the
' value of the slider during user interactions (by determing the slider position relative to these two values), etc.  The minimum
' position is constant once the control is created, but the max position can change if the control size changes.
Private Function getTrackMinPos() As Long
    getTrackMinPos = m_sliderDiameter \ 2 + 2
End Function

Private Function getTrackMaxPos() As Long
    getTrackMaxPos = m_SliderAreaWidth - (m_sliderDiameter \ 2) - 2
End Function

'Given a user-supplied DIB, resize it to the area of the background track.  When using a custom-drawn slider, first call this function
' (and supply your owner-drawn DIB, obviously), so you know how big of an area is required.
Public Sub sizeDIBToTrackArea(ByRef targetDIB As pdDIB)
    
    Set targetDIB = New pdDIB
    targetDIB.createBlank (getTrackMaxPos - getTrackMinPos) + m_trackDiameter, m_SliderAreaHeight, 32, ConvertSystemColor(vbWindowBackground), 255
    
End Sub

'After a component of this control gets or loses focus, it needs to call this function.  This function is responsible for raising
' Got/LostFocusAPI events, which are important as an API text box is part of this control.
Private Sub evaluateFocusCount(ByVal focusCountJustIncremented As Boolean)

    If focusCountJustIncremented Then
        
        'If just incremented from 0 to 1, raise a GotFocusAPI event
        If m_ControlFocusCount = 1 Then RaiseEvent GotFocusAPI
        
    Else
    
        'If just decremented from 1 to 0, raise a LostFocusAPI event
        If m_ControlFocusCount = 0 Then RaiseEvent LostFocusAPI
    
    End If

End Sub

'External functions can call this to request a redraw.  This is helpful for live-updating theme settings, as in the Preferences dialog.
Public Sub UpdateAgainstCurrentTheme()
    
    'The text up/down can redraw itself
    tudPrimary.UpdateAgainstCurrentTheme
    
    If g_IsProgramRunning Then
            
        'Our tooltip object must also be refreshed (in case the language has changed)
        toolTipManager.UpdateAgainstCurrentTheme
        
        'The caption manager will also refresh itself
        m_Caption.UpdateAgainstCurrentTheme
        
    End If
    
    'Update the control's layout to account for new translations and/or theme changes
    updateControlLayout
    
    'Redraw the control to match any updated settings
    redrawSlider
    
End Sub

'Due to complex interactions between user controls and PD's translation engine, tooltips require this dedicated function.
' (IMPORTANT NOTE: the tooltip class will handle translations automatically.  Always pass the original English text!)
Public Sub AssignTooltip(ByVal newTooltip As String, Optional ByVal newTooltipTitle As String, Optional ByVal newTooltipIcon As TT_ICON_TYPE = TTI_NONE)
    toolTipManager.setTooltip Me.hWnd, UserControl.containerHwnd, newTooltip, newTooltipTitle, newTooltipIcon
    toolTipManager.setTooltip picScroll.hWnd, UserControl.containerHwnd, newTooltip, newTooltipTitle, newTooltipIcon
End Sub
