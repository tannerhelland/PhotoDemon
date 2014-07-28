VERSION 5.00
Begin VB.UserControl sliderTextCombo 
   AutoRedraw      =   -1  'True
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
   Begin VB.PictureBox picScroll 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   120
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   329
      TabIndex        =   1
      Top             =   60
      Width           =   4935
   End
   Begin VB.TextBox txtPrimary 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Left            =   5160
      TabIndex        =   0
      Text            =   "0"
      Top             =   60
      Width           =   735
   End
   Begin VB.Shape shpError 
      BorderColor     =   &H000000FF&
      Height          =   465
      Left            =   5100
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
End
Attribute VB_Name = "sliderTextCombo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Text / Slider custom control
'Copyright ©2013-2014 by Tanner Helland
'Created: 19/April/13
'Last updated: 21/June/14
'Last update: dump the scroll bar; implement my own owner-drawn slider
'
'Software like PhotoDemon requires a lot of UI elements.  Ideally, every setting should be adjustable by at least
' two mechanisms: direct text entry, and some kind of slider or scroll bar, which allows for a quick method to
' make both large and small adjustments to a given parameter.
'
'Historically, I accomplished this by providing a scroll bar and text box for every parameter in the program.
' This got the job done, but it had a number of limitations - such as requiring an enormous amount of time if
' changes ever needed to be made, and custom code being required in every form to handle text / scroll synching.
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
' 1) Synching of text and scroll/slide values
' 2) Validation of text entries, including a function for external validation requests
' 3) Locale handling (like the aforementioned comma/decimal replacement in some countries)
' 4) A single "Change" event that fires for either scroll or text changes, and only if a text change is valid
' 5) Support for integer or floating-point values via the "SigDigits" property
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'This object provides a single raised event:
' - Change (which triggers when either the scrollbar or text box is modified in any way)
Public Event Change()

'This control supports font setting for the text box
Private WithEvents mFont As StdFont
Attribute mFont.VB_VarHelpID = -1

'Forecolor handling is not currently handled, but it may be in the future
Private origForecolor As Long

'Special classes are used to render themed and multiline tooltips
Private m_ToolTip As clsToolTip
Private m_ToolString As String

'Used to internally track value, min, and max values as floating-points
Private controlVal As Double, controlMin As Double, controlMax As Double

'The number of significant digits for this control.  0 means integer values.
Private significantDigits As Long

'If the text box is initiating a value change, we must track that so as to not overwrite the user's entry mid-typing
Private textBoxInitiated As Boolean

'Mouse input handler
Private WithEvents cMouseEvents As pdInput
Attribute cMouseEvents.VB_VarHelpID = -1

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

'Internal gradient DIB.  This is recreated as necessary to reflect the gradient colors and positions.
Private m_GradientDIB As pdDIB

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
    
'If the current text value is NOT valid, this will return FALSE
Public Property Get IsValid(Optional ByVal showError As Boolean = True) As Boolean
    
    Dim retVal As Boolean
    retVal = Not shpError.Visible
    
    'If the current text value is not valid, highlight the problem and optionally display an error message box
    If Not retVal Then
        AutoSelectText txtPrimary
        If showError Then IsTextEntryValid True
    End If
    
    IsValid = retVal
    
End Property

'The Enabled property is a bit unique; see http://msdn.microsoft.com/en-us/library/aa261357%28v=vs.60%29.aspx
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_UserMemId = -514
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal newValue As Boolean)
    
    UserControl.Enabled = newValue
    
    'Disable the text box
    txtPrimary.Enabled = newValue
    txtPrimary.Refresh
    
    'Redraw the slider; when disabled, the slider itself is not drawn (only the track behind it is)
    redrawSlider
    
    PropertyChanged "Enabled"
    
End Property

'Font handling is a bit specialized for user controls; see http://msdn.microsoft.com/en-us/library/aa261313%28v=vs.60%29.aspx
Public Property Get Font() As StdFont
Attribute Font.VB_ProcData.VB_Invoke_Property = "StandardFont;Font"
Attribute Font.VB_UserMemId = -512
    Set Font = mFont
End Property

Public Property Set Font(mNewFont As StdFont)
    With mFont
        .Bold = mNewFont.Bold
        .Italic = mNewFont.Italic
        .Name = mNewFont.Name
        .Size = mNewFont.Size
    End With
    PropertyChanged "Font"
End Property

'Arrow keys can be used to "nudge" the control value in single-unit increments.
Private Sub cMouseEvents_KeyDownArrows(ByVal Shift As ShiftConstants, ByVal upArrow As Boolean, ByVal rightArrow As Boolean, ByVal downArrow As Boolean, ByVal leftArrow As Boolean)
    If upArrow Or rightArrow Then Value = Value + getIncrementAmount
    If leftArrow Or downArrow Then Value = Value - getIncrementAmount
End Sub

Private Sub cMouseEvents_MouseDownCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    
    If ((Button And pdLeftButton) <> 0) And isMouseOverSlider(x, y) Then
    
        m_MouseDown = True
        
        'Retrieve the current slider x/y values, and store the mouse position relative to those values
        Dim sliderX As Single, sliderY As Single
        getSliderCoordinates sliderX, sliderY
        m_InitX = x - sliderX
        m_InitY = y - sliderY
    
    End If
    
End Sub

Private Sub cMouseEvents_MouseLeave(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    cMouseEvents.setSystemCursor IDC_ARROW
End Sub

Private Sub cMouseEvents_MouseMoveCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)

    'If the mouse is down, adjust the current control value accordingly.
    If m_MouseDown Then
                
        'Calculate a new control value relative to the current mouse position
        Value = (controlMax - controlMin) * (((x + m_InitX) - getTrackMinPos) / (getTrackMaxPos - getTrackMinPos)) + controlMin
            
    'If the LMB is not down, modify the cursor according to its position relative to the slider
    Else
    
        If isMouseOverSlider(x, y) Then
            cMouseEvents.setSystemCursor IDC_HAND
        Else
            cMouseEvents.setSystemCursor IDC_ARROW
        End If
    
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

Private Function isMouseOverSlider(ByVal mouseX As Single, ByVal mouseY As Single) As Boolean

    'Retrieve the current x/y position of the slider's CENTER
    Dim sliderX As Single, sliderY As Single
    getSliderCoordinates sliderX, sliderY
    
    'See if the mouse is within distance of the slider's center
    If distanceTwoPoints(sliderX, sliderY, mouseX, mouseY) < fixDPI(SLIDER_DIAMETER) \ 2 Then
        isMouseOverSlider = True
    Else
        isMouseOverSlider = False
    End If

End Function

'When the font is updated, change the text box font to match.  (We also change the user control font, but this doesn't do anything... yet!)
Private Sub mFont_FontChanged(ByVal PropertyName As String)
    Set UserControl.Font = mFont
    Set txtPrimary.Font = UserControl.Font
End Sub

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
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
            If IsTextEntryValid(False) Then
                
                'The text box contains valid numerical data.  If it matches the current control value, skip the refresh step.
                If StrComp(getFormattedStringValue(txtPrimary), CStr(controlVal), vbBinaryCompare) <> 0 Then
                    txtPrimary.Text = getFormattedStringValue(controlVal)
                    txtPrimary.Refresh
                End If
            
            'The text box is currently in an error state.  Copy the new text into place without a duplication check.
            Else
            
                If shpError.Visible Then shpError.Visible = False
                txtPrimary.Text = getFormattedStringValue(controlVal)
                txtPrimary.Refresh
            
            End If
            
        End If
                
        'Redraw the slider to reflect the new value
        redrawSlider
        
        'Mark the value property as being changed, and raise the corresponding event.
        PropertyChanged "Value"
        RaiseEvent Change
        
    End If
    
End Property

'Note: the control's minimum value is settable at run-time
Public Property Get Min() As Double
    Min = controlMin
End Property

Public Property Let Min(ByVal newValue As Double)
    
    controlMin = newValue
    
    'If the track style is some kind of custom gradient, recreate our internal gradient DIB now
    If (curSliderStyle = GradientTwoPoint) Or (curSliderStyle = GradientThreePoint) Then redrawInternalGradientDIB
    
    'If the current control value is less than the new minimum, update it to match (and raise a corresponding _Change event)
    If controlVal < controlMin Then Value = controlMin
        
    PropertyChanged "Min"
    
End Property

'Note: the control's maximum value is settable at run-time
Public Property Get Max() As Double
    Max = controlMax
End Property

Public Property Let Max(ByVal newValue As Double)
    
    controlMax = newValue
    
    'If the track style is some kind of custom gradient, recreate our internal gradient DIB now
    If (curSliderStyle = GradientTwoPoint) Or (curSliderStyle = GradientThreePoint) Then redrawInternalGradientDIB
    
    'If the current control value is greater than the new max, update it to match (and raise a corresponding _Change event)
    If controlVal > controlMax Then Value = controlMax
    
    PropertyChanged "Max"
    
End Property

'Significant digits determines whether the control allows float values or int values (and with how much precision)
Public Property Get SigDigits() As Long
    SigDigits = significantDigits
End Property

Public Property Let SigDigits(ByVal newValue As Long)
    significantDigits = newValue
    PropertyChanged "SigDigits"
End Property

'Forecolor may be used in the future as part of theming, but right now it serves no purpose
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = origForecolor
End Property

Public Property Let ForeColor(ByVal newColor As OLE_COLOR)
    origForecolor = newColor
    PropertyChanged "ForeColor"
End Property

Private Sub txtPrimary_GotFocus()
    AutoSelectText txtPrimary
End Sub

Private Sub txtPrimary_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If IsTextEntryValid() Then
        If shpError.Visible Then shpError.Visible = False
        textBoxInitiated = True
        Me.Value = CDblCustom(txtPrimary)
        textBoxInitiated = False
    Else
        shpError.Visible = True
    End If
    
End Sub

Private Sub UserControl_Initialize()
    
    'When compiled, manifest-themed controls need to be further subclassed so they can have transparent backgrounds.
    If g_IsProgramCompiled And g_IsThemingEnabled And g_IsVistaOrLater Then g_Themer.requestContainerSubclass UserControl.hWnd
    
    'When not in design mode, initialize a tracker for mouse events
    If g_UserModeFix Then
        Set cMouseEvents = New pdInput
        cMouseEvents.addInputTracker picScroll.hWnd, True, True, , True
        cMouseEvents.setSystemCursor IDC_HAND
        cMouseEvents.requestKeyTracking picScroll.hWnd
        cMouseEvents.setKeyTrackers picScroll.hWnd, True
    End If
    
    'Update the control-level track and slider diameters to reflect current screen DPI
    m_trackDiameter = fixDPI(TRACK_DIAMETER)
    m_sliderDiameter = fixDPI(SLIDER_DIAMETER)
    
    'Set slider area width/height
    m_SliderAreaWidth = picScroll.ScaleWidth
    m_SliderAreaHeight = picScroll.ScaleHeight
    
    'Forecolor tracking may be supported in the future, but for now it's irrelevant
    origForecolor = ForeColor
        
    'Prepare a font object for use
    Set mFont = New StdFont
    Set UserControl.Font = mFont
    
End Sub

'Initialize control properties for the first time
Private Sub UserControl_InitProperties()

    'Reset all controls to their default state.  For each public property, matching internal tracker variables are also updated;
    ' this is not necessary, but it's helpful for reminding me of the names of the internal tracker variables relevant to their
    ' connected property.
    Set mFont = UserControl.Font
    mFont.Name = "Tahoma"
    mFont.Size = 10
    mFont_FontChanged ("")
    
    ForeColor = &H404040
    origForecolor = ForeColor
    
    Value = 0
    controlVal = 0
    
    Min = 0
    controlMin = 0
    
    Max = 10
    controlMax = 10
    
    SigDigits = 0
    significantDigits = 0
    
    SliderTrackStyle = DefaultStyle
    curSliderStyle = DefaultStyle
    
    'These default gradient values are useless; if you're using a gradient style, MAKE CERTAIN TO SPECIFY ACTUAL COLORS!
    GradientColorLeft = RGB(0, 0, 0)
    gradColorLeft = RGB(0, 0, 0)
    
    GradientColorRight = RGB(255, 255, 25)
    gradColorRight = RGB(255, 255, 255)
    
    GradientColorMiddle = RGB(121, 131, 135)
    gradColorMiddle = RGB(121, 131, 135)
    
    'This default gradient middle value is useless; if you use the 3-color gradient style, MAKE CERTAIN TO SPECIFY THIS VALUE!
    GradientMiddleValue = 0
    gradMiddleValue = 0
    
    'Default notch position; for most controls, it should be set to AUTOMATIC.  If CUSTOM is set, make sure to supply whatever
    ' custom value you want in the corresponding property!
    NotchPosition = AutomaticPosition
    curNotchPosition = AutomaticPosition
    
    NotchValueCustom = 0
    customNotchValue = 0
    
End Sub

'Read control properties from file
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    With PropBag
        Set Font = .ReadProperty("Font", Ambient.Font)
        ForeColor = .ReadProperty("ForeColor", &H404040)
        Min = .ReadProperty("Min", 0)
        Max = .ReadProperty("Max", 10)
        SigDigits = .ReadProperty("SigDigits", 0)
        SliderTrackStyle = .ReadProperty("SliderTrackStyle", DefaultStyle)
        Value = .ReadProperty("Value", 0)
        GradientColorLeft = .ReadProperty("GradientColorLeft", RGB(0, 0, 0))
        GradientColorRight = .ReadProperty("GradientColorRight", RGB(255, 255, 255))
        GradientColorMiddle = .ReadProperty("GradientColorMiddle", RGB(121, 131, 135))
        GradientMiddleValue = .ReadProperty("GradientMiddleValue", 0)
        NotchPosition = .ReadProperty("NotchPosition", 0)
        NotchValueCustom = .ReadProperty("NotchValueCustom", 0)
    End With
    
    'These values should have already been set by their respective property read, but I list them here to help me remember
    ' which internal values correspond to which properties:
    'controlMin = Min
    'controlMax = Max
    'controlVal = Value
    'significantDigits = SigDigits
    'curSliderStyle = SliderTrackStyle
    '
    'gradColorLeft = GradientColorLeft
    'gradColorRight = GradientColorRight
    'gradColorMiddle = GradientColorMiddle
    'gradMiddleValue = GradientMiddleValue
    
    'curNotchPosition = NotchPosition
    'customNotchValue = NotchValueCustom
    
End Sub

Private Sub UserControl_Resize()

    'We want to keep the text box and scroll bar universally aligned.  Thus, I have hard-coded specific spacing values.
    txtPrimary.Left = UserControl.ScaleWidth - fixDPI(56)
    shpError.Left = txtPrimary.Left - fixDPI(4)
    
    'It's possible - but obviously not recommended - to shrink the control so much that the scroll bar is invisible.
    ' Please do not do this.
    If txtPrimary.Left - fixDPI(15) > 0 Then picScroll.Width = txtPrimary.Left - fixDPI(15)         '15 = 8 (scroll bar's .Left) + 7 (distance between scroll bar and text box)
    
    'Update slider area width/height to match the new picScroll size
    m_SliderAreaWidth = picScroll.ScaleWidth
    m_SliderAreaHeight = picScroll.ScaleHeight
    
    'If the track style is some kind of custom gradient, recreate our internal gradient DIB now
    If (curSliderStyle = GradientTwoPoint) Or (curSliderStyle = GradientThreePoint) Or (curSliderStyle = HueSpectrum360) Then redrawInternalGradientDIB
    
    'Redraw the control
    redrawSlider

End Sub

Private Sub UserControl_Show()
    
    'When the control is first made visible, remove the control's tooltip property and reassign it to the checkbox
    ' using a custom solution (which allows for linebreaks and theming).
    If Len(Extender.ToolTipText) > 0 Then assignTooltip Extender.ToolTipText
    
    'If the track style is some kind of custom gradient, recreate our internal gradient DIB now
    If (curSliderStyle = GradientTwoPoint) Or (curSliderStyle = GradientThreePoint) Or (curSliderStyle = HueSpectrum360) Then redrawInternalGradientDIB
    
    redrawSlider
        
End Sub

Private Sub UserControl_Terminate()
    
    'When the control is terminated, release the subclassing used for transparent backgrounds
    If g_IsProgramCompiled And g_IsThemingEnabled And g_IsVistaOrLater Then g_Themer.releaseContainerSubclass UserControl.hWnd
    
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    'Store all associated properties
    With PropBag
        .WriteProperty "Font", mFont, "Tahoma"
        .WriteProperty "ForeColor", ForeColor, &H404040
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

'Render a custom slider to the slider area picture box
Private Sub redrawSlider()

    'All drawing is done to a temporary DIB, which is flipped to the screen as the final rendering step
    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB
    
    tmpDIB.createBlank m_SliderAreaWidth, m_SliderAreaHeight, 32, 0
    
    'There are a few components to the slider:
    ' 1) The track that sits behind the slider.  It has two relevant parameters: a radius, and a color.  Its width is automatically
    '     calculated relevant to the width of the control as a whole.
    ' 2) The slider that sits atop the track.  It has two relevant parameters: a radius, a fill color, and an edge color.  Its width
    '     is constant from a programmatic standpoint, though it does get updated at run-time to account for screen DPI.
    
    'Set colors first.  In the future, these will be handled via a theming engine (so PD can support light-on-dark or dark-on-light
    ' themes, etc), but for now, set them manually.
    Dim trackColor As Long, sliderBackgroundColor As Long, sliderEdgeColor As Long
    trackColor = RGB(177, 186, 194)
    sliderBackgroundColor = RGB(255, 255, 255)
    sliderEdgeColor = RGB(60, 175, 230)
    
    'Retrieve the current slider x/y position.  Floating-point values are used so we can support sub-pixel positioning!
    Dim relevantSliderPosX As Single, relevantSliderPosY As Single
    getSliderCoordinates relevantSliderPosX, relevantSliderPosY
    
    'Additional draw variables utilized by multiple render styles
    Dim customX As Single, customY As Single
    Dim relevantMin As Single
    
    'Draw the background track according to the current SliderTrackStyle property.
    If Me.Enabled Then
    
        'This control supports a variety of different track styles.  Some of these styles require a DIB supplied by the owner, and
        ' they *will not* render properly until that DIB is provided!
        Select Case curSliderStyle
        
            'Default style: fill the "active" part of track with the control highlight color.  The "active part" is the chunk relative
            ' to zero, if the control supports 0 as a value; otherwise, it is relative to the control minimum.
            Case DefaultStyle
            
                'Start by drawing the default background track
                GDI_Plus.GDIPlusDrawLineToDC tmpDIB.getDIBDC, getTrackMinPos, m_SliderAreaHeight \ 2, getTrackMaxPos, m_SliderAreaHeight \ 2, trackColor, 255, m_trackDiameter + 1, True, LineCapRound
                
                'Next, determine a minimum value for the control, using the formula provided:
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
                GDI_Plus.GDIPlusDrawLineToDC tmpDIB.getDIBDC, customX, customY, relevantSliderPosX, customY, sliderEdgeColor, 255, m_trackDiameter + 1, True, LineCapRound
                
            'No-frills slider: plain gray background (boooring - use only if absolutely necessary)
            Case NoFrills
                GDI_Plus.GDIPlusDrawLineToDC tmpDIB.getDIBDC, getTrackMinPos, m_SliderAreaHeight \ 2, getTrackMaxPos, m_SliderAreaHeight \ 2, trackColor, 255, m_trackDiameter + 1, True, LineCapRound
            
            Case GradientTwoPoint, GradientThreePoint, HueSpectrum360
            
                'As a failsafe, make sure our internal gradient DIB exists
                If m_GradientDIB Is Nothing Then redrawInternalGradientDIB
                
                'Draw a stock trackline onto the target DIB.  This will serve as the border of the gradient track area.
                GDI_Plus.GDIPlusDrawLineToDC tmpDIB.getDIBDC, getTrackMinPos, m_SliderAreaHeight \ 2, getTrackMaxPos, m_SliderAreaHeight \ 2, trackColor, 255, m_trackDiameter + 1, True, LineCapRound
                
                'Next, draw the gradient effect DIB to the location where we'd normally draw the track line.  Alpha has already been
                ' calculated for the gradient DIB, so it will sit precisely inside the trackline drawn above, giving the track a
                ' sharp 1px border.
                m_GradientDIB.alphaBlendToDC tmpDIB.getDIBDC, 255, getTrackMinPos - (m_trackDiameter \ 2), 0
                
            Case CustomOwnerDrawn
        
        End Select
        
        'Before carrying on, draw a slight notch above and below the slider track, using the value specified by the associated property
        drawNotchToDIB tmpDIB, trackColor
        
    'Control is disabled; draw a plain track in the background, but no notch or other frills
    Else
        GDI_Plus.GDIPlusDrawLineToDC tmpDIB.getDIBDC, getTrackMinPos, m_SliderAreaHeight \ 2, getTrackMaxPos, m_SliderAreaHeight \ 2, trackColor, 255, m_trackDiameter + 1, True, LineCapRound
    End If
        
    'The slider itself is only drawn if the control is enabled; otherwise, we do not display it at all.
    If Me.Enabled Then
    
        'Draw the background (interior fill) circle of the slider
        GDI_Plus.GDIPlusDrawEllipseToDC tmpDIB.getDIBDC, relevantSliderPosX - (m_sliderDiameter \ 2), relevantSliderPosY - (m_sliderDiameter \ 2), m_sliderDiameter, m_sliderDiameter, sliderBackgroundColor, True
        
        'Draw the edge (exterior) circle around the slider
        GDI_Plus.GDIPlusDrawCircleToDC tmpDIB.getDIBDC, relevantSliderPosX, relevantSliderPosY, m_sliderDiameter \ 2, sliderEdgeColor, 255, 1.5, True
        
    End If
    
    'Composite the slider buffer against the specified background color.  In the future, the background color will be set by PD's theming engine,
    ' but for now it is hard-coded against the standard "window background" color.
    Dim backDIB As pdDIB
    Set backDIB = New pdDIB
    backDIB.createBlank m_SliderAreaWidth, m_SliderAreaHeight, 24, ConvertSystemColor(vbWindowBackground)
    tmpDIB.alphaBlendToDC backDIB.getDIBDC, 255
    
    'Flip the fully composited scroller image onto the owner picture box
    BitBlt picScroll.hDC, 0, 0, m_SliderAreaWidth, m_SliderAreaHeight, backDIB.getDIBDC, 0, 0, vbSrcCopy
    picScroll.Picture = picScroll.Image
    picScroll.Refresh
    
    Set backDIB = Nothing
    Set tmpDIB = Nothing

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
    ' incorrect sa we need some padding for the rounded edge of the track area).  Note that hue gradients automatically
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
    'alphaMask.fixPremultipliedAlpha False
    m_GradientDIB.copyAlphaFromExistingDIB alphaMask
    
    'Release the alpha-mask
    Set alphaMask = Nothing
    
    'Premultiply the gradient DIB, so we can successfully alpha-blend it later
    m_GradientDIB.fixPremultipliedAlpha True
    
    'The gradient mask is now complete!
    
End Sub

'To workaround a translation issue, the control's original English text can be manually backed up; this allows us
' to change the language at run-time and still have translations work as expected.
Public Sub assignTooltip(ByVal newTooltip As String)
    m_ToolString = newTooltip
    If Len(m_ToolString) > 0 Then refreshTooltipObject
End Sub

'When the program language is changed, the object's tooltip must be retranslated to match.  External functions can
' call this sub to have it automatically fixed.
Public Sub refreshTooltipObject()
    
    If Not (m_ToolTip Is Nothing) Then
        m_ToolTip.RemoveTool picScroll
    End If
    
    Set m_ToolTip = New clsToolTip
    With m_ToolTip
        .Create Me
        .MaxTipWidth = PD_MAX_TOOLTIP_WIDTH
        .DelayTime(ttDelayShow) = 10000
        If g_Language.translationActive Then
            .AddTool picScroll, g_Language.TranslateMessage(m_ToolString)
        Else
            .AddTool picScroll, m_ToolString
        End If
    End With
        
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
' numeric, and allow the user to display a warning message if necessary.
Private Function IsTextEntryValid(Optional ByVal displayErrorMsg As Boolean = False) As Boolean
    
    'Some locales use a comma as a decimal separator.  Check for this and replace as necessary.
    Dim chkString As String
    chkString = txtPrimary
    
    'Remember the current cursor position as necessary
    Dim cursorPos As Long
    cursorPos = txtPrimary.SelStart
    
    'It may be possible for the user to enter consecutive ",." characters, which then cause the CDbl() below to fail.
    ' Check for this and fix it as necessary.
    If InStr(1, chkString, "..") Then
        chkString = Replace(chkString, "..", ".")
        txtPrimary = chkString
        If cursorPos >= Len(txtPrimary) Then cursorPos = Len(txtPrimary)
        txtPrimary.SelStart = cursorPos
    End If
    
    'If our owner wants a message displayed on invalid input, raise one now.
    If Not IsNumeric(chkString) Then
        If displayErrorMsg Then pdMsgBox "%1 is not a valid entry." & vbCrLf & "Please enter a numeric value.", vbExclamation + vbOKOnly + vbApplicationModal, "Invalid entry", txtPrimary
        IsTextEntryValid = False
    Else
        
        Dim checkVal As Double
        checkVal = CDblCustom(chkString)
    
        If (checkVal >= controlMin) And (checkVal <= controlMax) Then
            IsTextEntryValid = True
        Else
            If displayErrorMsg Then pdMsgBox "%1 is not a valid entry." & vbCrLf & "Please enter a value between %2 and %3.", vbExclamation + vbOKOnly + vbApplicationModal, "Invalid entry", txtPrimary, getFormattedStringValue(controlMin), getFormattedStringValue(controlMax)
            IsTextEntryValid = False
        End If
    End If
    
End Function

'Retrieve the current coordinates of the slider.  Note that the x/y pair returned are the slider's *center point*.
Private Sub getSliderCoordinates(ByRef sliderX As Single, ByRef sliderY As Single)
    
    'This dumb catch exists for when sliders are first loaded, and their max/min may both be zero.  This causes a divide-by-zero
    ' error in the horizontal slider position calculation, so if that happens, simply set the slider to its minimum position and exit.
    If controlMin <> controlMax Then
        sliderX = getTrackMinPos + ((controlVal - controlMin) / (controlMax - controlMin)) * (getTrackMaxPos - getTrackMinPos)
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
