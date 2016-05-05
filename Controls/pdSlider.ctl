VERSION 5.00
Begin VB.UserControl pdSlider 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6000
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
   MousePointer    =   99  'Custom
   ScaleHeight     =   33
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   400
   ToolboxBitmap   =   "pdSlider.ctx":0000
   Begin PhotoDemon.pdSliderStandalone pdssPrimary 
      Height          =   360
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   635
   End
   Begin PhotoDemon.pdSpinner tudPrimary 
      Height          =   345
      Left            =   4800
      TabIndex        =   0
      Top             =   45
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   609
   End
End
Attribute VB_Name = "pdSlider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Slider+Spinner custom control
'Copyright 2013-2016 by Tanner Helland
'Created: 19/April/13
'Last updated: 12/February/16
'Last update: migrate slider-specific code into pdSliderStandalone, which hugely reduces the complexity of
'             synchronizing the slider and spinner elements of this control.
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
' - Change (which triggers when the slider or spinner values are modified by any mechanism)
Public Event Change()

'Because VB focus events are wonky, especially when we use CreateWindow within a UC, this control raises its own
' specialized focus events.  If you need to track focus, use these instead of the default VB functions.
Public Event GotFocusAPI()
Public Event LostFocusAPI()

'If the text box is initiating a value change, we must track that so as to not overwrite the user's entry mid-typing
Private textBoxInitiated As Boolean

'User control support class.  Historically, many classes (and associated subclassers) were required by each user control,
' but I've since attempted to wrap these into a single master control support class.
Private WithEvents ucSupport As pdUCSupport
Attribute ucSupport.VB_VarHelpID = -1

'Tracks whether the control (any component) has focus.  This is helpful as we must synchronize between VB's focus events and API
' focus events.  Every time an individual component gains focus, we increment this counter by 1.  Every time an individual component
' loses focus, we decrement the counter by 1.  When the counter hits 0, we report a control-wide Got/LostFocusAPI event.
Private m_ControlFocusCount As Long

'Used to prevent recursive redraws
Private m_InternalResizeActive As Boolean

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
End Property

'The Enabled property is a bit unique; see http://msdn.microsoft.com/en-us/library/aa261357%28v=vs.60%29.aspx
Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal newValue As Boolean)
    pdssPrimary.Enabled = newValue
    tudPrimary.Enabled = newValue
    UserControl.Enabled = newValue
    PropertyChanged "Enabled"
End Property

Public Property Get FontSizeCaption() As Single
    FontSizeCaption = ucSupport.GetCaptionFontSize
End Property

Public Property Let FontSizeCaption(ByVal newSize As Single)
    ucSupport.SetCaptionFontSize newSize
    PropertyChanged "FontSizeCaption"
End Property

Public Property Get FontSizeTUD() As Single
    FontSizeTUD = tudPrimary.FontSize
End Property

Public Property Let FontSizeTUD(ByVal newSize As Single)
    If newSize <> tudPrimary.FontSize Then
        tudPrimary.FontSize = newSize
        PropertyChanged "FontSizeTUD"
    End If
End Property

'Gradient colors.  For the two-color gradient style, only colors Left and Right are relevant.  Color Middle is used for the
' 3-color style only, and note that it *must* be accompanied by an owner-supplied middle position value.
Public Property Get GradientColorLeft() As OLE_COLOR
    GradientColorLeft = pdssPrimary.GradientColorLeft
End Property

Public Property Get GradientColorMiddle() As OLE_COLOR
    GradientColorMiddle = pdssPrimary.GradientColorMiddle
End Property

Public Property Get GradientColorRight() As OLE_COLOR
    GradientColorRight = pdssPrimary.GradientColorRight
End Property

Public Property Let GradientColorLeft(ByVal newColor As OLE_COLOR)
    pdssPrimary.GradientColorLeft = newColor
    PropertyChanged "GradientColorLeft"
End Property

Public Property Let GradientColorMiddle(ByVal newColor As OLE_COLOR)
    pdssPrimary.GradientColorMiddle = newColor
    PropertyChanged "GradientColorMiddle"
End Property

Public Property Let GradientColorRight(ByVal newColor As OLE_COLOR)
    pdssPrimary.GradientColorRight = newColor
    PropertyChanged "GradientColorRight"
End Property

'Custom middle value for the 3-color gradient style.  This value is ignored for all other styles.
Public Property Get GradientMiddleValue() As Double
    GradientMiddleValue = pdssPrimary.GradientMiddleValue
End Property

Public Property Let GradientMiddleValue(ByVal newValue As Double)
    pdssPrimary.GradientMiddleValue = newValue
    PropertyChanged "GradientMiddleValue"
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
    Max = pdssPrimary.Max
End Property

Public Property Let Max(ByVal newValue As Double)
    pdssPrimary.Max = newValue
    tudPrimary.Max = newValue
    PropertyChanged "Max"
End Property

'Note: the control's minimum value is settable at run-time
Public Property Get Min() As Double
    Min = pdssPrimary.Min
End Property

Public Property Let Min(ByVal newValue As Double)
    pdssPrimary.Min = newValue
    tudPrimary.Min = newValue
    PropertyChanged "Min"
End Property

'Notch positioning technique.  If CUSTOM is set, make sure to supply a custom value to match!
Public Property Get NotchPosition() As SLIDER_NOTCH_POSITION
    NotchPosition = pdssPrimary.NotchPosition
End Property

Public Property Let NotchPosition(ByVal newPosition As SLIDER_NOTCH_POSITION)
    pdssPrimary.NotchPosition = newPosition
    PropertyChanged "NotchPosition"
End Property

'Custom notch value.  This value is only used if NotchPosition = CustomPosition.
Public Property Get NotchValueCustom() As Double
    NotchValueCustom = pdssPrimary.NotchValueCustom
End Property

Public Property Let NotchValueCustom(ByVal newValue As Double)
    pdssPrimary.NotchValueCustom = newValue
    PropertyChanged "NotchValueCustom"
End Property

'Significant digits determines whether the control allows float values or int values (and with how much precision)
Public Property Get SigDigits() As Long
    SigDigits = pdssPrimary.SigDigits
End Property

Public Property Let SigDigits(ByVal newValue As Long)
    pdssPrimary.SigDigits = newValue
    tudPrimary.SigDigits = newValue
    PropertyChanged "SigDigits"
End Property

Public Property Get SliderTrackStyle() As SLIDER_TRACK_STYLE
    SliderTrackStyle = pdssPrimary.SliderTrackStyle
End Property

Public Property Let SliderTrackStyle(ByVal newStyle As SLIDER_TRACK_STYLE)
    pdssPrimary.SliderTrackStyle = newStyle
    PropertyChanged "SliderTrackStyle"
End Property

'The control's value is simply a reflection of the embedded scroll bar and text box
Public Property Get Value() As Double
Attribute Value.VB_UserMemId = 0
    Value = pdssPrimary.Value
End Property

Public Property Let Value(ByVal newValue As Double)
    
    'Don't make any changes unless the new value deviates from the existing one
    If (newValue <> pdssPrimary.Value) Or (newValue <> tudPrimary.Value) Then
        pdssPrimary.Value = newValue
        tudPrimary.Value = newValue
        If Me.Enabled Then RaiseEvent Change
        PropertyChanged "Value"
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

Private Sub pdssPrimary_Change()
    Me.Value = pdssPrimary.Value
End Sub

Private Sub pdssPrimary_GotFocusAPI()
    m_ControlFocusCount = m_ControlFocusCount + 1
    EvaluateFocusCount True
End Sub

Private Sub pdssPrimary_LostFocusAPI()
    m_ControlFocusCount = m_ControlFocusCount - 1
    EvaluateFocusCount True
End Sub

Private Sub ucSupport_GotFocusAPI()
    m_ControlFocusCount = m_ControlFocusCount + 1
    EvaluateFocusCount True
End Sub

Private Sub ucSupport_LostFocusAPI()
    m_ControlFocusCount = m_ControlFocusCount - 1
    EvaluateFocusCount False
End Sub

Private Sub tudPrimary_Change()
    If (Not textBoxInitiated) Then
        If tudPrimary.IsValid(False) Then
            textBoxInitiated = True
            Me.Value = tudPrimary.Value
            textBoxInitiated = False
        End If
    End If
End Sub

Private Sub tudPrimary_GotFocusAPI()
    m_ControlFocusCount = m_ControlFocusCount + 1
    EvaluateFocusCount True
End Sub

Private Sub tudPrimary_LostFocusAPI()
    m_ControlFocusCount = m_ControlFocusCount - 1
    EvaluateFocusCount False
End Sub

Private Sub tudPrimary_Resize()
    UpdateControlLayout
End Sub

Private Sub ucSupport_WindowResize(ByVal newWidth As Long, ByVal newHeight As Long)
    If (Not m_InternalResizeActive) Then UpdateControlLayout
End Sub

Private Sub UserControl_Initialize()
    
    'Initialize a master user control support class
    Set ucSupport = New pdUCSupport
    ucSupport.RegisterControl UserControl.hWnd
    ucSupport.RequestCaptionSupport False
        
End Sub

'Initialize control properties for the first time
Private Sub UserControl_InitProperties()
    FontSizeTUD = 10
    FontSizeCaption = 12
    Caption = ""
        
    Min = 0
    Max = 10
    SigDigits = 0
    Value = 0
    
    SliderTrackStyle = DefaultStyle
    GradientColorLeft = RGB(0, 0, 0)
    GradientColorRight = RGB(255, 255, 25)
    GradientColorMiddle = RGB(121, 131, 135)
    GradientMiddleValue = 0
    NotchPosition = AutomaticPosition
    NotchValueCustom = 0
End Sub

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

'At run-time, painting is handled by PD's pdWindowPainter class.  In the IDE, however, we must rely on VB's internal paint event.
Private Sub UserControl_Paint()
    ucSupport.RequestIDERepaint UserControl.hDC
End Sub

Private Sub UserControl_Resize()
    If Not g_IsProgramRunning Then ucSupport.RequestRepaint True
End Sub

Private Sub UserControl_Show()
    UpdateControlLayout
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    With PropBag
        .WriteProperty "Caption", Me.Caption, ""
        .WriteProperty "FontSizeCaption", Me.FontSizeCaption, 12
        .WriteProperty "FontSizeTUD", Me.FontSizeTUD, 10
        .WriteProperty "Min", Me.Min, 0
        .WriteProperty "Max", Me.Max, 10
        .WriteProperty "SigDigits", Me.SigDigits, 0
        .WriteProperty "SliderTrackStyle", Me.SliderTrackStyle, DefaultStyle
        .WriteProperty "Value", Me.Value, 0
        .WriteProperty "GradientColorLeft", Me.GradientColorLeft, RGB(0, 0, 0)
        .WriteProperty "GradientColorRight", Me.GradientColorRight, RGB(255, 255, 255)
        .WriteProperty "GradientColorMiddle", Me.GradientColorMiddle, RGB(121, 131, 135)
        .WriteProperty "GradientMiddleValue", Me.GradientMiddleValue, 0
        .WriteProperty "NotchPosition", Me.NotchPosition, 0
        .WriteProperty "NotchValueCustom", Me.NotchValueCustom, 0
    End With
    
End Sub

'When the control is resized, the caption is changed, or font sizes for either the caption or text up/down are modified,
' this function should be called.  It controls the physical positioning of various control sub-elements
' (specifically, the caption area, the slider area, and the text up/down area).
Private Sub UpdateControlLayout()
    
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
    If ucSupport.IsCaptionActive Then
        
        'We now have all the information necessary to calculate caption positioning (and by extension, slider and
        ' text up/down positioning, too!)
        
        'Calculate a new height for the usercontrol as a whole.  This is simple formula:
        ' (height of text up/down) + (2 px padding around text up/down) + (height of caption) + (1 px padding around caption)
        Dim textHeight As Long
        textHeight = ucSupport.GetCaptionHeight
        newControlHeight = tudPrimary.GetHeight + FixDPI(4) + textHeight + FixDPI(2)
        
        'Calculate a new top position for the slider box (which will be vertically centered in the space below the caption)
        newTop_Slider = ((newControlHeight - (textHeight + FixDPI(4))) - tudPrimary.GetHeight) \ 2
        newTop_Slider = textHeight + FixDPI(4) + newTop_Slider
        
    'When a slider lacks a caption, we hard-code its height to preset values
    Else
        
        'Start by setting the control height
        newControlHeight = tudPrimary.GetHeight + FixDPI(4)
        
        'Center the slider box inside the newly calculated height
        newTop_Slider = (newControlHeight - pdssPrimary.GetHeight) \ 2
                
    End If
    
    'Apply the new height to this UC instance
    If (ucSupport.GetControlHeight <> newControlHeight) Then ucSupport.RequestNewSize , newControlHeight
    
    'With height correctly set, we next want to left-align the spinner against the slider region
    newLeft_TUD = ucSupport.GetControlWidth - (tudPrimary.GetWidth + FixDPI(2))
    
    'Because the slider width is contingent on the spinner position, calculate it next, then move it into place
    newWidth_Slider = newLeft_TUD - FixDPI(10)
    If (newTop_Slider <> pdssPrimary.GetTop) Then pdssPrimary.SetTop newTop_Slider
    If (newWidth_Slider > 0) And (newWidth_Slider <> pdssPrimary.GetWidth) Then pdssPrimary.SetWidth newWidth_Slider
    
    'Vertically center the spinner relative to the slider
    Dim sliderVerticalCenter As Single
    sliderVerticalCenter = pdssPrimary.GetTop + (CSng(pdssPrimary.GetHeight) / 2)
    newTop_TUD = sliderVerticalCenter - Int(CSng(tudPrimary.GetHeight) / 2) + 1
    
    'Now that we've calculated new text up/down positioning, we can apply it as necessary
    If (tudPrimary.GetTop <> newTop_TUD) Then tudPrimary.SetTop newTop_TUD
    If (tudPrimary.GetLeft <> newLeft_TUD) Then tudPrimary.SetLeft newLeft_TUD
    
    'Inside the IDE, use a line of dummy code to force a redraw of the control outline
    If (Not g_IsProgramRunning) Then
        Dim bufferDC As Long
        bufferDC = ucSupport.GetBackBufferDC(True)
    End If
    
    ucSupport.RequestRepaint
    m_InternalResizeActive = False
    
End Sub

'Check a passed value against a min and max value to see if it is valid.  Additionally, make sure the value is
' numeric, and allow the user to display a warning message if necessary.  (As of v6.6, all validation is off-loaded
' to the embedded text up/down control.)
Private Function IsTextEntryValid(Optional ByVal displayErrorMsg As Boolean = False) As Boolean
    IsTextEntryValid = tudPrimary.IsValid(displayErrorMsg)
End Function

'After a component of this control gets or loses focus, it needs to call this function.  This function is responsible for raising
' Got/LostFocusAPI events, which are important as an API text box is part of this control.
Private Sub EvaluateFocusCount(ByVal focusCountJustIncremented As Boolean)

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
    pdssPrimary.UpdateAgainstCurrentTheme
    tudPrimary.UpdateAgainstCurrentTheme
    ucSupport.UpdateAgainstThemeAndLanguage
    
    'Update the control's layout to account for new translations and/or theme changes
    UpdateControlLayout
End Sub

'Due to complex interactions between user controls and PD's translation engine, tooltips require this dedicated function.
' (IMPORTANT NOTE: the tooltip class will handle translations automatically.  Always pass the original English text!)
Public Sub AssignTooltip(ByVal newTooltip As String, Optional ByVal newTooltipTitle As String, Optional ByVal newTooltipIcon As TT_ICON_TYPE = TTI_NONE)
    pdssPrimary.AssignTooltip newTooltip, newTooltipTitle, newTooltipIcon
    tudPrimary.AssignTooltip newTooltip, newTooltipTitle, newTooltipIcon
End Sub



