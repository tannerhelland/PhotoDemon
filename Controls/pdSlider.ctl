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
   HasDC           =   0   'False
   ScaleHeight     =   33
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   400
   ToolboxBitmap   =   "pdSlider.ctx":0000
   Begin PhotoDemon.pdSliderStandalone pdssPrimary 
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   60
      Width           =   4335
      _ExtentX        =   8281
      _ExtentY        =   635
   End
   Begin PhotoDemon.pdSpinner tudPrimary 
      Height          =   345
      Left            =   4440
      TabIndex        =   1
      Top             =   45
      Width           =   1260
      _ExtentX        =   2223
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
'Copyright 2013-2026 by Tanner Helland
'Created: 19/April/13
'Last updated: 09/December/21
'Last update: new CaptionPadding property, so I can fix minor alignment annoyances on toolpanels
'
'Software like PhotoDemon requires a lot of UI elements.  Ideally, every setting should be adjustable
' by at least two mechanisms: direct text entry, and some kind of slider or scroll bar, which provides
' quick input for both large and small adjustments.  This slider control provides this capability across
' almost every dialog in the project.
'
'Generally speaking, this control handles the following things automatically:
' 1) Syncing of text and scroll/slide values
' 2) Validation of text entries, including a function for external validation requests
' 3) Locale handling (including comma/decimal translation across locales)
' 4) A single "Change" event that fires for either scroll or text changes, and only if a text change is valid
' 5) Support for integer or floating-point values via the "SigDigits" property
' 6) Several different drawing modes, including support for 2- or 3-point gradients
' 7) Self-captioning, to remove the need for a redundant label control next to this one
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Change vs FinalChange: change is fired whenever the scroller value changes at all (e.g. during every
' mouse movement); FinalChange is fired only when a mouse or key is released.  If the slider controls a
' particularly time-consuming operation, it may be preferable to lean on FinalChange instead of Change,
' but note the caveat that FinalChange *only* triggers on MouseUp/KeyUp - *not* on external .Value changes
' - so you may still need to handle the regular Change event, if you are externally setting values.
' This oddity is necessary because otherwise, the spinner and slider controls constantly trigger each
' other's .Value properties, causing endless FinalChange triggers.
Public Event Change()
Public Event FinalChange()
Public Event ResetClick()

'Because VB focus events are wonky, especially when we use CreateWindow within a UC, this control raises its own
' specialized focus events.  If you need to track focus, use these instead of the default VB functions.
Public Event GotFocusAPI()
Public Event LostFocusAPI()
Public Event SetCustomTabTarget(ByVal shiftTabWasPressed As Boolean, ByRef newTargetHwnd As Long)

'If this is an owner-drawn slider, the slider will raise events when it needs an updated track image.
' (This event is irrelevant for normal sliders.)
Public Event RenderTrackImage(ByRef dstDIB As pdDIB, ByVal leftBoundary As Single, ByVal rightBoundary As Single)

'If the text box is initiating a value change, we must track that so as to not overwrite the user's entry mid-typing
Private m_textBoxInitiated As Boolean

'Added in 2021 to allow for minor tweaks in caption vs scrollbar padding; this helps me perfectly
' align some tricky run-time elements in PhotoDemon's toolpanels.  (Can be negative.)
Private m_captionPadding As Long

'User control support class.  Historically, many classes (and associated subclassers) were required by each user control,
' but I've since wrapped these into a single central support class.
Private WithEvents ucSupport As pdUCSupport
Attribute ucSupport.VB_VarHelpID = -1

'Tracks whether the control (any component) has focus.  This is helpful as this control contains a number of child controls,
' and we want to raise focus events only if *none of our children* have focus (or alternatively, if *one of our children*
' gains focus).
Private m_LastFocusState As Boolean

'Used to prevent recursive redraws
Private m_InternalResizeActive As Boolean

Public Function GetControlType() As PD_ControlType
    GetControlType = pdct_Slider
End Function

Public Function GetControlName() As String
    GetControlName = UserControl.Extender.Name
End Function

'Workaround for VB6 quirks; see VBHacks.InControlArray()
Public Function IsChildInControlArray(ByRef ctlChild As Object) As Boolean
    IsChildInControlArray = Not UserControl.Controls(ctlChild.Name) Is ctlChild
End Function

'Caption is handled just like a label's caption property.  It is valid at design-time, and any translation,
' if present, will not be processed until run-time.
Public Property Get Caption() As String
Attribute Caption.VB_UserMemId = -518
    Caption = ucSupport.GetCaptionText
End Property

Public Property Let Caption(ByRef newCaption As String)
    ucSupport.SetCaptionText newCaption
    PropertyChanged "Caption"
End Property

Public Property Get CaptionPadding() As Long
    CaptionPadding = m_captionPadding
End Property

Public Property Let CaptionPadding(ByVal newPadding As Long)
    m_captionPadding = newPadding
    If PDMain.IsProgramRunning() Then UpdateControlLayout
    PropertyChanged "CaptionPadding"
End Property

Public Property Get DefaultValue() As Double
    DefaultValue = tudPrimary.DefaultValue
End Property

Public Property Let DefaultValue(ByVal newValue As Double)
    tudPrimary.DefaultValue = newValue
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
    If (newSize <> tudPrimary.FontSize) Then
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

Public Property Get HasFocus() As Boolean
    HasFocus = ucSupport.DoIHaveFocus() Or pdssPrimary.HasFocus() Or tudPrimary.HasFocus()
End Property

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

Public Property Get hWndSlider() As Long
    hWndSlider = pdssPrimary.hWnd
End Property

Public Property Get hWndSpinner() As Long
    hWndSpinner = tudPrimary.hWnd
End Property

'If the current text value is NOT valid, this will return FALSE.  Note that this property is read-only.
Public Property Get IsValid(Optional ByVal showError As Boolean = True) As Boolean
    IsValid = tudPrimary.IsValid(showError)
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

'Scale style determines whether the slider knob moves linearly, or exponentially
Public Property Get ScaleStyle() As PD_SLIDER_SCALESTYLE
    ScaleStyle = pdssPrimary.ScaleStyle
End Property

Public Property Let ScaleStyle(ByVal newStyle As PD_SLIDER_SCALESTYLE)
    pdssPrimary.ScaleStyle = newStyle
End Property

Public Property Get ScaleExponent() As Single
    ScaleExponent = pdssPrimary.ScaleExponent
End Property

Public Property Let ScaleExponent(ByVal newExponent As Single)
    pdssPrimary.ScaleExponent = newExponent
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

Public Property Get SliderKnobStyle() As SLIDER_KNOB_STYLE
    SliderKnobStyle = pdssPrimary.SliderKnobStyle
End Property

Public Property Let SliderKnobStyle(ByVal newStyle As SLIDER_KNOB_STYLE)
    pdssPrimary.SliderKnobStyle = newStyle
    PropertyChanged "SliderKnobStyle"
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

Public Sub Reset()
    tudPrimary.Reset
End Sub

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

'Shortcut function for setting the slider L/R gradient colors and Value all at once.  I've gone back and forth on the
' best way to do this (in PD, the color dialog needs this ability), because setting each property individually
' obviously works, but it also causes a lot of redundant redraws, which isn't great performance-wise.  This helper
' function seems like the least of many evils.
Public Sub SetGradientColorsAndValueAtOnce(ByVal leftGradientColor As OLE_COLOR, ByVal rightGradientColor As OLE_COLOR, ByVal newValue As Single)
    pdssPrimary.SetGradientColorsAndValueAtOnce leftGradientColor, rightGradientColor, newValue
End Sub

Private Sub pdssPrimary_Change()
    Me.Value = pdssPrimary.Value
End Sub

Private Sub pdssPrimary_FinalChange()
    RaiseEvent FinalChange
End Sub

Private Sub pdssPrimary_GotFocusAPI()
    EvaluateFocusCount
End Sub

Private Sub pdssPrimary_LostFocusAPI()
    EvaluateFocusCount
End Sub

Private Sub pdssPrimary_RenderTrackImage(dstDIB As pdDIB, ByVal leftBoundary As Single, ByVal rightBoundary As Single)
    RaiseEvent RenderTrackImage(dstDIB, leftBoundary, rightBoundary)
End Sub

'During owner-draw mode, our parent can call this sub if they need to modify their owner-drawn track image.
Public Sub RequestOwnerDrawChange()
    pdssPrimary.RequestOwnerDrawChange
End Sub

Private Sub pdssPrimary_SetCustomTabTarget(ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If shiftTabWasPressed Then RaiseEvent SetCustomTabTarget(True, newTargetHwnd)
End Sub

Private Sub tudPrimary_SetCustomTabTarget(ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If (Not shiftTabWasPressed) Then RaiseEvent SetCustomTabTarget(False, newTargetHwnd)
End Sub

Private Sub ucSupport_GotFocusAPI()
    EvaluateFocusCount
End Sub

Private Sub ucSupport_KeyDownSystem(ByVal Shift As ShiftConstants, ByVal whichSysKey As PD_NavigationKey, markEventHandled As Boolean)
    
    'Enter/Esc get reported directly to the system key handler.  Note that we track the return, because TRUE
    ' means the key was successfully forwarded to the relevant handler.  (If FALSE is returned, no control
    ' accepted the keypress, meaning we should forward the event down the line.)
    markEventHandled = NavKey.NotifyNavKeypress(Me, whichSysKey, Shift)
    
End Sub

Private Sub ucSupport_LostFocusAPI()
    EvaluateFocusCount
End Sub

Private Sub tudPrimary_Change()
    If (Not m_textBoxInitiated) Then
        If tudPrimary.IsValid(False) Then
            m_textBoxInitiated = True
            Me.Value = tudPrimary.Value
            m_textBoxInitiated = False
        End If
    End If
End Sub

Private Sub tudPrimary_FinalChange()
    RaiseEvent FinalChange
End Sub

Private Sub tudPrimary_GotFocusAPI()
    EvaluateFocusCount
End Sub

Private Sub tudPrimary_LostFocusAPI()
    EvaluateFocusCount
End Sub

Private Sub tudPrimary_ResetClick()
    RaiseEvent ResetClick
End Sub

Private Sub tudPrimary_Resize()
    UpdateControlLayout
End Sub

Private Sub ucSupport_RepaintRequired(ByVal updateLayoutToo As Boolean)
    If updateLayoutToo Then UpdateControlLayout Else ucSupport.RequestRepaint True
End Sub

Private Sub ucSupport_WindowResize(ByVal newWidth As Long, ByVal newHeight As Long)
    If (Not m_InternalResizeActive) Then UpdateControlLayout
End Sub

Private Sub UserControl_Initialize()
    
    'Initialize a user control support class
    Set ucSupport = New pdUCSupport
    ucSupport.RegisterControl UserControl.hWnd, False
    ucSupport.RequestExtraFunctionality True, , , False
    ucSupport.RequestCaptionSupport False
        
End Sub

'Initialize control properties for the first time
Private Sub UserControl_InitProperties()
    
    FontSizeTUD = 10
    FontSizeCaption = 12
    Caption = vbNullString
    CaptionPadding = 0
    
    Min = 0
    Max = 10
    SigDigits = 0
    Value = 0
    
    ScaleStyle = DefaultScaleLinear
    ScaleExponent = 2#
    SliderKnobStyle = DefaultKnobStyle
    SliderTrackStyle = DefaultTrackStyle
    GradientColorLeft = RGB(0, 0, 0)
    GradientColorRight = RGB(255, 255, 25)
    GradientColorMiddle = RGB(121, 131, 135)
    GradientMiddleValue = 0
    NotchPosition = AutomaticPosition
    NotchValueCustom = 0
    DefaultValue = NotchValueCustom
    
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    With PropBag
        Caption = .ReadProperty("Caption", vbNullString)
        CaptionPadding = .ReadProperty("CaptionPadding", 0)
        FontSizeCaption = .ReadProperty("FontSizeCaption", 12)
        FontSizeTUD = .ReadProperty("FontSizeTUD", 10)
        SigDigits = .ReadProperty("SigDigits", 0)
        Max = .ReadProperty("Max", 10)
        Min = .ReadProperty("Min", 0)
        ScaleStyle = .ReadProperty("ScaleStyle", DefaultScaleLinear)
        ScaleExponent = .ReadProperty("ScaleExponent", 2#)
        SliderKnobStyle = .ReadProperty("SliderKnobStyle", DefaultKnobStyle)
        SliderTrackStyle = .ReadProperty("SliderTrackStyle", DefaultTrackStyle)
        Value = .ReadProperty("Value", 0)
        GradientColorLeft = .ReadProperty("GradientColorLeft", RGB(0, 0, 0))
        GradientColorRight = .ReadProperty("GradientColorRight", RGB(255, 255, 255))
        GradientColorMiddle = .ReadProperty("GradientColorMiddle", RGB(121, 131, 135))
        GradientMiddleValue = .ReadProperty("GradientMiddleValue", 0)
        NotchPosition = .ReadProperty("NotchPosition", 0)
        NotchValueCustom = .ReadProperty("NotchValueCustom", 0)
        DefaultValue = .ReadProperty("DefaultValue", NotchValueCustom)
    End With
    
End Sub

'At run-time, painting is handled by PD's pdWindowPainter class.  In the IDE, however, we must rely on VB's internal paint event.
Private Sub UserControl_Paint()
    ucSupport.RequestIDERepaint UserControl.hDC
End Sub

Private Sub UserControl_Resize()
    If (Not PDMain.IsProgramRunning()) Then ucSupport.NotifyIDEResize UserControl.Width, UserControl.Height
End Sub

Private Sub UserControl_Show()
    UpdateControlLayout
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    With PropBag
        .WriteProperty "Caption", Me.Caption, vbNullString
        .WriteProperty "CaptionPadding", Me.CaptionPadding, 0
        .WriteProperty "FontSizeCaption", Me.FontSizeCaption, 12
        .WriteProperty "FontSizeTUD", Me.FontSizeTUD, 10
        .WriteProperty "Min", Me.Min, 0
        .WriteProperty "Max", Me.Max, 10
        .WriteProperty "SigDigits", Me.SigDigits, 0
        .WriteProperty "ScaleStyle", Me.ScaleStyle, DefaultScaleLinear
        .WriteProperty "ScaleExponent", Me.ScaleExponent, 2#
        .WriteProperty "SliderKnobStyle", Me.SliderKnobStyle, DefaultKnobStyle
        .WriteProperty "SliderTrackStyle", Me.SliderTrackStyle, DefaultTrackStyle
        .WriteProperty "Value", Me.Value, 0
        .WriteProperty "GradientColorLeft", Me.GradientColorLeft, RGB(0, 0, 0)
        .WriteProperty "GradientColorRight", Me.GradientColorRight, RGB(255, 255, 255)
        .WriteProperty "GradientColorMiddle", Me.GradientColorMiddle, RGB(121, 131, 135)
        .WriteProperty "GradientMiddleValue", Me.GradientMiddleValue, 0
        .WriteProperty "NotchPosition", Me.NotchPosition, 0
        .WriteProperty "NotchValueCustom", Me.NotchValueCustom, 0
        .WriteProperty "DefaultValue", Me.DefaultValue, Me.NotchValueCustom
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
    Dim newLeft_TUD As Long, newTop_TUD As Long, newTop_Slider As Long, newWidth_Slider As Long
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
        newControlHeight = tudPrimary.GetHeight + Interface.FixDPI(4) + textHeight + Interface.FixDPI(2) + m_captionPadding
        
        'Calculate a new top position for the slider box (which will be vertically centered in the space below the caption)
        newTop_Slider = ((newControlHeight - (textHeight + Interface.FixDPI(4))) - tudPrimary.GetHeight) \ 2
        newTop_Slider = textHeight + Interface.FixDPI(4) + newTop_Slider + m_captionPadding
        
    'When a slider lacks a caption, we hard-code its height to preset values
    Else
        
        'Start by setting the control height
        newControlHeight = tudPrimary.GetHeight + Interface.FixDPI(4)
        
        'Center the slider box inside the newly calculated height
        newTop_Slider = (newControlHeight - pdssPrimary.GetHeight) \ 2
        
    End If
    
    'Apply the new height to this UC instance
    If (ucSupport.GetControlHeight <> newControlHeight) Then ucSupport.RequestNewSize , newControlHeight
    
    'If the text up/down control has a huge upper limit, increase its width (to ensure all digits are visible)
    If (tudPrimary.Max >= 100000) Then tudPrimary.SetWidth Interface.FixDPI(92) Else tudPrimary.SetWidth Interface.FixDPI(84)
    
    'With height correctly set, we next want to left-align the spinner against the slider region
    newLeft_TUD = ucSupport.GetControlWidth - (tudPrimary.GetWidth + Interface.FixDPI(2))
    
    'Because the slider width is contingent on the spinner position, calculate it next, then move it into place
    newWidth_Slider = newLeft_TUD - Interface.FixDPI(6)
    If (newTop_Slider <> pdssPrimary.GetTop) Then pdssPrimary.SetTop newTop_Slider
    If (newWidth_Slider > 0) And (newWidth_Slider <> pdssPrimary.GetWidth) Then pdssPrimary.SetWidth newWidth_Slider
    
    'Vertically center the spinner relative to the slider
    Dim sliderVerticalCenter As Single
    sliderVerticalCenter = pdssPrimary.GetTop + (CSng(pdssPrimary.GetHeight) / 2)
    newTop_TUD = sliderVerticalCenter - Int(CSng(tudPrimary.GetHeight) / 2) + 1
    
    'Now that we've calculated new text up/down positioning, we can apply it as necessary
    If (tudPrimary.GetTop <> newTop_TUD) Then tudPrimary.SetTop newTop_TUD
    If (tudPrimary.GetLeft <> newLeft_TUD) Then tudPrimary.SetLeft newLeft_TUD
    
    'Failsafe check for text up/down visibility
    If (ucSupport.GetControlHeight < tudPrimary.GetTop + tudPrimary.GetHeight) Then ucSupport.RequestNewSize , tudPrimary.GetTop + tudPrimary.GetHeight
    
    'Inside the IDE, use a line of dummy code to force a redraw of the control outline
    If (Not PDMain.IsProgramRunning()) Then
        pdssPrimary.Visible = False
        Dim bufferDC As Long
        bufferDC = ucSupport.GetBackBufferDC(True)
    End If
    
    ucSupport.RequestRepaint True
    m_InternalResizeActive = False
    
End Sub

'After a component of this control gets or loses focus, it needs to call this function.  This function is responsible for raising
' Got/LostFocusAPI events, which are important as an API text box is part of this control.
Private Sub EvaluateFocusCount()

    If (Not m_LastFocusState) And Me.HasFocus() Then
        m_LastFocusState = True
        RaiseEvent GotFocusAPI
    ElseIf m_LastFocusState And (Not Me.HasFocus()) Then
        m_LastFocusState = False
        RaiseEvent LostFocusAPI
    End If

End Sub

'External functions can call this to request a redraw.  This is helpful for live-updating theme settings, as in the Preferences dialog.
Public Sub UpdateAgainstCurrentTheme(Optional ByVal hostFormhWnd As Long = 0)
    If ucSupport.ThemeUpdateRequired Then
        pdssPrimary.UpdateAgainstCurrentTheme
        tudPrimary.UpdateAgainstCurrentTheme
        If PDMain.IsProgramRunning() Then NavKey.NotifyControlLoad Me, hostFormhWnd, False
        If PDMain.IsProgramRunning() Then ucSupport.UpdateAgainstThemeAndLanguage
    End If
End Sub

'Due to complex interactions between user controls and PD's translation engine, tooltips require this dedicated function.
' (IMPORTANT NOTE: the tooltip class will handle translations automatically.  Always pass the original English text!)
Public Sub AssignTooltip(ByRef newTooltip As String, Optional ByRef newTooltipTitle As String = vbNullString, Optional ByVal raiseTipsImmediately As Boolean = False)
    pdssPrimary.AssignTooltip newTooltip, newTooltipTitle
    tudPrimary.AssignTooltip newTooltip, newTooltipTitle
End Sub
