VERSION 5.00
Begin VB.UserControl textUpDown 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BackStyle       =   0  'Transparent
   ClientHeight    =   420
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1125
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   28
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   75
   ToolboxBitmap   =   "textUpDown.ctx":0000
   Begin VB.Timer tmrDownButton 
      Enabled         =   0   'False
      Left            =   1080
      Top             =   0
   End
   Begin VB.Timer tmrUpButton 
      Enabled         =   0   'False
      Left            =   1080
      Top             =   120
   End
   Begin PhotoDemon.pdTextBox txtPrimary 
      Height          =   315
      Left            =   15
      TabIndex        =   1
      Top             =   15
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   556
      TabBehavior     =   1
   End
   Begin VB.PictureBox picScroll 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   720
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   17
      TabIndex        =   0
      Top             =   0
      Width           =   255
   End
   Begin VB.Shape shpError 
      BorderColor     =   &H000000FF&
      Height          =   390
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1005
   End
End
Attribute VB_Name = "textUpDown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Text / UpDown custom control
'Copyright 2013-2015 by Tanner Helland
'Created: 19/April/13
'Last updated: 06/January/15
'Last update: replace scroll bar with custom buttons that behave like a scroll bar
'
'Software like PhotoDemon requires a lot of controls.  Ideally, every setting should be adjustable by at least
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
'This control handles the following things automatically:
' 1) Synching of text and scroll/slide values
' 2) Validation of text entries, including a function for external validation requests
' 3) Locale handling (like the aforementioned comma/decimal replacement in some countries)
' 4) A single "Change" event that fires for either scroll or text changes, and only if a text change is valid
' 5) Support for floating-point values (scroll bar max/min values are automatically adjusted to mimic this)
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'This object can raise a Change (which triggers when the Value property is changed by ANY means)
Public Event Change()

'Because we have multiple components on this user control, including an API text box, we report our own Got/Lost focus events.
Public Event GotFocusAPI()
Public Event LostFocusAPI()

'Reliable focus detection on the spin control requires a specialized subclasser
Private WithEvents cFocusDetector As pdFocusDetector
Attribute cFocusDetector.VB_VarHelpID = -1

'For performance reasons, this object can always raise an event I call "FinalChange".  This triggers under the same conditions as Change,
' *EXCEPT* when the mouse button is held down.  FinalChange will not fire until the mouse button is released, which makes it ideal
' for things like syncing time-consuming UI elements.
Public Event FinalChange()

'The only exposed font setting is size.  All other settings are handled automatically, by the themer.
Private m_FontSize As Single

'Additional helper for rendering themed and multiline tooltips
Private toolTipManager As pdToolTip

'Used to track value, min, and max values as floating-points
Private controlVal As Double, controlMin As Double, controlMax As Double

'The number of significant digits for this control.  0 means integer values.
Private significantDigits As Long

'If the text box is initiating a value change, we must track that so as to not overwrite the user's entry mid-typing
Private textBoxInitiated As Boolean

'To simplify mouse_down handling, size events fill two rects: one for the "up" spin button, and another for the "down" spin button.
' These are relative to the picScroll object - not the underlying usercontrol!  (This is necessary due to the way VB handles focus
' for user controls with child objects on them.)
Private upRect As RECT, downRect As RECT

'Flicker-free painter for the spin button area
Private WithEvents cPainter As pdWindowPainter
Attribute cPainter.VB_VarHelpID = -1

'All spin button painting is performed on this DIB
Private buttonDIB As pdDIB

'Mouse handler for the spin button area
Private WithEvents cMouseEvents As pdInputMouse
Attribute cMouseEvents.VB_VarHelpID = -1

'Mouse state for the spin button area
Private m_MouseDownUpButton As Boolean, m_MouseDownDownButton As Boolean
Private m_MouseOverUpButton As Boolean, m_MouseOverDownButton As Boolean

'Tracks whether the control (any component) has focus.  This is helpful as we must synchronize between VB's focus events and API
' focus events.  Every time an individual component gains focus, we increment this counter by 1.  Every time an individual component
' loses focus, we decrement the counter by 1.  When the counter hits 0, we report a control-wide Got/LostFocusAPI event.
Private m_ControlFocusCount As Long

Private Declare Function IntersectRect Lib "user32" (ByRef lpDestRect As RECTL, ByRef lpSrc1Rect As RECTL, ByRef lpSrc2Rect As RECTL) As Long

'If the current text value is NOT valid, this will return FALSE
Public Property Get IsValid(Optional ByVal showError As Boolean = True) As Boolean
    
    Dim retVal As Boolean
    retVal = Not shpError.Visible
    
    'If the current text value is not valid, highlight the problem and optionally display an error message box
    If Not retVal Then
        txtPrimary.selectAll
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
    
    'Mirror the new enabled setting across child controls
    UserControl.Enabled = newValue
    txtPrimary.Enabled = newValue
    txtPrimary = getFormattedStringValue(controlVal)
    
    'Request a button redraw
    RedrawButton
    
    PropertyChanged "Enabled"
    
End Property

Public Property Get FontSize() As Single
Attribute FontSize.VB_ProcData.VB_Invoke_Property = "StandardFont;Font"
Attribute FontSize.VB_UserMemId = -512
    FontSize = m_FontSize
End Property

Public Property Let FontSize(ByVal newSize As Single)
    If m_FontSize <> newSize Then
        m_FontSize = newSize
        txtPrimary.FontSize = m_FontSize
        PropertyChanged "FontSize"
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

Private Sub cMouseEvents_MouseDownCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    
    'Determine mouse button state for the up and down button areas
    If (Button = pdLeftButton) And Me.Enabled Then
    
        If isPointInRect(x, y, upRect) Then
            m_MouseDownUpButton = True
            
            'Adjust the value immediately
            moveValueDown
            
            'Start the repeat timer as well
            tmrUpButton.Interval = Interface.GetKeyboardDelay() * 1000
            tmrUpButton.Enabled = True
            
        Else
            m_MouseDownUpButton = False
        End If
        
        If isPointInRect(x, y, downRect) Then
            m_MouseDownDownButton = True
            moveValueUp
            tmrDownButton.Interval = Interface.GetKeyboardDelay() * 1000
            tmrDownButton.Enabled = True
        Else
            m_MouseDownDownButton = False
        End If
        
        'Request a button redraw
        RedrawButton
        
    End If
    
End Sub

Private Sub cMouseEvents_MouseEnter(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    cMouseEvents.setSystemCursor IDC_HAND
End Sub

Private Sub cMouseEvents_MouseLeave(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    
    cMouseEvents.setSystemCursor IDC_DEFAULT
    
    m_MouseOverUpButton = False
    m_MouseOverDownButton = False
    
    'Request a button redraw
    RedrawButton
    
End Sub

Private Sub cMouseEvents_MouseMoveCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    
    'Determine mouse hover state for the up and down button areas
    If isPointInRect(x, y, upRect) Then
        m_MouseOverUpButton = True
    Else
        m_MouseOverUpButton = False
    End If
    
    If isPointInRect(x, y, downRect) Then
        m_MouseOverDownButton = True
    Else
        m_MouseOverDownButton = False
    End If
    
    'Request a button redraw
    RedrawButton
    
End Sub

'Reset spin control button state on a mouse up event
Private Sub cMouseEvents_MouseUpCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal ClickEventAlsoFiring As Boolean)
    
    If Button = pdLeftButton Then
        
        m_MouseDownUpButton = False
        m_MouseDownDownButton = False
        tmrUpButton.Enabled = False
        tmrDownButton.Enabled = False
        
        'When the mouse is release, raise a "FinalChange" event, which lets the caller know that they can perform any
        ' long-running actions now.
        RaiseEvent FinalChange
        
        'Request a button redraw
        RedrawButton
        
    End If
        
End Sub

Private Sub cPainter_PaintWindow(ByVal winLeft As Long, ByVal winTop As Long, ByVal winWidth As Long, ByVal winHeight As Long)
    
    'Flip the button back buffer to the screen
    If Not (buttonDIB Is Nothing) Then
        BitBlt cPainter.getPaintStructDC, 0, 0, buttonDIB.getDIBWidth, buttonDIB.getDIBHeight, buttonDIB.getDIBDC, 0, 0, vbSrcCopy
    End If
    
End Sub

Private Sub tmrDownButton_Timer()

    'If this is the first time the button is firing, we want to reset the button's interval to the repeat rate instead
    ' of the delay rate.
    If tmrDownButton.Interval = Interface.GetKeyboardDelay * 1000 Then
        tmrDownButton.Interval = Interface.GetKeyboardRepeatRate * 1000
    End If
    
    'It's a little counter-intuitive, but the DOWN button actually moves the control value UP
    moveValueUp

End Sub

Private Sub tmrUpButton_Timer()
    
    'If this is the first time the button is firing, we want to reset the button's interval to the repeat rate instead
    ' of the delay rate.
    If tmrUpButton.Interval = Interface.GetKeyboardDelay * 1000 Then
        tmrUpButton.Interval = Interface.GetKeyboardRepeatRate * 1000
    End If
    
    'It's a little counter-intuitive, but the UP button actually moves the control value DOWN
    moveValueDown
    
End Sub

'When the control value is moved UP via button, this function is called
Private Sub moveValueUp()
    Value = controlVal - (1 / (10 ^ significantDigits))
End Sub

'When the control value is moved DOWN via button, this function is called
Private Sub moveValueDown()
    Value = controlVal + (1 / (10 ^ significantDigits))
End Sub

Private Sub txtPrimary_Change()
    
    If IsTextEntryValid() Then
        If shpError.Visible Then shpError.Visible = False
        textBoxInitiated = True
        Value = CDblCustom(txtPrimary)
        textBoxInitiated = False
    Else
        If Me.Enabled Then shpError.Visible = True
    End If
    
End Sub

Private Sub txtPrimary_GotFocusAPI()
    
    m_ControlFocusCount = m_ControlFocusCount + 1
    evaluateFocusCount True
    
    'As a convenience to the user, select all text when first clicked
    txtPrimary.selectAll
    
End Sub

Private Sub txtPrimary_LostFocusAPI()
    m_ControlFocusCount = m_ControlFocusCount - 1
    evaluateFocusCount False
End Sub

Private Sub txtPrimary_Resize()
    
    If UserControl.ScaleHeight <> txtPrimary.Height + 2 Then
        UserControl.Extender.Height = txtPrimary.Height + 2
    End If
    
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
    If (newValue <> controlVal) Or (Not IsValid(False)) Then
        
        controlVal = newValue
                
        'While running, perform bounds-checking.  (It's less important in the designer, as the assumption is that the
        ' developer will momentarily bring everything into order.)
        If g_IsProgramRunning Then
                
            'To prevent RTEs, perform an additional bounds check.  Don't assign the value if it's invalid.
            If controlVal < controlMin Then
                'Debug.Print "Control value forcibly changed to bring it in-bounds (too low)"
                controlVal = controlMin
            End If
            
            If controlVal > controlMax Then
                'Debug.Print "Control value forcibly changed to bring it in-bounds (too high)"
                controlVal = controlMax
            End If
            
        End If
                
        'Mirror the value to the text box
        If Not textBoxInitiated Then
            If (Not IsValid(False)) Then
                txtPrimary = getFormattedStringValue(controlVal)
                shpError.Visible = False
            Else
                If Len(txtPrimary) > 0 Then
                    If StrComp(getFormattedStringValue(txtPrimary), CStr(controlVal), vbBinaryCompare) <> 0 Then txtPrimary.Text = getFormattedStringValue(controlVal)
                End If
            End If
            
        End If
    
        'Mark the value property as being changed, and raise the corresponding event.
        PropertyChanged "Value"
        RaiseEvent Change
        
        'If the mouse button is *not* currently down, raise the "FinalChange" event too
        If (Not m_MouseDownUpButton) And (Not m_MouseDownDownButton) Then RaiseEvent FinalChange
        
    End If
                
End Property

'Note: the control's minimum value is settable at run-time
Public Property Get Min() As Double
    Min = controlMin
End Property

Public Property Let Min(ByVal newValue As Double)
        
    controlMin = newValue
    
    'If the current control .Value is less than the new minimum, change it to match
    If controlVal < controlMin Then
        controlVal = controlMin
        txtPrimary = CStr(controlVal)
        RaiseEvent Change
    End If
    
    PropertyChanged "Min"
    
End Property

'Note: the control's maximum value is settable at run-time
Public Property Get Max() As Double
    Max = controlMax
End Property

Public Property Let Max(ByVal newValue As Double)
        
    controlMax = newValue
    
    'If the current control .Value is greater than the new max, change it to match
    If controlVal > controlMax Then
        controlVal = controlMax
        txtPrimary = CStr(controlVal)
        RaiseEvent Change
    End If
    
    PropertyChanged "Max"
    
End Property

'Significant digits determines whether the control allows float values or int values (and with how much precision)
Public Property Get SigDigits() As Long
    SigDigits = significantDigits
End Property

Public Property Let SigDigits(ByVal newValue As Long)
        
    significantDigits = newValue
        
    'Update the text display to reflect the new significant digit amount, including any decimal places
    txtPrimary = getFormattedStringValue(controlVal)
    
    PropertyChanged "SigDigits"
    
End Property

'Mirror the code from the change event, but force a formatted text sync
Private Sub txtPrimary_Validate(Cancel As Boolean)
    If IsTextEntryValid() Then
        If shpError.Visible Then shpError.Visible = False
        Value = CDblCustom(txtPrimary)
    Else
        If Me.Enabled Then shpError.Visible = True
    End If
End Sub

Private Sub UserControl_GotFocus()
    m_ControlFocusCount = m_ControlFocusCount + 1
    evaluateFocusCount True
End Sub

Private Sub UserControl_Initialize()
        
    'Prepare a default font size
    m_FontSize = 10
    txtPrimary.FontSize = m_FontSize
        
    'Prep the spin button back buffer
    Set buttonDIB = New pdDIB
    If g_IsProgramRunning Then buttonDIB.createBlank picScroll.ScaleWidth, picScroll.ScaleHeight, 24
    
    'Prepare a window painter for the spin button area
    Set cPainter = New pdWindowPainter
    If g_IsProgramRunning Then cPainter.startPainter picScroll.hWnd
    
    'Prepare an input handler for the spin button area
    Set cMouseEvents = New pdInputMouse
    If g_IsProgramRunning Then cMouseEvents.addInputTracker picScroll.hWnd, True, True, False, True, False
    
    'Also start a focus detector for the spinner picture box
    Set cFocusDetector = New pdFocusDetector
    If g_IsProgramRunning Then cFocusDetector.startFocusTracking picScroll.hWnd
    
    'Reset the focus count
    m_ControlFocusCount = 0
    
    'Create a tooltip engine
    Set toolTipManager = New pdToolTip
                    
End Sub

Private Sub UserControl_InitProperties()
    
    FontSize = 10
    m_FontSize = 10
        
    Value = 0
    controlVal = 0
    
    Min = 0
    controlMin = 0
    
    Max = 10
    controlMax = 10
    
    SigDigits = 0
    significantDigits = 0
    
End Sub

Private Sub UserControl_LostFocus()
    m_ControlFocusCount = m_ControlFocusCount - 1
    evaluateFocusCount False
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    With PropBag
        FontSize = .ReadProperty("FontSize", 10)
        ForeColor = .ReadProperty("ForeColor", &H404040)
        SigDigits = .ReadProperty("SigDigits", 0)
        Max = .ReadProperty("Max", 10)
        Min = .ReadProperty("Min", 0)
        Value = .ReadProperty("Value", 0)
    End With
        
End Sub

Private Sub UserControl_Resize()
    resizeControl
End Sub

Private Sub resizeControl()

    'The goal here is to keep the text box and scroll bar nicely aligned, with a 1px border for the red "error" box
    picScroll.Width = FixDPI(18)
    picScroll.Top = 1
    picScroll.Height = txtPrimary.Height
    
    'Leave a 1px border around the text box, to be used for displaying red during range and numeric errors
    txtPrimary.Left = 1
    txtPrimary.Top = 1
    txtPrimary.Width = UserControl.ScaleWidth - 2 - picScroll.Width
    
    'Align the scroll bar container to the right of the text box
    picScroll.Left = txtPrimary.Left + txtPrimary.Width
    
    'Calculate new rects for the up/down buttons
    With upRect
        .Left = 0
        .Right = picScroll.ScaleWidth - 1
        .Top = 0
        .Bottom = (picScroll.ScaleHeight \ 2) - 1
    End With
    
    With downRect
        .Left = 0
        .Right = picScroll.ScaleWidth - 1
        .Top = upRect.Bottom + 1
        .Bottom = picScroll.ScaleHeight - 1
    End With
    
    'Make the shape control (used for errors) the size of the user control
    shpError.Left = 0
    shpError.Top = 0
    shpError.Height = UserControl.ScaleHeight
    shpError.Width = UserControl.ScaleWidth
    
    'Resize the button back buffer to match
    If Not (buttonDIB Is Nothing) Then
        If (buttonDIB.getDIBWidth <> picScroll.ScaleWidth) Or (buttonDIB.getDIBHeight <> picScroll.ScaleHeight) Then
            buttonDIB.createBlank picScroll.ScaleWidth, picScroll.ScaleHeight, 24
        End If
    End If
    
    'Request a redraw of the button
    RedrawButton
    
End Sub

'Redraw the spin button area of the control
Private Sub RedrawButton()
    
    'Start by determining what color to use for the background.  (In the IDE, we have to supply all colors manually.)
    Dim buttonBackColor As Long, buttonBorderColor As Long
    
    Dim upButtonBorderColor As Long, downButtonBorderColor As Long
    Dim upButtonFillColor As Long, downButtonFillColor As Long
    Dim upButtonArrowColor As Long, downButtonArrowColor As Long
    
    If Not (g_Themer Is Nothing) Then
        buttonBackColor = g_Themer.getThemeColor(PDTC_BACKGROUND_DEFAULT)
        buttonBorderColor = g_Themer.getThemeColor(PDTC_GRAY_DEFAULT)
    Else
        buttonBackColor = vbWindowBackground
        buttonBorderColor = RGB(128, 128, 128)
    End If
    
    'Start by erasing the buffer (which will have already been sized correctly by a previous function) and drawing
    ' a default border around the entire control.
    GDI_Plus.GDIPlusFillDIBRect buttonDIB, 0, 0, buttonDIB.getDIBWidth, buttonDIB.getDIBHeight, buttonBackColor
    
    'Next, figure out button colors.  These are affected by hover and press state.
    If m_MouseOverUpButton And Me.Enabled And (Not (g_Themer Is Nothing)) Then
    
        If m_MouseDownUpButton Then
            upButtonBorderColor = g_Themer.getThemeColor(PDTC_ACCENT_DEFAULT)
            upButtonArrowColor = g_Themer.getThemeColor(PDTC_TEXT_INVERT)
            upButtonFillColor = g_Themer.getThemeColor(PDTC_ACCENT_DEFAULT)
        Else
            upButtonBorderColor = g_Themer.getThemeColor(PDTC_ACCENT_SHADOW)
            upButtonArrowColor = g_Themer.getThemeColor(PDTC_ACCENT_DEFAULT)
            upButtonFillColor = g_Themer.getThemeColor(PDTC_BACKGROUND_DEFAULT)
        End If
    
    Else
        If Not (g_Themer Is Nothing) Then
            upButtonBorderColor = g_Themer.getThemeColor(PDTC_BACKGROUND_DEFAULT)
            If Me.Enabled Then upButtonArrowColor = g_Themer.getThemeColor(PDTC_GRAY_DEFAULT) Else upButtonArrowColor = g_Themer.getThemeColor(PDTC_GRAY_HIGHLIGHT)
            upButtonFillColor = g_Themer.getThemeColor(PDTC_BACKGROUND_DEFAULT)
        Else
            upButtonBorderColor = vbWindowBackground
            upButtonArrowColor = RGB(128, 128, 128)
            upButtonFillColor = vbWindowBackground
        End If
    End If
    
    If m_MouseOverDownButton And Me.Enabled And (Not (g_Themer Is Nothing)) Then
    
        If m_MouseDownDownButton Then
            downButtonBorderColor = g_Themer.getThemeColor(PDTC_ACCENT_DEFAULT)
            downButtonArrowColor = g_Themer.getThemeColor(PDTC_TEXT_INVERT)
            downButtonFillColor = g_Themer.getThemeColor(PDTC_ACCENT_DEFAULT)
        Else
            downButtonBorderColor = g_Themer.getThemeColor(PDTC_ACCENT_SHADOW)
            downButtonArrowColor = g_Themer.getThemeColor(PDTC_ACCENT_DEFAULT)
            downButtonFillColor = g_Themer.getThemeColor(PDTC_BACKGROUND_DEFAULT)
        End If
    
    Else
        If Not (g_Themer Is Nothing) Then
            downButtonBorderColor = g_Themer.getThemeColor(PDTC_BACKGROUND_DEFAULT)
            If Me.Enabled Then downButtonArrowColor = g_Themer.getThemeColor(PDTC_GRAY_DEFAULT) Else downButtonArrowColor = g_Themer.getThemeColor(PDTC_GRAY_HIGHLIGHT)
            downButtonFillColor = g_Themer.getThemeColor(PDTC_BACKGROUND_DEFAULT)
        Else
            downButtonBorderColor = vbWindowBackground
            downButtonArrowColor = RGB(128, 128, 128)
            downButtonFillColor = vbWindowBackground
        End If
    End If
    
    'Paint both button backgrounds and borders
    GDI_Plus.GDIPlusFillDIBRect buttonDIB, upRect.Left, upRect.Top, upRect.Right - upRect.Left, upRect.Bottom - upRect.Top, upButtonFillColor
    GDI_Plus.GDIPlusDrawRectOutlineToDC buttonDIB.getDIBDC, upRect.Left, upRect.Top, upRect.Right, upRect.Bottom, upButtonBorderColor
    
    GDI_Plus.GDIPlusFillDIBRect buttonDIB, downRect.Left, downRect.Top, downRect.Right - downRect.Left, downRect.Bottom - downRect.Top, downButtonFillColor
    GDI_Plus.GDIPlusDrawRectOutlineToDC buttonDIB.getDIBDC, downRect.Left, downRect.Top, downRect.Right, downRect.Bottom, downButtonBorderColor
    
    'Finally, paint the arrows themselves
    Dim buttonPt1 As POINTFLOAT, buttonPt2 As POINTFLOAT, buttonPt3 As POINTFLOAT
                
    'Start with the up-pointing arrow
    buttonPt1.x = upRect.Left + FixDPIFloat(5)
    buttonPt1.y = (upRect.Bottom - upRect.Top) / 2 + FixDPIFloat(2)
    
    buttonPt3.x = upRect.Right - FixDPIFloat(5)
    buttonPt3.y = buttonPt1.y
    
    buttonPt2.x = buttonPt1.x + (buttonPt3.x - buttonPt1.x) / 2
    buttonPt2.y = buttonPt1.y - FixDPIFloat(3)
    
    GDI_Plus.GDIPlusDrawLineToDC buttonDIB.getDIBDC, buttonPt1.x, buttonPt1.y, buttonPt2.x, buttonPt2.y, upButtonArrowColor, 255, 2, True, LineCapRound
    GDI_Plus.GDIPlusDrawLineToDC buttonDIB.getDIBDC, buttonPt2.x, buttonPt2.y, buttonPt3.x, buttonPt3.y, upButtonArrowColor, 255, 2, True, LineCapRound
                
    'Next, the down-pointing arrow
    buttonPt1.x = downRect.Left + FixDPIFloat(5)
    buttonPt1.y = downRect.Top + (downRect.Bottom - downRect.Top) / 2 - FixDPIFloat(1)
    
    buttonPt3.x = downRect.Right - FixDPIFloat(5)
    buttonPt3.y = buttonPt1.y
    
    buttonPt2.x = buttonPt1.x + (buttonPt3.x - buttonPt1.x) / 2
    buttonPt2.y = buttonPt1.y + FixDPIFloat(3)
    
    GDI_Plus.GDIPlusDrawLineToDC buttonDIB.getDIBDC, buttonPt1.x, buttonPt1.y, buttonPt2.x, buttonPt2.y, downButtonArrowColor, 255, 2, True, LineCapRound
    GDI_Plus.GDIPlusDrawLineToDC buttonDIB.getDIBDC, buttonPt2.x, buttonPt2.y, buttonPt3.x, buttonPt3.y, downButtonArrowColor, 255, 2, True, LineCapRound
    
    'As a final step, request a repaint onto the button's container
    cPainter.requestRepaint

End Sub

Private Sub UserControl_Show()
        
    'Also, force a resize to modify its layout
    If g_IsProgramRunning Then UserControl_Resize
        
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    'Store all associated properties
    With PropBag
        .WriteProperty "Min", controlMin, 0
        .WriteProperty "Max", controlMax, 10
        .WriteProperty "SigDigits", significantDigits, 0
        .WriteProperty "Value", controlVal, 0
        .WriteProperty "FontSize", m_FontSize, 10
        .WriteProperty "ForeColor", ForeColor, &H404040
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
    
    'Perform a final check for control enablement.  If the control is disabled, we do not (currently) display anything.
    If Not Me.Enabled Then getFormattedStringValue = ""

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
        
    If Not IsNumeric(chkString) Then
        If displayErrorMsg Then PDMsgBox "%1 is not a valid entry." & vbCrLf & "Please enter a numeric value.", vbExclamation + vbOKOnly + vbApplicationModal, "Invalid entry", txtPrimary
        IsTextEntryValid = False
    Else
        
        Dim checkVal As Double
        checkVal = CDblCustom(chkString)
    
        If (checkVal >= controlMin) And (checkVal <= controlMax) Then
            IsTextEntryValid = True
        Else
            If displayErrorMsg Then PDMsgBox "%1 is not a valid entry." & vbCrLf & "Please enter a value between %2 and %3.", vbExclamation + vbOKOnly + vbApplicationModal, "Invalid entry", txtPrimary, getFormattedStringValue(controlMin), getFormattedStringValue(controlMax)
            IsTextEntryValid = False
        End If
    End If
    
End Function

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
    
    'Text boxes handle their own updating
    If g_IsProgramRunning Then txtPrimary.UpdateAgainstCurrentTheme
    
    'Our tooltip object must also be refreshed (in case the language has changed)
    If g_IsProgramRunning Then toolTipManager.UpdateAgainstCurrentTheme
    
    'Request a repaint
    If Not cPainter Is Nothing Then cPainter.requestRepaint
    
End Sub

'Due to complex interactions between user controls and PD's translation engine, tooltips require this dedicated function.
' (IMPORTANT NOTE: the tooltip class will handle translations automatically.  Always pass the original English text!)
Public Sub AssignTooltip(ByVal newTooltip As String, Optional ByVal newTooltipTitle As String, Optional ByVal newTooltipIcon As TT_ICON_TYPE = TTI_NONE)
    toolTipManager.setTooltip Me.hWnd, UserControl.containerHwnd, newTooltip, newTooltipTitle, newTooltipIcon
    toolTipManager.setTooltip picScroll.hWnd, UserControl.containerHwnd, newTooltip, newTooltipTitle, newTooltipIcon
End Sub
