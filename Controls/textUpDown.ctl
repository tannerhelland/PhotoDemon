VERSION 5.00
Begin VB.UserControl textUpDown 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BackStyle       =   0  'Transparent
   ClientHeight    =   420
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1125
   ClipControls    =   0   'False
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
   Begin VB.VScrollBar vsPrimary 
      Height          =   360
      Left            =   720
      TabIndex        =   1
      Top             =   15
      Width           =   255
   End
   Begin VB.TextBox txtPrimary 
      Alignment       =   2  'Center
      CausesValidation=   0   'False
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
      Left            =   15
      TabIndex        =   0
      Text            =   "0"
      Top             =   15
      Width           =   735
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
'Copyright ©2013-2014 by Tanner Helland
'Created: 19/April/13
'Last updated: 13/September/13
'Last update: fix non-96dpi layout issues
'
'Software like PhotoDemon requires a lot of controls.  Ideally, every setting should be adjustable by at least
' two mechanisms: direct text entry, and some kind of slider or scroll bar, which allows for a quick method to
' make both large and small adjustments to a given parameter.
'
'Historically, I accomplished this by providing a scroll bar and text box for every parameter in the program.
' This got the job done, but it had a number of limitations - such as requiring an enormous amount of time if
' changes ever needed to be made, and custom code being required in every form to handle text / scroll synching.
'
'In April 2014, it was brought to my attention that some locales (e.g. Italy) use a comma instead of a decimal
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

Private WithEvents mFont As StdFont
Attribute mFont.VB_VarHelpID = -1

Private origForecolor As Long

'Used to render themed and multiline tooltips
Private m_ToolTip As clsToolTip
Private m_ToolString As String

'Used to track value, min, and max values as floating-points
Private controlVal As Double, controlMin As Double, controlMax As Double

'The number of significant digits for this control.  0 means integer values.
Private significantDigits As Long

'If the text box is initiating a value change, we must track that so as to not overwrite the user's entry mid-typing
Private textBoxInitiated As Boolean

'Used to prevent the obnoxious blinking effect of the vertical scroll bar
Private Declare Function DestroyCaret Lib "user32" () As Long

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

Public Property Let Enabled(ByVal NewValue As Boolean)
    UserControl.Enabled = NewValue
    vsPrimary.Enabled = -NewValue
    txtPrimary.Enabled = NewValue
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

Private Sub vsPrimary_Change()
    If Not textBoxInitiated Then copyValToTextBox -1 * vsPrimary.Value
    Value = -1 * (vsPrimary / (10 ^ significantDigits))
End Sub

Private Sub vsPrimary_Scroll()
    copyValToTextBox -1 * vsPrimary.Value
    Value = -1 * (vsPrimary / (10 ^ significantDigits))
End Sub

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

Public Property Let Value(ByVal NewValue As Double)
    
    'Don't make any changes unless the new value deviates from the existing one
    If NewValue <> controlVal Then
    
        'Internally track the value of the control
        controlVal = NewValue
        
        'Assign the scroll bar the "same" value.  This will vary based on the number of significant digits in use; because
        ' scroll bars cannot hold float values, we have to multiple by 10^n where n is the number of significant digits
        ' in use for this control.
        Dim newScrollVal As Long
        newScrollVal = -1 * CLng(controlVal * (10 ^ significantDigits))
        
        If vsPrimary <> newScrollVal Then
            
            'To prevent RTEs, perform an additional bounds check.  Don't assign the value if it's invalid.
            If newScrollVal >= vsPrimary.Min And newScrollVal <= vsPrimary.Max Then vsPrimary = newScrollVal
            
        End If
        
        'Mirror the value to the text box
        If Not textBoxInitiated Then
            If StrComp(getFormattedStringValue(txtPrimary), CStr(controlVal), vbBinaryCompare) <> 0 Then txtPrimary.Text = getFormattedStringValue(controlVal)
        End If
        
        'Mark the value property as being changed, and raise the corresponding event.
        PropertyChanged "Value"
        RaiseEvent Change
        
    End If
    
End Property

'Note: the control's minimum value is settable at run-time
Public Property Get Min() As Double
    Min = controlMin
End Property

Public Property Let Min(ByVal NewValue As Double)
    
    controlMin = NewValue
    vsPrimary.Max = -1 * controlMin * (10 ^ significantDigits)
    
    'If the current control .Value is less than the new minimum, change it to match
    If controlVal < controlMin Then
        controlVal = controlMin
        vsPrimary = -1 * controlVal * (10 ^ significantDigits)
        txtPrimary = CStr(controlVal)
        RaiseEvent Change
    End If
    
    PropertyChanged "Min"
    
End Property

'Note: the control's maximum value is settable at run-time
Public Property Get Max() As Double
    Max = controlMax
End Property

Public Property Let Max(ByVal NewValue As Double)
    
    controlMax = NewValue
    vsPrimary.Min = -1 * controlMax * (10 ^ significantDigits)
    
    'If the current control .Value is greater than the new max, change it to match
    If controlVal > controlMax Then
        controlVal = controlMax
        vsPrimary = -1 * controlVal * (10 ^ significantDigits)
        txtPrimary = CStr(controlVal)
        RaiseEvent Change
    End If
    
    PropertyChanged "Max"
    
End Property

'Significant digits determines whether the control allows float values or int values (and with how much precision)
Public Property Get SigDigits() As Long
    SigDigits = significantDigits
End Property

Public Property Let SigDigits(ByVal NewValue As Long)
    
    significantDigits = NewValue
    
    'Update the scroll bar's min and max values accordingly
    vsPrimary.Max = -1 * controlMin * (10 ^ significantDigits)
    vsPrimary.Min = -1 * controlMax * (10 ^ significantDigits)
    
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
        vsPrimary.Value = Val(txtPrimary) * -1 * (10 ^ significantDigits)
        textBoxInitiated = False
    Else
        shpError.Visible = True
    End If
    
End Sub

Private Sub UserControl_Initialize()
    
    'Apply a hand cursor to the scroll bar
    setHandCursor vsPrimary
    
    'When compiled, manifest-themed controls need to be further subclassed so they can have transparent backgrounds.
    If g_IsProgramCompiled And g_IsThemingEnabled And g_IsVistaOrLater Then
        g_Themer.requestContainerSubclass UserControl.hWnd
    End If
            
    'If fancy fonts are being used, increase the horizontal scroll bar height by one pixel equivalent (to make it fit better)
    'If g_UseFancyFonts Then hsPrimary.Height = 23 Else hsPrimary.Height = 22
    
    origForecolor = ForeColor
        
    'Prepare a font object for use
    Set mFont = New StdFont
    Set UserControl.Font = mFont
                    
End Sub

Private Sub UserControl_InitProperties()
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
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    With PropBag
        Set Font = .ReadProperty("Font", Ambient.Font)
        ForeColor = .ReadProperty("ForeColor", &H404040)
        Min = .ReadProperty("Min", 0)
        Max = .ReadProperty("Max", 10)
        SigDigits = .ReadProperty("SigDigits", 0)
        Value = .ReadProperty("Value", 0)
    End With
    
    controlMin = Min
    controlMax = Max
    controlVal = Value
    significantDigits = SigDigits
    
End Sub

Private Sub UserControl_Resize()

    'Keep the text box and scroll bar perfectly aligned, with a 1px border for the red "error" box
    If g_IsProgramCompiled And g_IsVistaOrLater And g_IsThemingEnabled Then
        vsPrimary.Width = fixDPI(19)
        vsPrimary.Top = 0
        vsPrimary.Height = UserControl.ScaleHeight
    Else
        vsPrimary.Width = fixDPI(17)
        vsPrimary.Top = fixDPI(1)
        vsPrimary.Height = UserControl.ScaleHeight - fixDPI(2)
    End If
    
    txtPrimary.Left = 1
    txtPrimary.Top = 1
    txtPrimary.Width = UserControl.ScaleWidth - 2 - vsPrimary.Width
    txtPrimary.Height = UserControl.ScaleHeight - 2
    
    vsPrimary.Left = txtPrimary.Left + txtPrimary.Width
            
    shpError.Left = 0
    shpError.Top = 0
    shpError.Height = UserControl.ScaleHeight
    If g_IsProgramCompiled And g_IsVistaOrLater And g_IsThemingEnabled Then shpError.Width = UserControl.ScaleWidth - 2 Else shpError.Width = UserControl.ScaleWidth
    
End Sub

Private Sub UserControl_Show()
        
    'Apply a hand cursor to the scroll bar.  (It is necessary to do this here, in addition to the _Initialize event,
    ' if the control exists inside a container like a picture box.)
    setHandCursor vsPrimary
        
    'When the control is first made visible, remove the control's tooltip property and reassign it to the checkbox
    ' using a custom solution (which allows for linebreaks and theming).
    m_ToolString = Extender.ToolTipText
    
    If m_ToolString <> "" Then
    
        Set m_ToolTip = New clsToolTip
        With m_ToolTip
        
            .Create Me
            .MaxTipWidth = PD_MAX_TOOLTIP_WIDTH
            .AddTool vsPrimary, m_ToolString
            .AddTool txtPrimary, m_ToolString
            
        End With
        
    End If
        
    'Also, force a resize to modify its layout
    UserControl_Resize
        
End Sub

Private Sub UserControl_Terminate()
    
    'When the control is terminated, release the subclassing used for transparent backgrounds
    If g_IsProgramCompiled And g_IsThemingEnabled And g_IsVistaOrLater Then g_Themer.releaseContainerSubclass UserControl.hWnd
    
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    'Store all associated properties
    With PropBag
        .WriteProperty "Min", controlMin, 0
        .WriteProperty "Max", controlMax, 10
        .WriteProperty "SigDigits", significantDigits, 0
        .WriteProperty "Value", controlVal, 0
        .WriteProperty "Font", mFont, "Tahoma"
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

End Function

'Populate the text box with a given integer value.
Private Sub copyValToTextBox(ByVal srcValue As Double)

    'Remember the current cursor position
    Dim cursorPos As Long
    cursorPos = txtPrimary.SelStart

    'Overwrite the current text box value with the new value
    txtPrimary = getFormattedStringValue(srcValue / (10 ^ significantDigits))
    txtPrimary.Refresh
    
    'Restore the cursor to its original position
    If cursorPos >= Len(txtPrimary) Then cursorPos = Len(txtPrimary)
    txtPrimary.SelStart = cursorPos
    
    'Hide the error box - we know it's not needed, as the value has been set via scroll bar
    If shpError.Visible Then shpError.Visible = False

End Sub

'Check a passed value against a min and max value to see if it is valid.  Additionally, make sure the value is
' numeric, and allow the user to display a warning message if necessary.
Private Function IsTextEntryValid(Optional ByVal displayErrorMsg As Boolean = False) As Boolean
        
    'Some locales use a comma as a decimal separator.  Check for this and replace as necessary.
    Dim chkString As String
    chkString = txtPrimary
    
    'Remember the current cursor position as necessary
    Dim cursorPos As Long
    cursorPos = txtPrimary.SelStart
    
    If InStr(1, chkString, ",") Then
        chkString = Replace(chkString, ",", ".")
        txtPrimary = chkString
        If cursorPos >= Len(txtPrimary) Then cursorPos = Len(txtPrimary)
        txtPrimary.SelStart = cursorPos
    End If
    
    'It may be possible for the user to enter consecutive ",." characters, which then cause the CDbl() below to fail.
    ' Check for this and fix it as necessary.
    If InStr(1, chkString, "..") Then
        chkString = Replace(chkString, "..", ".")
        txtPrimary = chkString
        If cursorPos >= Len(txtPrimary) Then cursorPos = Len(txtPrimary)
        txtPrimary.SelStart = cursorPos
    End If
        
    If Not IsNumeric(chkString) Then
        If displayErrorMsg Then pdMsgBox "%1 is not a valid entry." & vbCrLf & "Please enter a numeric value.", vbExclamation + vbOKOnly + vbApplicationModal, "Invalid entry", txtPrimary
        IsTextEntryValid = False
    Else
        
        Dim checkVal As Double
        checkVal = CDbl(chkString)
    
        If (checkVal >= controlMin) And (checkVal <= controlMax) Then
            IsTextEntryValid = True
        Else
            If displayErrorMsg Then pdMsgBox "%1 is not a valid entry." & vbCrLf & "Please enter a value between %2 and %3.", vbExclamation + vbOKOnly + vbApplicationModal, "Invalid entry", txtPrimary, getFormattedStringValue(controlMin), getFormattedStringValue(controlMax)
            IsTextEntryValid = False
        End If
    End If
    
End Function

Private Sub vsPrimary_GotFocus()
    DestroyCaret
End Sub
