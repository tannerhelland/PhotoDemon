VERSION 5.00
Begin VB.UserControl textSliderCombo 
   BackColor       =   &H80000005&
   ClientHeight    =   495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6015
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   33
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   401
   Begin VB.HScrollBar hsPrimary 
      Height          =   285
      Left            =   120
      Max             =   100
      Min             =   -100
      TabIndex        =   1
      Top             =   90
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
      MaxLength       =   4
      TabIndex        =   0
      Text            =   "0"
      Top             =   60
      Width           =   735
   End
End
Attribute VB_Name = "textSliderCombo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Text / Slider custom control
'Copyright ©2012-2013 by Tanner Helland
'Created: 19/April/13
'Last updated: 19/April/13
'Last update: initial build
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
' 2) Hand cursor automatically applied to the slider
' 3) Validation of text entries
' 4) Locale handling (like the aforementioned comma/decimal replacement in some locales)
' 5) A single "Change" event that fires for either scroll or text changes, and only if a text change is valid
'
'***************************************************************************

Option Explicit

'This object really only needs one event raised - Change.  It triggers when either the scrollbar or text box
' is modified.
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

'The Enabled property is a bit unique; see http://msdn.microsoft.com/en-us/library/aa261357%28v=vs.60%29.aspx
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_UserMemId = -514
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal NewValue As Boolean)
    UserControl.Enabled = NewValue
    hsPrimary.Enabled = NewValue
    txtPrimary.Enabled = NewValue
    PropertyChanged "Enabled"
End Property

'Font handling is a bit specialized for user controls; see http://msdn.microsoft.com/en-us/library/aa261313%28v=vs.60%29.aspx
Public Property Get Font() As StdFont
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

'Private Sub chkBox_Click()
'    If chkBox.Value = vbChecked Then Value = vbChecked Else Value = vbUnchecked
'End Sub

'Setting Value to true will automatically raise all necessary external events and redraw the control
'Private Sub lblCaption_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'If Value = vbChecked Then
    '    chkBox.Value = vbUnchecked
    '    Value = vbUnchecked
    'Else
    '    chkBox.Value = vbChecked
    '    Value = vbChecked
    'End If
'End Sub

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
    controlVal = NewValue
    hsPrimary.Value = CLng(controlVal)
    txtPrimary.Text = CStr(controlVal)
    PropertyChanged "Value"
    RaiseEvent Change
End Property

'Note: the control's minimum value is settable at run-time
Public Property Get Min() As Double
    Min = controlMin
End Property

Public Property Let Min(ByVal NewValue As Double)
    controlMin = NewValue
    If hsPrimary < CLng(controlMin) Then
        controlVal = controlMin
        hsPrimary = controlVal
        txtPrimary = controlVal
        RaiseEvent Change
    End If
    hsPrimary.Min = CLng(controlMin)
    PropertyChanged "Min"
End Property

'Note: the control's maximum value is settable at run-time
Public Property Get Max() As Double
    Max = controlMax
End Property

Public Property Let Max(ByVal NewValue As Double)
    controlMax = NewValue
    If hsPrimary > CLng(controlMax) Then
        controlVal = controlMax
        hsPrimary = controlVal
        txtPrimary = controlVal
        RaiseEvent Change
    End If
    hsPrimary.Max = CLng(controlMax)
    PropertyChanged "Max"
End Property

'Significant digits determines whether the control allows float values or int values (and with how much precision)
Public Property Get SigDigits() As Double
    SigDigits = significantDigits
End Property

Public Property Let SigDigits(ByVal NewValue As Long)
    significantDigits = SigDigits
    PropertyChanged "SigDigits"
End Property

'Forecolor may be used in the future as part of theming, but right now it serves no purpose
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = lblCaption.ForeColor
End Property

Public Property Let ForeColor(ByVal newColor As OLE_COLOR)
    origForecolor = newColor
    PropertyChanged "ForeColor"
End Property

Private Sub UserControl_Initialize()
    
    'Apply a hand cursor to the entire control (good enough for the IDE) and also the option button (when compiled)
    setHandCursor hsPrimary
    
    'When compiled, manifest-themed controls need to be further subclassed so they can have transparent backgrounds.
    If g_IsProgramCompiled And g_IsThemingEnabled And g_IsVistaOrLater Then
        SubclassFrame UserControl.hWnd, False
    End If
      
    origForecolor = ForeColor
        
    'Prepare a font object for use
    Set mFont = New StdFont
    Set UserControl.Font = mFont
                
End Sub

Private Sub UserControl_InitProperties()
    Set mFont = UserControl.Font
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

'For responsiveness, MouseDown is used instead of Click
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Value = vbChecked Then
        chkBox.Value = vbUnchecked
        Value = vbUnchecked
    Else
        chkBox.Value = vbChecked
        Value = vbChecked
    End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    With PropBag
        Set Font = .ReadProperty("Font", Ambient.Font)
        ForeColor = .ReadProperty("ForeColor", &H404040)
        Value = .ReadProperty("Value", 0)
        Min = .ReadProperty("Min", 0)
        Max = .ReadProperty("Max", 10)
        SigDigits = .ReadProperty("SigDigits", 0)
    End With
    
    controlMin = Min
    controlMax = Max
    controlVal = Value
    significantDigits = SigDigits

End Sub

Private Sub UserControl_Show()
    
    'When the control is first made visible, remove the control's tooltip property and reassign it to the checkbox
    ' using a custom solution (which allows for linebreaks and theming).
    m_ToolString = Extender.ToolTipText
    
    If m_ToolString <> "" Then
    
        Set m_ToolTip = New clsToolTip
        With m_ToolTip
        
            .Create Me
            .MaxTipWidth = 400
            .AddTool hsPrimary, m_ToolString
            .AddTool txtPrimary, m_ToolString
            
        End With
        
    End If
        
End Sub

Private Sub UserControl_Terminate()
    
    'When the control is terminated, release the subclassing used for transparent backgrounds
    If g_IsProgramCompiled Then SubclassFrame UserControl.hWnd, True
    
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    'Store all associated properties
    With PropBag
        .WriteProperty "Min", controlMin, 0
        .WriteProperty "Max", controlMax, 10
        .WriteProperty "Value", controlVal, 0
        .WriteProperty "SigDigits", significantDigits, 0
        .WriteProperty "Font", mFont, "Tahoma"
        .WriteProperty "ForeColor", ForeColor, &H404040
    End With
    
End Sub
