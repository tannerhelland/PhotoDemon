VERSION 5.00
Begin VB.UserControl smartCheckBox 
   BackColor       =   &H80000005&
   ClientHeight    =   375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2520
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   25
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   168
   ToolboxBitmap   =   "smartCheckBox.ctx":0000
   Begin VB.CheckBox chkBox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   30
      Width           =   255
   End
   Begin VB.Label lblCaption 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "caption"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   345
      TabIndex        =   0
      Top             =   0
      UseMnemonic     =   0   'False
      Width           =   765
   End
End
Attribute VB_Name = "smartCheckBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon "Smart" Check Box custom control
'Copyright ©2012-2013 by Tanner Helland
'Created: 29/January/13
'Last updated: 29/January/13
'Last update: initial build
'
'Intrinsic VB checkboxes have a number of limitations.  Most obnoxious is the lack of an "autosize" for the caption.
' Now that PhotoDemon has full translation support, checkboxes are frequently resized at run-time, and if a
' translated phrase is longer than the original one, the text may be cut off.
'
'So I wrote my own replacement checkbox.  My version has a few important benefits:
' 1) Autosize based on the caption is properly supported.
' 2) A hand cursor is automatically applied, and clicks on both the box and label are registered properly.
' 3) The text forecolor can be changed even when a manifest is applied, unlike regular checkboxes.
' 4) Subclassing is used to make the checkbox background properly transparent - but only when compiled.
' 5) Font changes are automatically handled internally.
'
'This control is quite a bit simpler than its option button counterpart, on account of not having to interact with
' other controls on the form.  That said, the two share many features, so changes to one should probably be
' mirrored to the other.
'
'***************************************************************************

Option Explicit

'This object really only needs one event raised - Click
Public Event Click()

Private WithEvents mFont As StdFont
Attribute mFont.VB_VarHelpID = -1

Private origForecolor As Long

'Used to render themed and multiline tooltips
Private m_ToolTip As clsToolTip
Private m_ToolString As String

'The Enabled property is a bit unique; see http://msdn.microsoft.com/en-us/library/aa261357%28v=vs.60%29.aspx
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_UserMemId = -514
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal newValue As Boolean)
    UserControl.Enabled = newValue
    chkBox.Enabled = newValue
    
    'Also change the label color to help indicate disablement
    If newValue Then lblCaption.ForeColor = origForecolor Else lblCaption.ForeColor = vbGrayText
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

Private Sub chkBox_Click()
    If chkBox.Value = vbChecked Then Value = vbChecked Else Value = vbUnchecked
End Sub

'Setting Value to true will automatically raise all necessary external events and redraw the control
Private Sub lblCaption_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Value = vbChecked Then
        chkBox.Value = vbUnchecked
        Value = vbUnchecked
    Else
        chkBox.Value = vbChecked
        Value = vbChecked
    End If
End Sub

Private Sub mFont_FontChanged(ByVal PropertyName As String)
    Set UserControl.Font = mFont
    Set lblCaption.Font = UserControl.Font
    updateControlSize
End Sub

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

'The control's value is simply a reflection of the embedded check box
Public Property Get Value() As CheckBoxConstants
Attribute Value.VB_UserMemId = 0
    Value = chkBox.Value
End Property

Public Property Let Value(ByVal newValue As CheckBoxConstants)
    chkBox.Value = newValue
    PropertyChanged "Value"
    RaiseEvent Click
End Property

'The control's caption is simply passed on to the label
Public Property Get Caption() As String
Attribute Caption.VB_UserMemId = -518
    Caption = lblCaption.Caption
End Property

Public Property Let Caption(ByVal newCaption As String)
    lblCaption.Caption = newCaption
    lblCaption.Refresh
    updateControlSize
    PropertyChanged "Caption"
End Property

'Forecolor is used to control the color of only the label; nothing else is affected by it
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = lblCaption.ForeColor
End Property

Public Property Let ForeColor(ByVal newColor As OLE_COLOR)
    lblCaption.ForeColor = newColor
    origForecolor = newColor
    PropertyChanged "ForeColor"
End Property

'I handle focus a little differently in the IDE, namely by underlining the label caption.  (This makes debugging simpler.)  When compiled, a focus rectangle is drawn around the control only when it HAS FOCUS and IS NOT TRUE.  If
Private Sub UserControl_EnterFocus()
    If Not g_IsProgramCompiled Then lblCaption.Font.Underline = True
End Sub

Private Sub UserControl_ExitFocus()
    If Not g_IsProgramCompiled Then lblCaption.Font.Underline = False
End Sub

Private Sub UserControl_Initialize()
    
    'Apply a hand cursor to the entire control (good enough for the IDE) and also the option button (when compiled)
    setHandCursorToHwnd UserControl.hWnd
    setHandCursor chkBox
    
    'When compiled, manifest-themed controls need to be further subclassed so they can have transparent backgrounds.
    If g_IsProgramCompiled And g_IsThemingEnabled And g_IsVistaOrLater Then
        SubclassFrame UserControl.hWnd, False
        chkBox.ZOrder 0
    End If
      
    origForecolor = ForeColor
    lblCaption.ForeColor = ForeColor
        
    'Prepare a font object for use
    Set mFont = New StdFont
    Set UserControl.Font = mFont
                
End Sub

Private Sub UserControl_InitProperties()
    Caption = "caption"
    Set mFont = UserControl.Font
    mFont_FontChanged ("")
    ForeColor = &H404040
    origForecolor = ForeColor
    Value = vbUnchecked
    updateControlSize
End Sub

'For responsiveness, MouseDown is used instead of Click
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
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
        Caption = .ReadProperty("Caption", "")
        Set Font = .ReadProperty("Font", Ambient.Font)
        ForeColor = .ReadProperty("ForeColor", &H404040)
        Value = .ReadProperty("Value", vbUnchecked)
    End With

End Sub

Private Sub UserControl_Show()
    
    'When the control is first made visible, remove the control's tooltip property and reassign it to the checkbox
    ' using a custom solution (which allows for linebreaks and theming).
    m_ToolString = Extender.ToolTipText
    
    If m_ToolString <> "" Then
    
        Set m_ToolTip = New clsToolTip
        With m_ToolTip
        
            .Create Me
            .MaxTipWidth = PD_MAX_TOOLTIP_WIDTH
            .AddTool chkBox, m_ToolString
            
        End With
        
    End If
        
End Sub

Private Sub UserControl_Terminate()
    
    'When the control is terminated, release the subclassing used for transparent backgrounds
    If g_IsProgramCompiled Then SubclassFrame UserControl.hWnd, True
    
End Sub

'The control dynamically resizes to match the dimensions of the caption.  The size cannot be set by the user.
Private Sub updateControlSize()

    'Force the height to match that of the label
    Dim newHeight As Long
    newHeight = (lblCaption.Height * 2) '* Screen.TwipsPerPixelY
    If newHeight < chkBox.Height Then newHeight = chkBox.Height
    UserControl.Height = newHeight * Screen.TwipsPerPixelY
    UserControl.Width = (lblCaption.Left + lblCaption.Width + 2) * Screen.TwipsPerPixelX
    
    'Center the option button vertically
    Dim hModifier As Long
    hModifier = -1
    If g_IsProgramCompiled And g_UseFancyFonts Then hModifier = 0
    
    'Center-align the check box vertically
    chkBox.Top = (UserControl.ScaleHeight - chkBox.Height) \ 2
    
    'Align the label with the check box
    lblCaption.Top = (chkBox.Top + chkBox.Height) - lblCaption.Height + hModifier
    
    'When compiled, set the option button to be the full size of the user control.  Thanks to subclassing, the option
    ' button will still be fully transparent.  This allows the caption to be seen, while also allowing Vista/7's
    ' "hover" animation to still work with the mouse.  In the IDE, an underline is used to display focus.
    If g_IsProgramCompiled And g_IsThemingEnabled And g_IsVistaOrLater Then chkBox.Width = UserControl.ScaleWidth - 2
            
    lblCaption.Refresh
    chkBox.Refresh
            
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    'Store all associated properties
    With PropBag
        .WriteProperty "Caption", Caption, "caption"
        .WriteProperty "Value", Value, vbUnchecked
        .WriteProperty "Font", mFont, "Tahoma"
        .WriteProperty "ForeColor", ForeColor, &H404040
    End With
    
End Sub

