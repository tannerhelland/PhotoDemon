VERSION 5.00
Begin VB.UserControl smartOptionButton 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3270
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
   ScaleHeight     =   26
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   218
   ToolboxBitmap   =   "smartOptionButton.ctx":0000
   Begin VB.CheckBox chkFirst 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1440
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.OptionButton optButton 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   15
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   30
      Width           =   240
   End
   Begin VB.Label lblCaption 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "caption"
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   360
      TabIndex        =   2
      Top             =   0
      UseMnemonic     =   0   'False
      Width           =   765
   End
End
Attribute VB_Name = "smartOptionButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon "Smart" Option Button custom control
'Copyright ©2012-2013 by Tanner Helland
'Created: 28/January/13
'Last updated: 28/January/13
'Last update: initial build
'
'Intrinsic VB option buttons have a number of limitations.  Most obnoxious is the lack of an "autosize" for the caption.
' Now that PhotoDemon has full translation support, option buttons are frequently resized at run-time, and if a
' translated phrase is longer than the original one, the text may be cut off.
'
'So I wrote my own replacement option button.  This option button has a few important benefits:
' 1) Autosize based on the caption is properly supported.
' 2) A hand cursor is automatically applied, and clicks on both the button and label are registered properly.
' 3) The text forecolor can be changed even when a manifest is applied, unlike regular option buttons.
' 4) Subclassing is used to make the option button properly transparent - but only when compiled.
' 5) Font changes are automatically handled internally.
'
'Note that some odd workarounds are required to handle quirks of the way VB treats focus events for option buttons.
' If an option button is the only one in a container, when it receives focus, it will automatically have its value
' set to TRUE.  This obviously doesn't work for our control, which has a single option button in its container.
'
'I work around this by using a hidden check box to receive focus initially.  The user control is still drawn as
' "in focus" (via caption underlining in the IDE), despite a hidden control receiving focus.  If the spacebar is
' used to toggle the option button value, we also handle this properly.
'
'One other quirk of VB option buttons is that whenever they are set to TRUE, their TabStop value is also set to
' TRUE.  We forcibly return it to FALSE - otherwise, the user would have to tab through both the option button
' and the check box on each user control, and we don't want that.
'
'***************************************************************************

Option Explicit

'API technique for drawing a focus rectangle
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Declare Function DrawFocusRect Lib "user32" (ByVal hDC As Long, lpRect As RECT) As Long
Private drewFocusRect As Boolean

'This function really only needs one event raised - Click
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

Public Property Let Enabled(ByVal NewValue As Boolean)
    UserControl.Enabled = NewValue
    optButton.Enabled = NewValue
    
    'Also change the label color to help indicate disablement
    If NewValue Then lblCaption.ForeColor = origForecolor Else lblCaption.ForeColor = vbGrayText
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

Private Sub mFont_FontChanged(ByVal PropertyName As String)
    Set UserControl.Font = mFont
    Set lblCaption.Font = UserControl.Font
    updateControlSize
End Sub

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

'The control's value is simply a reflection of the embedded option button
Public Property Get Value() As Boolean
Attribute Value.VB_UserMemId = 0
    Value = optButton.Value
End Property

Public Property Let Value(ByVal NewValue As Boolean)
    optButton.Value = NewValue
    updateValue
    PropertyChanged "Value"
    If Value Then
        optButton.TabStop = False
        RaiseEvent Click
    End If
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

'A hidden checkbox is used for reasons mentioned at the top of this page
Private Sub chkFirst_Click()
    If Value <> True Then Value = True
    optButton.TabStop = False
End Sub

'Setting Value to true will automatically raise all necessary external events and redraw the control
Private Sub lblCaption_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Value = True
End Sub

Private Sub optButton_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Value = True
End Sub

'I handle focus a little differently than Windows; in the IDE, the label caption is underlined.  (This makes debugging
' simpler.)  When compiled, a focus rectangle is drawn around the control only when it HAS FOCUS and IS NOT TRUE.  If
' the control is set to TRUE, there is no need for a focus rectangle.  But if it isn't TRUE, the focus rectangle makes
' sense, becuase the user needs to be notified what control will be toggled.
Private Sub UserControl_EnterFocus()
    If Not g_IsProgramCompiled Then lblCaption.Font.Underline = True
    If g_IsProgramCompiled And (Value = False) Then updateFocusRect True
End Sub

Private Sub UserControl_ExitFocus()
    If Not g_IsProgramCompiled Then lblCaption.Font.Underline = False
    If drewFocusRect Then updateFocusRect False
End Sub

'This can be used to draw or remove a focus rectangle
Private Sub updateFocusRect(ByVal newStatus As Boolean)
    Dim tmpRect As RECT
    With tmpRect
        .Left = 0
        .Top = 0
        .Right = UserControl.ScaleWidth
        .Bottom = UserControl.ScaleHeight
    End With
    DrawFocusRect UserControl.hDC, tmpRect
    drewFocusRect = newStatus
End Sub

Private Sub UserControl_Initialize()
    
    'Apply a hand cursor to the entire control (good enough for the IDE) and also the option button (when compiled)
    setHandCursorToHwnd UserControl.hWnd
    setHandCursor optButton
    
    'When compiled, manifest-themed controls need to be further subclassed so they can have transparent backgrounds.
    If g_IsProgramCompiled And g_IsThemingEnabled And g_IsVistaOrLater Then
        SubclassFrame UserControl.hWnd, False
        optButton.ZOrder 0
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
    Value = False
    updateControlSize
End Sub

'For responsiveness, MouseDown is used instead of Click
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Value = True
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    With PropBag
        Caption = .ReadProperty("Caption", "")
        Set Font = .ReadProperty("Font", Ambient.Font)
        ForeColor = .ReadProperty("ForeColor", &H404040)
        Value = .ReadProperty("Value", False)
    End With

End Sub

'The control dynamically resizes to match the dimensions of the caption.  The size cannot be set by the user.
Private Sub UserControl_Resize()
    updateControlSize
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
            .AddTool optButton, m_ToolString
            
        End With
        
    End If
    
End Sub

Private Sub UserControl_Terminate()
    
    'When the control is terminated, release the subclassing used for transparent backgrounds
    If g_IsProgramCompiled Then SubclassFrame UserControl.hWnd, True
    
End Sub

Private Sub updateControlSize()

    'Segoe UI requires a slightly different layout
    If g_UseFancyFonts Then
        If g_IsProgramCompiled Then lblCaption.Top = 0 Else lblCaption.Top = 2
    Else
        If g_IsProgramCompiled Then lblCaption.Top = 2 Else lblCaption.Top = 1
    End If
    
    'Force the height to match that of the label
    Dim fontModifier As Long
    If (g_UseFancyFonts And g_IsProgramCompiled) Then fontModifier = 1 Else fontModifier = 4
    UserControl.Height = (lblCaption.Height + lblCaption.Top * 2 + fontModifier) * Screen.TwipsPerPixelY
    UserControl.Width = (lblCaption.Left + lblCaption.Width + 2) * Screen.TwipsPerPixelX
    
    'Center the option button vertically
    optButton.Top = (UserControl.ScaleHeight - optButton.Height) \ 2
    
    'When compiled, set the option button to be the full size of the user control.  Thanks to subclassing, the option
    ' button will still be fully transparent.  This allows the caption to be seen, while also allowing Vista/7's
    ' "hover" animation to still work with the mouse.  In the IDE, an underline is used to display focus.
    If g_IsProgramCompiled And g_IsThemingEnabled And g_IsVistaOrLater Then optButton.Width = UserControl.ScaleWidth - 2
            
    lblCaption.Refresh
    optButton.Refresh
            
End Sub

'Because this is an option control (not a checkbox), other option controls need to be turned off when it is clicked
Private Sub updateValue()

    'If the option button is set to TRUE, turn off all other option buttons on a form
    If optButton.Value Then

        'Enumerate through each control on the form; if it's another option button, turn it OFF
        Dim eControl As Object
        For Each eControl In Parent.Controls
            If TypeOf eControl Is smartOptionButton Then
                If eControl.Container.hWnd = UserControl.ContainerHwnd Then
                    If Not (eControl.hWnd = UserControl.hWnd) Then eControl.Value = False
                End If
            End If
        Next eControl
    
    End If
    
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    'Store all associated properties
    With PropBag
        .WriteProperty "Caption", Caption, "caption"
        .WriteProperty "Value", Value, False
        .WriteProperty "Font", mFont, "Tahoma"
        .WriteProperty "ForeColor", ForeColor, &H404040
    End With
    
End Sub
