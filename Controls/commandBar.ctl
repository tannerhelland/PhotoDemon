VERSION 5.00
Begin VB.UserControl commandBar 
   Alignable       =   -1  'True
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   ClientHeight    =   750
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9555
   DefaultCancel   =   -1  'True
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   50
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   637
   ToolboxBitmap   =   "commandBar.ctx":0000
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "Reset all options to their default values."
      Top             =   120
      Width           =   900
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   510
      Left            =   8070
      TabIndex        =   1
      ToolTipText     =   "Exit this tool.  No changes will be made to the image."
      Top             =   120
      Width           =   1365
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   510
      Left            =   6600
      TabIndex        =   0
      ToolTipText     =   "Apply the selected action to the current image."
      Top             =   120
      Width           =   1365
   End
End
Attribute VB_Name = "commandBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Tool Dialog Command Bar custom control
'Copyright ©2012-2013 by Tanner Helland
'Created: 14/August/13
'Last updated: 14/August/13
'Last update: initial build
'
'For the first decade of its life, PhotoDemon relied on a simple OK and CANCEL button at the bottom of each tool
' dialog.  These two buttons were dutifully copy+pasted every time a new tool was built, and beyond that they
' received little thought.
'
'As the program has grown more complex, I have wanted to add a variety of new features to each tool - things like
' dedicated "Help" and "Reset" buttons.  Tool presets.  Maybe even a Randomize button.  Adding all these features
' to each tool individually would be a RIDICULOUSLY time-consuming task, so rather than do that, I have wrapped
' all universal tool features into a single command bar, which can be dropped onto any new tool form at will.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://www.tannerhelland.com/photodemon/#license
'
'***************************************************************************

Option Explicit

'Clicking the OK and CANCEL buttons raise their respective events
Public Event OKClick()
Public Event CancelClick()

'Clicking the RESET event raises the corresponding event; note that EACH TOOL MUST IMPLEMENT THIS FUNCTION.
' There is no magical way for me to know default values in advance, so each tool needs to have reset values
' added manually.  Additionally, if no preset values are found, this event will be automatically triggered.
Public Event ResetClick()

'All custom PD controls are auto-validated when OK is pressed.  If other custom items need validation, the OK
' function will trigger this event, which the parent can use as necessary.
Public Event ExtraValidations()

'Used to render images onto the command buttons
Private cImgCtl As clsControlImage

'Custom tooltip class allows for things like multiline, theming, and multiple monitor support
Dim m_ToolTip As clsToolTip

'Font handling for user controls requires some extra work; see below for details
Private WithEvents mFont As StdFont
Attribute mFont.VB_VarHelpID = -1

'Results of extra user validations will be stored here
Private userValidationFailed As Boolean

'If the user wants us to postpone a Cancel-initiated unload, this will let us know
Private dontShutdownYet As Boolean

'If the user wants to postpone a Cancel-initiated unload for some reason, they can call this function during their
' Cancel event.
Public Sub doNotUnloadForm()
    dontShutdownYet = True
End Sub

'If any user-applied extra validations failed, they can call this sub to notify us, and we will adjust our behavior accordingly
Public Sub validationFailed()
    userValidationFailed = True
End Sub

'An hWnd is needed for external tooltip handling
Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
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
    Set cmdOK.Font = mFont
    Set cmdCancel.Font = mFont
    Set cmdReset.Font = mFont
End Sub

'Backcolor is used to control the color of the base user control; nothing else is affected by it
Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal newColor As OLE_COLOR)
    UserControl.BackColor = newColor
    PropertyChanged "BackColor"
End Property

'CANCEL button
Private Sub CmdCancel_Click()

    'The user may have Cancel actions they want to apply - let them do that
    RaiseEvent CancelClick
    
    'If the user asked us to not shutdown yet, obey - otherwise, unload the parent form
    If dontShutdownYet Then
        dontShutdownYet = False
        Exit Sub
    End If
    
    'Automatically unload our parent
    Unload UserControl.Parent
    
End Sub

'OK button
Private Sub CmdOK_Click()
    
    'Automatically validate all relevant controls on the parent object.  This is a huge perk, because it saves us
    ' from having to write validation code individually.
    Dim validateCheck As Boolean
    validateCheck = True
    
    'To validate everything, start by enumerating through each control on our parent form.
    Dim eControl As Object
    For Each eControl In Parent.Controls
        
        'Obviously, we can only validate our own custom objects that have built-in auto-validate functions.
        If (TypeOf eControl Is sliderTextCombo) Or (TypeOf eControl Is textUpDown) Then
            
            'Just to be safe, verify matching container hWnd properties
            If eControl.Container.hWnd = UserControl.ContainerHwnd Then
                
                'Finally, ask the control to validate itself
                If Not eControl.IsValid Then
                    validateCheck = False
                    Exit For
                End If
                
            End If
            
        End If
    Next eControl
    
    'If validation failed, do not proceed further
    If Not validateCheck Then Exit Sub
    
    'Raise an extra validation process, which the parent form can use if necessary to check additional controls.
    RaiseEvent ExtraValidations
    
    'Make sure any extra user validations succeeded
    If userValidationFailed Then
        userValidationFailed = False
        Exit Sub
    End If
    
    'At this point, we are now free to proceed like any normal OK click.
    
    'Hide the parent form from view
    UserControl.Parent.Visible = False
    
    'Finally, let the user proceed with whatever comes next!
    RaiseEvent OKClick
    
    'When everything is done, unload our parent form
    Unload UserControl.Parent
    
End Sub

'RESET button
Private Sub cmdReset_Click()
    RaiseEvent ResetClick
End Sub

Private Sub UserControl_Initialize()

    'Apply the hand cursor to all command buttons
    setHandCursorToHwnd cmdOK.hWnd
    setHandCursorToHwnd cmdCancel.hWnd
    setHandCursorToHwnd cmdReset.hWnd

    'Certain actions are only applied in the compiled EXE
    If g_IsProgramCompiled Then
    
        'Extract relevant icons from the resource file, and render them onto the buttons at run-time.
        ' (NOTE: because the icons require manifest theming, they will not appear in the IDE.)
        Set cImgCtl = New clsControlImage
        With cImgCtl
            cmdReset.Caption = ""
            .LoadImageFromStream cmdReset.hWnd, LoadResData("RESETBUTTON", "CUSTOM"), 24, 24
            '.SetMargins cmdReset.hWnd, , 12
            .Align(cmdReset.hWnd) = Icon_Center
        End With
        
    End If
    
    UserControl.BackColor = BackColor
    
    'Prepare a font object for use
    Set mFont = New StdFont
    Set UserControl.Font = mFont
    
    'Validations succeed by default
    userValidationFailed = False
    
    'Parent forms will be unloaded by default when pressing Cancel
    dontShutdownYet = False
    
End Sub

Private Sub UserControl_InitProperties()
    
    Set mFont = UserControl.Font
    mFont_FontChanged ("")
    BackColor = &HEEEEEE
    
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    
    With PropBag
        Set Font = .ReadProperty("Font", Ambient.Font)
        BackColor = .ReadProperty("BackColor", &HEEEEEE)
    End With
    
End Sub

Private Sub UserControl_Resize()
    updateControlLayout
End Sub

'The command bar's layout is all handled programmatically.  This lets it look good, regardless of the parent form's size.
Private Sub updateControlLayout()

    'Force a standard user control size
    UserControl.Height = 50 * Screen.TwipsPerPixelY
    
    'Make the control the same width as its parent
    UserControl.Width = UserControl.Parent.ScaleWidth * Screen.TwipsPerPixelX
    
    'Right-align the Cancel and OK buttons
    cmdCancel.Left = UserControl.Parent.ScaleWidth - cmdCancel.Width - 8
    cmdOK.Left = cmdCancel.Left - cmdOK.Width - 8

End Sub

Private Sub UserControl_Show()
    
    'When the control is first made visible, rebuild individual tooltips using a custom solution
    ' (which allows for linebreaks and theming).
    If g_UserModeFix Then
        Set m_ToolTip = New clsToolTip
        With m_ToolTip
        
            .Create Me
            .MaxTipWidth = PD_MAX_TOOLTIP_WIDTH
            .AddTool cmdOK, cmdOK.ToolTipText
            cmdOK.ToolTipText = ""
            .AddTool cmdCancel, cmdCancel.ToolTipText
            cmdCancel.ToolTipText = ""
            .AddTool cmdReset, cmdReset.ToolTipText
            cmdReset.ToolTipText = ""
                
        End With
    End If
    
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    
    'Store all associated properties
    With PropBag
        .WriteProperty "Font", mFont, "Tahoma"
        .WriteProperty "BackColor", BackColor, &HEEEEEE
    End With
    
End Sub
