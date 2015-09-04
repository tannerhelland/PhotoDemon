VERSION 5.00
Begin VB.UserControl commandBarMini 
   Alignable       =   -1  'True
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   ClientHeight    =   750
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9555
   ClipControls    =   0   'False
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
   ToolboxBitmap   =   "commandBarMini.ctx":0000
   Begin PhotoDemon.pdButton cmdOK 
      Height          =   510
      Left            =   6600
      TabIndex        =   0
      Top             =   120
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   900
      Caption         =   "&OK"
   End
   Begin PhotoDemon.pdButton cmdCancel 
      Height          =   510
      Left            =   8070
      TabIndex        =   1
      Top             =   120
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   900
      Caption         =   "&Cancel"
   End
End
Attribute VB_Name = "commandBarMini"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon "Mini" Command Bar control
'Copyright 2013-2015 by Tanner Helland
'Created: 14/August/13
'Last updated: 02/September/15
'Last update: separate from the main command bar, to allow for simpler code.
'
'This control is a stripped-down version of the primary CommandBar user control.  It is meant for dialogs where
' save/load preset support is irrelevant, while still supporting the same theming and translation options as
' the standard command bar.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Clicking the OK and CANCEL buttons raise their respective events
Public Event OKClick()
Public Event CancelClick()

'If the user wants us to postpone the automated unload after OK or Cancel is pressed, this will let us know to suspend it.
' (This is controlled via the doNotUnloadForm sub, below, which should be called during the OK or CANCEL events this control raises.)
Private m_dontShutdownYet As Boolean

'If the parent does not want the command bar to auto-unload it when OK or CANCEL is pressed, this will be set to TRUE.
' (This is controlled via property.)
Private m_dontAutoUnloadParent As Boolean

'The command bar is set to auto-unload its parent object when OK or CANCEL is pressed.  In some instances (e.g. forms prefaced with
' "dialog_", which return a VBMsgBoxResult), this behavior is not desirable.  It can be overridden by setting this property to TRUE.
Public Property Get dontAutoUnloadParent() As Boolean
    dontAutoUnloadParent = m_dontAutoUnloadParent
End Property

Public Property Let dontAutoUnloadParent(ByVal newValue As Boolean)
    m_dontAutoUnloadParent = newValue
    PropertyChanged "dontAutoUnloadParent"
End Property

'If the user wants to postpone an OK or Cancel-initiated unload for some reason, they can call this function during their
' Cancel event.
Public Sub doNotUnloadForm()
    m_dontShutdownYet = True
End Sub

'hWnd is used for external focus tracking
Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

'Backcolor is used to control the color of the base user control; nothing else is affected by it.
' Note that - by design - the back color is hardcoded.  Still TODO is integrating it with theming.
Public Property Get BackColor() As OLE_COLOR
    BackColor = RGB(220, 220, 225)
End Property

Public Property Let BackColor(ByVal newColor As OLE_COLOR)
    
    UserControl.BackColor = RGB(220, 220, 225)
    PropertyChanged "BackColor"
    
    'Update all button backgrounds to match
    cmdOK.BackColor = RGB(235, 235, 240)
    cmdCancel.BackColor = RGB(235, 235, 240)
    
End Property

'CANCEL button
Private Sub CmdCancel_Click()

    'The user may have Cancel actions they want to apply - let them do that
    RaiseEvent CancelClick
    
    'If the user asked us to not shutdown yet, obey - otherwise, unload the parent form
    If m_dontShutdownYet Then
        m_dontShutdownYet = False
        Exit Sub
    End If
        
    'Hide the parent form from view
    If UserControl.Parent.Visible Then UserControl.Parent.Hide
        
    'When everything is done, unload our parent form (unless the override property is set, as it is by default)
    If Not m_dontAutoUnloadParent Then Unload UserControl.Parent
    
End Sub

'OK button
Private Sub CmdOK_Click()
    
    'Let the caller know that OK was pressed
    RaiseEvent OKClick
    
    'If the user asked us to not shutdown yet, obey - otherwise, unload the parent form
    If m_dontShutdownYet Then
        m_dontShutdownYet = False
        Exit Sub
    End If
    
    'Hide the parent form from view
    If UserControl.Parent.Visible Then UserControl.Parent.Hide
        
    'When everything is done, unload our parent form (unless the override property is set, as it is by default)
    If Not m_dontAutoUnloadParent Then Unload UserControl.Parent
    
End Sub

Private Sub UserControl_Initialize()
    
    UserControl.BackColor = BackColor
    
    'Parent forms will be unloaded by default when pressing Cancel
    m_dontShutdownYet = False
    
End Sub

Private Sub UserControl_InitProperties()
    BackColor = &HEEEEEE
    dontAutoUnloadParent = False
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    
    With PropBag
        BackColor = .ReadProperty("BackColor", &HEEEEEE)
        dontAutoUnloadParent = .ReadProperty("dontAutoUnloadParent", False)
    End With
    
End Sub

Private Sub UserControl_Resize()
    updateControlLayout
End Sub

'The command bar's layout is all handled programmatically.  This lets it look good, regardless of the parent form's size or
' the current monitor's DPI setting.
Private Sub updateControlLayout()

    On Error GoTo skipUpdateLayout

    'Force a standard user control size
    UserControl.Height = fixDPI(50) * TwipsPerPixelYFix
    
    'Make the control the same width as its parent
    If g_IsProgramRunning Then
    
        UserControl.Width = UserControl.Parent.ScaleWidth * TwipsPerPixelXFix
        
        'Right-align the Cancel and OK buttons
        cmdCancel.Left = UserControl.Parent.ScaleWidth - cmdCancel.Width - fixDPI(8)
        cmdOK.Left = cmdCancel.Left - cmdOK.Width - fixDPI(8)
        
    End If
    
'NOTE: this error catch is important, as VB will attempt to update the user control's size even after the parent has
'       been unloaded, raising error 398 "Client site not available". If we don't catch the error, the compiled .exe
'       will fail every time a command bar is unloaded (e.g. on almost every tool).
skipUpdateLayout:

End Sub

Private Sub UserControl_Show()
    
    'At run-time, give the OK button focus by default.  (Note that using the .Default property to do this will
    ' BREAK THINGS.  .Default overrides catching the Enter key anywhere else in the form, so we cannot do things
    ' like save a preset via Enter keypress, because the .Default control will always eat the Enter keypress.)
    
    'Additional note: some forms may chose to explicitly set focus away from the OK button.  If that happens, the line below
    ' will throw a critical error.  To avoid that, simply ignore any errors that arise from resetting focus.
    On Error GoTo somethingStoleFocus
    If g_IsProgramRunning Then cmdOK.SetFocus

somethingStoleFocus:
    
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    
    'Store all associated properties
    With PropBag
        .WriteProperty "BackColor", BackColor, &HEEEEEE
        .WriteProperty "dontAutoUnloadParent", m_dontAutoUnloadParent, False
    End With
    
End Sub
