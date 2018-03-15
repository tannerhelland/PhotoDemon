VERSION 5.00
Begin VB.UserControl pdCommandBarMini 
   Alignable       =   -1  'True
   Appearance      =   0  'Flat
   ClientHeight    =   750
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9555
   ClipControls    =   0   'False
   DrawStyle       =   5  'Transparent
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HasDC           =   0   'False
   ScaleHeight     =   50
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   637
   ToolboxBitmap   =   "pdCommandBarMini.ctx":0000
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
Attribute VB_Name = "pdCommandBarMini"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon "Mini" Command Bar control
'Copyright 2013-2018 by Tanner Helland
'Created: 14/August/13
'Last updated: 23/August/17
'Last update: add automatic handling for Enter/Esc keypresses from child controls
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

'Like other PD controls, this control raises its own specialized focus events.  If you need to track focus,
' use these instead of the default VB functions.
Public Event GotFocusAPI()
Public Event LostFocusAPI()

'If the user wants us to postpone the automated unload after OK or Cancel is pressed, this will let us know to suspend it.
' (This is controlled via the doNotUnloadForm sub, below, which should be called during the OK or CANCEL events this control raises.)
Private m_dontShutdownYet As Boolean

'If the parent does not want the command bar to auto-unload it when OK or CANCEL is pressed, this will be set to TRUE.
' (This is controlled via property.)
Private m_dontAutoUnloadParent As Boolean

'To avoid "Client Site not available (Error 398)", we wait to access certain parent properties until
' Init/ReadProperty events have fired.  (See MSDN: https://msdn.microsoft.com/en-us/library/aa243344(v=vs.60).aspx)
Private m_ParentAvailable As Boolean

'User control support class.  Historically, many classes (and associated subclassers) were required by each user control,
' but I've since attempted to wrap these into a single master control support class.
Private WithEvents ucSupport As pdUCSupport
Attribute ucSupport.VB_VarHelpID = -1

'Local list of themable colors.  This list includes all potential colors used by this class, regardless of state change
' or internal control settings.  The list is updated by calling the UpdateColorList function.
' (Note also that this list does not include variants, e.g. "BorderColor" vs "BorderColor_Hovered".  Variant values are
'  automatically calculated by the color management class, and they are retrieved by passing boolean modifiers to that
'  class, rather than treating every imaginable variant as a separate constant.)
Private Enum PDCB_COLOR_LIST
    [_First] = 0
    PDCB_Background = 0
    [_Last] = 0
    [_Count] = 1
End Enum

'Color retrieval and storage is handled by a dedicated class; this allows us to optimize theme interactions,
' without worrying about the details locally.
Private m_Colors As pdThemeColors

Public Function GetControlType() As PD_ControlType
    GetControlType = pdct_CommandBarMini
End Function

Public Function GetControlName() As String
    GetControlName = UserControl.Extender.Name
End Function

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

'The command bar is set to auto-unload its parent object when OK or CANCEL is pressed.  In some instances (e.g. forms prefaced with
' "dialog_", which return a VBMsgBoxResult), this behavior is not desirable.  It can be overridden by setting this property to TRUE.
Public Property Get DontAutoUnloadParent() As Boolean
    DontAutoUnloadParent = m_dontAutoUnloadParent
End Property

Public Property Let DontAutoUnloadParent(ByVal newValue As Boolean)
    m_dontAutoUnloadParent = newValue
    PropertyChanged "dontAutoUnloadParent"
End Property

'If the user wants to postpone an OK or Cancel-initiated unload for some reason, they can call this function during their
' Cancel event.
Public Sub DoNotUnloadForm()
    m_dontShutdownYet = True
End Sub

'hWnd is used for external focus tracking
Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

'The Enabled property is a bit unique; see http://msdn.microsoft.com/en-us/library/aa261357%28v=vs.60%29.aspx
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_UserMemId = -514
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal newValue As Boolean)
    UserControl.Enabled = newValue
    PropertyChanged "Enabled"
End Property

'CANCEL button
Private Sub cmdCancel_Click()
    HandleCancelButton
End Sub

Private Sub HandleCancelButton()

    'The user may have Cancel actions they want to apply - let them do that
    RaiseEvent CancelClick
    
    'If the user asked us to not shutdown yet, obey - otherwise, unload the parent form
    If m_dontShutdownYet Then
        m_dontShutdownYet = False
        Exit Sub
    End If
    
    'Notify the central Interface handler that CANCEL was clicked; this lets other functions bypass a UI sync
    Interface.NotifyShowDialogResult vbCancel
    
    'Hide the parent form from view
    If UserControl.Parent.Visible Then UserControl.Parent.Hide
        
    'When everything is done, unload our parent form (unless the override property is set, as it is by default)
    If Not m_dontAutoUnloadParent Then Unload UserControl.Parent
    
End Sub

'OK button
Private Sub CmdOK_Click()
    HandleOKButton
End Sub

Private Sub HandleOKButton()

    'Let the caller know that OK was pressed
    RaiseEvent OKClick
    
    'If the user asked us to not shutdown yet, obey - otherwise, unload the parent form
    If m_dontShutdownYet Then
        m_dontShutdownYet = False
        Exit Sub
    End If
    
    'Notify the central Interface handler that OK was clicked; this lets other functions know that a UI sync is required
    Interface.NotifyShowDialogResult vbOK
    
    'Hide the parent form from view
    On Error Resume Next
    UserControl.Parent.Visible = False
        
    'When everything is done, unload our parent form (unless the override property is set, as it is by default)
    If (Not m_dontAutoUnloadParent) Then Unload UserControl.Parent
    
End Sub

'This control subclasses some internal PD messages, which is how we support "OK" and "Cancel" shortcuts via
' "Enter" and "Esc" keypresses
Private Sub ucSupport_CustomMessage(ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturn As Long)

    If (wMsg = WM_PD_DIALOG_NAVKEY) Then
    
        'This is a relevant navigation key!
        
        'Interpret Enter as OK...
        If (wParam = pdnk_Enter) Then
            HandleOKButton
            bHandled = True
            
        '...and Esc as CANCEL.
        ElseIf (wParam = pdnk_Escape) Then
            HandleCancelButton
            bHandled = True
        End If
        
    End If

End Sub

Private Sub ucSupport_GotFocusAPI()
    RaiseEvent GotFocusAPI
End Sub

'If the command bar itself has focus, manually handle Enter/Esc as OK/Cancel events
Private Sub ucSupport_KeyDownSystem(ByVal Shift As ShiftConstants, ByVal whichSysKey As PD_NavigationKey, markEventHandled As Boolean)

    'Interpret Enter as OK...
    If (whichSysKey = pdnk_Enter) Then
        markEventHandled = True
        HandleOKButton
        
    '...and Esc as CANCEL.
    ElseIf (whichSysKey = pdnk_Escape) Then
        markEventHandled = True
        HandleCancelButton
        
    End If

End Sub

Private Sub ucSupport_LostFocusAPI()
    RaiseEvent LostFocusAPI
End Sub

Private Sub ucSupport_RepaintRequired(ByVal updateLayoutToo As Boolean)
    If updateLayoutToo Then UpdateControlLayout
    RedrawBackBuffer
End Sub

Private Sub UserControl_Initialize()
    
    'Parent forms will be unloaded by default when pressing Cancel
    m_dontShutdownYet = False
    
    'Initialize a master user control support class
    Set ucSupport = New pdUCSupport
    ucSupport.RegisterControl UserControl.hWnd, True
    
    'This control can automatically handle "Enter" and "Esc" keypresses coming from its child form.
    ucSupport.SubclassCustomMessage WM_PD_DIALOG_NAVKEY, True
    
    'Prep the color manager and load default colors
    Set m_Colors = New pdThemeColors
    Dim colorCount As PDCB_COLOR_LIST: colorCount = [_Count]
    m_Colors.InitializeColorList "PDCommandBar", colorCount
    If Not pdMain.IsProgramRunning() Then UpdateColorList
    
    'Update the control size parameters at least once
    UpdateControlLayout
    
End Sub

Private Sub UserControl_InitProperties()
    DontAutoUnloadParent = False
End Sub

'At run-time, painting is handled by the support class.  In the IDE, however, we must rely on VB's internal paint event.
Private Sub UserControl_Paint()
    If Not pdMain.IsProgramRunning() Then ucSupport.RequestIDERepaint UserControl.hDC
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    DontAutoUnloadParent = PropBag.ReadProperty("dontAutoUnloadParent", False)
    m_ParentAvailable = True
End Sub

Private Sub UserControl_Resize()
    If Not pdMain.IsProgramRunning() Then ucSupport.RequestRepaint True
End Sub

Private Sub UserControl_Show()
    
    'At run-time, give the OK button focus by default.  (Note that using the .Default property to do this will
    ' BREAK THINGS.  .Default overrides catching the Enter key anywhere else in the form, so we cannot do things
    ' like save a preset via Enter keypress, because the .Default control will always eat the Enter keypress.)
    
    'Additional note: some forms may chose to explicitly set focus away from the OK button.  If that happens, the line below
    ' will throw a critical error.  To avoid that, simply ignore any errors that arise from resetting focus.
    On Error GoTo SomethingStoleFocus
    If pdMain.IsProgramRunning() And (Not g_WindowManager Is Nothing) Then g_WindowManager.SetFocusAPI cmdOK.hWnd

SomethingStoleFocus:
    
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "DontAutoUnloadParent", m_dontAutoUnloadParent, False
End Sub

'The command bar's layout is all handled programmatically.  This lets it look good, regardless of the parent form's size or
' the current monitor's DPI setting.
Private Sub UpdateControlLayout()

    On Error GoTo SkipUpdateLayout
    
    If (Not m_ParentAvailable) Then Exit Sub
    
    'Retrieve DPI-aware control dimensions from the support class
    Dim bWidth As Long, bHeight As Long
    bWidth = ucSupport.GetControlWidth
    bHeight = ucSupport.GetControlHeight
    
    'Force a standard user control size and bottom-alignment
    Dim parentWindowWidth As Long, parentWindowHeight As Long
    parentWindowWidth = g_WindowManager.GetClientWidth(UserControl.Parent.hWnd)
    parentWindowHeight = g_WindowManager.GetClientHeight(UserControl.Parent.hWnd)
    
    Dim moveRequired As Boolean
    If bHeight <> FixDPI(50) Then moveRequired = True
    If ucSupport.GetControlTop <> parentWindowHeight - FixDPI(50) Then moveRequired = True
    
    If moveRequired Then
        ucSupport.RequestNewSize , FixDPI(50)
        ucSupport.RequestNewPosition 0, parentWindowHeight - ucSupport.GetControlHeight
    End If
    
    'Make the control the same width as its parent
    If pdMain.IsProgramRunning() Then
        
        If bWidth <> parentWindowWidth Then ucSupport.RequestNewSize parentWindowWidth
        
        'Right-align the Cancel and OK buttons
        cmdCancel.SetLeft parentWindowWidth - cmdCancel.GetWidth - FixDPI(8)
        cmdOK.SetLeft cmdCancel.GetLeft - cmdOK.GetWidth - FixDPI(8)
        
    End If
    
'NOTE: this error catch is important, as VB will attempt to update the user control's size even after the parent has
'       been unloaded, raising error 398 "Client site not available". If we don't catch the error, the compiled .exe
'       will fail every time a command bar is unloaded (e.g. on almost every tool).
SkipUpdateLayout:

End Sub

'Primary rendering function.  Note that ucSupport handles a number of rendering duties (like maintaining a back buffer for us).
Private Sub RedrawBackBuffer()
    
    'We can improve shutdown performance by ignoring redraw requests
    If g_ProgramShuttingDown Then
        If (g_Themer Is Nothing) Then Exit Sub
    End If
    
    'Request the back buffer DC, and ask the support module to erase any existing rendering for us.
    Dim bufferDC As Long
    bufferDC = ucSupport.GetBackBufferDC(True, m_Colors.RetrieveColor(PDCB_Background, Me.Enabled))
    
    'Paint the final result to the screen, as relevant
    ucSupport.RequestRepaint
    
End Sub

'Before this control does any painting, we need to retrieve relevant colors from PD's primary theming class.  Note that this
' step must also be called if/when PD's visual theme settings change.
Private Sub UpdateColorList()
    m_Colors.LoadThemeColor PDCB_Background, "Background", IDE_GRAY
End Sub

'External functions can call this to request a redraw.  This is helpful for live-updating theme settings, as in the Preferences dialog.
Public Sub UpdateAgainstCurrentTheme(Optional ByVal hostFormhWnd As Long = 0)
    
    If ucSupport.ThemeUpdateRequired Then
        
        'Because all controls on the command bar are synchronized against a non-standard backcolor, we need to make sure any new
        ' colors are loaded FIRST
        UpdateColorList
        If pdMain.IsProgramRunning() Then NavKey.NotifyControlLoad Me, hostFormhWnd
        If pdMain.IsProgramRunning() Then ucSupport.UpdateAgainstThemeAndLanguage
        
        Dim cbBackgroundColor As Long
        cbBackgroundColor = m_Colors.RetrieveColor(PDCB_Background, Me.Enabled)
        
        'Synchronize the background color of individual controls against the command bar's backcolor
        cmdOK.BackgroundColor = cbBackgroundColor
        cmdCancel.BackgroundColor = cbBackgroundColor
        cmdOK.UseCustomBackgroundColor = True
        cmdCancel.UseCustomBackgroundColor = True
        cmdOK.UpdateAgainstCurrentTheme
        cmdCancel.UpdateAgainstCurrentTheme
        
    End If
    
End Sub
