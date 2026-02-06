VERSION 5.00
Begin VB.UserControl pdCommandBarMini 
   Alignable       =   -1  'True
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
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
      Caption         =   "OK"
   End
   Begin PhotoDemon.pdButton cmdCancel 
      Height          =   510
      Left            =   8070
      TabIndex        =   1
      Top             =   120
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   900
      Caption         =   "Cancel"
   End
End
Attribute VB_Name = "pdCommandBarMini"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon "Mini" Command Bar control
'Copyright 2013-2026 by Tanner Helland
'Created: 14/August/13
'Last updated: 23/August/17
'Last update: add automatic handling for Enter/Esc keypresses from child controls
'
'This control is a stripped-down version of the primary CommandBar user control.  It is meant for dialogs where
' save/load preset support is irrelevant, while still supporting the same theming and translation options as
' the standard command bar.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
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
' but I've since wrapped these into a single central support class.
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

'Workaround for VB6 quirks; see VBHacks.InControlArray()
Public Function IsChildInControlArray(ByRef ctlChild As Object) As Boolean
    IsChildInControlArray = Not UserControl.Controls(ctlChild.Name) Is ctlChild
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

'See dialog_InputBox for a use-case for this function
Public Sub ClickCancelForMe()
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
    If (Not m_dontAutoUnloadParent) Then Unload UserControl.Parent
    
End Sub

'OK button
Private Sub CmdOK_Click()
    HandleOKButton
End Sub

'See dialog_InputBox for a use-case for this function
Public Sub ClickOKForMe()
    HandleOKButton
End Sub

Private Sub HandleOKButton()
    
    'Save the current window location
    If (Not g_WindowManager Is Nothing) Then g_WindowManager.SaveWindowLocation UserControl.Parent, False
    
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
    On Error GoTo ParentUnloadedAlready
    UserControl.Parent.Visible = False
ParentUnloadedAlready:
        
    'When everything is done, unload our parent form (unless the override property is set, as it is by default)
    If (Not m_dontAutoUnloadParent) Then Unload UserControl.Parent
    
End Sub

'This control subclasses some internal PD messages, which is how we support "OK" and "Cancel" shortcuts via
' "Enter" and "Esc" keypresses
Private Sub ucSupport_CustomMessage(ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturn As Long)

    If (wMsg = WM_PD_DIALOG_NAVKEY) Then
    
        'This is a relevant navigation key!
        
        'Interpret Enter or Space as OK...
        If (wParam = pdnk_Enter) Or (wParam = pdnk_Space) Then
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

    'Interpret Enter or Space as OK...
    If (whichSysKey = pdnk_Enter) Or (whichSysKey = pdnk_Space) Then
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
    
    'Initialize a user control support class
    Set ucSupport = New pdUCSupport
    ucSupport.RegisterControl UserControl.hWnd, True
    
    'This control can automatically handle "Enter" and "Esc" keypresses coming from its child form.
    ucSupport.SubclassCustomMessage WM_PD_DIALOG_NAVKEY, True
    
    'Prep the color manager and load default colors
    Set m_Colors = New pdThemeColors
    Dim colorCount As PDCB_COLOR_LIST: colorCount = [_Count]
    m_Colors.InitializeColorList "PDCommandBar", colorCount
    If Not PDMain.IsProgramRunning() Then UpdateColorList
    
End Sub

Private Sub UserControl_InitProperties()
    DontAutoUnloadParent = False
End Sub

'At run-time, painting is handled by the support class.  In the IDE, however, we must rely on VB's internal paint event.
Private Sub UserControl_Paint()
    If Not PDMain.IsProgramRunning() Then ucSupport.RequestIDERepaint UserControl.hDC
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    DontAutoUnloadParent = PropBag.ReadProperty("dontAutoUnloadParent", False)
    m_ParentAvailable = True
End Sub

Private Sub UserControl_Resize()
    If Not PDMain.IsProgramRunning() Then ucSupport.RequestRepaint True
End Sub

Private Sub UserControl_Show()
    
    'At run-time, give the OK button focus by default.  (Note that using the .Default property to do this will
    ' BREAK THINGS.  .Default overrides catching the Enter key anywhere else in the form, so we cannot do things
    ' like save a preset via Enter keypress, because the .Default control will always eat the Enter keypress.)
    
    'Additional note: some forms may chose to explicitly set focus away from the OK button.  If that happens, the line below
    ' will throw a critical error.  To avoid that, simply ignore any errors that arise from resetting focus.
    On Error GoTo SomethingStoleFocus
    If PDMain.IsProgramRunning() And (Not g_WindowManager Is Nothing) Then g_WindowManager.SetFocusAPI cmdOK.hWnd

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
    
    'Current command bar height is an arbitrary value
    Dim cmdBarSize As Long
    cmdBarSize = Interface.FixDPI(50)
    
    Dim moveRequired As Boolean
    moveRequired = (bHeight <> cmdBarSize)
    If (Not moveRequired) And (ucSupport.GetControlTop <> parentWindowHeight - cmdBarSize) Then moveRequired = True
    
    If moveRequired Then
        ucSupport.RequestNewSize , cmdBarSize
        ucSupport.RequestNewPosition 0, parentWindowHeight - ucSupport.GetControlHeight()
    End If
    
    'Make the control the same width as its parent
    If PDMain.IsProgramRunning() Then
        
        If (bWidth <> parentWindowWidth) Then ucSupport.RequestNewSize parentWindowWidth
        
        'Right-align the Cancel and OK buttons
        cmdCancel.SetLeft parentWindowWidth - cmdCancel.GetWidth - Interface.FixDPI(8)
        cmdOK.SetLeft cmdCancel.GetLeft - cmdOK.GetWidth - Interface.FixDPI(8)
        
    End If
    
'NOTE: this error catch is important, as VB will attempt to update the user control's size even after the parent has
'       been unloaded, raising error 398 "Client site not available". If we don't catch the error, the compiled .exe
'       will fail every time a command bar is unloaded (e.g. on almost every tool).
SkipUpdateLayout:

End Sub

'Primary rendering function.  Note that ucSupport handles a number of rendering duties (like maintaining a back buffer for us).
Private Sub RedrawBackBuffer()
    
    'Request the back buffer DC, and ask the support module to erase any existing rendering for us.
    Dim bufferDC As Long
    bufferDC = ucSupport.GetBackBufferDC(True, m_Colors.RetrieveColor(PDCB_Background, Me.Enabled))
    If (bufferDC = 0) Then Exit Sub
    
    'Paint the final result to the screen, as relevant
    ucSupport.RequestRepaint
    
End Sub

'Before this control does any painting, we need to retrieve relevant colors from PD's primary theming class.  Note that this
' step must also be called if/when PD's visual theme settings change.
Private Sub UpdateColorList()
    m_Colors.LoadThemeColor PDCB_Background, "Background", IDE_GRAY
End Sub

'External functions can call this to set custom OK button text.  The OK button
' *will* be resized to account for the text change, but the text will *not* get
' auto-translated - that's up to the caller to handle.
Public Sub SetCustomOKText(ByRef newText As String, Optional ByVal desiredPaddingAt96DPI As Long = 16)

    'Before assigning the text, we need to measure it and ensure the OK button is
    ' large enough to fit.
    Dim btnResized As Boolean
    
    'Button resizing requires a valid theme object
    If (Not g_Themer Is Nothing) Then
    
        Dim cFont As pdFont
        Set cFont = New pdFont
        cFont.SetFontSize cmdOK.FontSize()
        cFont.SetFontFace Fonts.GetUIFontName()
        
        Dim pxLength As Long
        pxLength = cFont.GetWidthOfString(newText)
        
        'Minimum amount of required padding on either side of the caption
        ' (including any button border rendering)
        Dim capPadding As Long
        capPadding = Interface.FixDPI(desiredPaddingAt96DPI)
        
        'If the caption is too small, enlarge the button to fit.
        If ((pxLength + capPadding * 2) > cmdOK.GetWidth) Then
            btnResized = True
            cmdOK.SetWidth pxLength + capPadding * 2
        End If
    
    End If
    
    cmdOK.Caption = newText
    
    'Always update the control layout after the button has (likely) been resized
    If btnResized Then UpdateControlLayout

End Sub

'External functions can call this to request a redraw.  This is helpful for live-updating theme settings, as in the Preferences dialog.
Public Sub UpdateAgainstCurrentTheme(Optional ByVal hostFormhWnd As Long = 0)
    
    If ucSupport.ThemeUpdateRequired Then
        
        'Because all controls on the command bar are synchronized against a non-standard backcolor, we need to make sure any new
        ' colors are loaded FIRST
        UpdateColorList
        If PDMain.IsProgramRunning() Then NavKey.NotifyControlLoad Me, hostFormhWnd, False
        If PDMain.IsProgramRunning() Then ucSupport.UpdateAgainstThemeAndLanguage
        
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
