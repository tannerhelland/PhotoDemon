VERSION 5.00
Begin VB.UserControl pdGradientSelector 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
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
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "pdGradientSelector.ctx":0000
End
Attribute VB_Name = "pdGradientSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Gradient Selector custom control
'Copyright 2015-2018 by Tanner Helland
'Created: 23/July/15
'Last updated: 01/February/16
'Last update: finalize theming support
'
'This thin user control is basically an empty control that when clicked, displays a gradient editor window.  If a
' gradient is selected (e.g. Cancel is not pressed), it updates its appearance to match, and raises a "GradientChanged"
' event.
'
'Though simple, this control solves a lot of problems.  It is especially helpful for improving interaction with the
' command bar user control, as it easily supports gradient reset/randomize/preset events.  It is also nice to be able
' to update a single master function for gradient selection, then have the change propagate to all tool windows.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************


Option Explicit

'This control doesn't really do anything interesting, besides allow a gradient to be selected.
Public Event GradientChanged()

'Because VB focus events are wonky, especially when we use CreateWindow within a UC, this control raises its own
' specialized focus events.  If you need to track focus, use these instead of the default VB functions.
Public Event GotFocusAPI()
Public Event LostFocusAPI()

'The control's current gradient settings
Private m_curGradient As String

'Temporary brush object, used to render the gradient preview
Private m_Brush As pd2DBrush

'When the "select gradient" dialog is live, this will be set to TRUE
Private isDialogLive As Boolean

'The rectangle where the gradient preview is actually rendered, and a boolean to track whether the mouse is inside that rect
Private m_GradientRect As RectF, m_MouseInsideGradientRect As Boolean, m_MouseDownGradientRect As Boolean

'User control support class.  Historically, many classes (and associated subclassers) were required by each user control,
' but I've since attempted to wrap these into a single master control support class.
Private WithEvents ucSupport As pdUCSupport
Attribute ucSupport.VB_VarHelpID = -1

'Local list of themable colors.  This list includes all potential colors used by this class, regardless of state change
' or internal control settings.  The list is updated by calling the UpdateColorList function.
' (Note also that this list does not include variants, e.g. "BorderColor" vs "BorderColor_Hovered".  Variant values are
'  automatically calculated by the color management class, and they are retrieved by passing boolean modifiers to that
'  class, rather than treating every imaginable variant as a separate constant.)
Private Enum PDGS_COLOR_LIST
    [_First] = 0
    PDGS_Border = 0
    [_Last] = 0
    [_Count] = 1
End Enum

'Color retrieval and storage is handled by a dedicated class; this allows us to optimize theme interactions,
' without worrying about the details locally.
Private m_Colors As pdThemeColors

Public Function GetControlType() As PD_ControlType
    GetControlType = pdct_GradientSelector
End Function

Public Function GetControlName() As String
    GetControlName = UserControl.Extender.Name
End Function

'Caption is handled just like the common control label's caption property.  It is valid at design-time, and any translation,
' if present, will not be processed until run-time.
' IMPORTANT NOTE: only the ENGLISH caption is returned.  I don't have a reason for returning a translated caption (if any),
'                  but I can revisit in the future if it ever becomes relevant.
Public Property Get Caption() As String
Attribute Caption.VB_UserMemId = -518
    Caption = ucSupport.GetCaptionText()
End Property

Public Property Let Caption(ByRef newCaption As String)
    ucSupport.SetCaptionText newCaption
    PropertyChanged "Caption"
End Property

'The Enabled property is a bit unique; see http://msdn.microsoft.com/en-us/library/aa261357%28v=vs.60%29.aspx
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_UserMemId = -514
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal newValue As Boolean)
    UserControl.Enabled = newValue
    RedrawBackBuffer
    PropertyChanged "Enabled"
End Property

Public Property Get FontSize() As Single
    FontSize = ucSupport.GetCaptionFontSize()
End Property

Public Property Let FontSize(ByVal newSize As Single)
    ucSupport.SetCaptionFontSize newSize
    PropertyChanged "FontSize"
End Property

'You can retrieve the gradient param string (not a pd2DGradient object!) via this property
Public Property Get Gradient() As String
    Gradient = m_curGradient
End Property

Public Property Let Gradient(ByVal NewGradient As String)
    m_curGradient = NewGradient
    RedrawBackBuffer
    RaiseEvent GradientChanged
    PropertyChanged "Gradient"
End Property

Public Property Get hWnd() As Long
Attribute hWnd.VB_UserMemId = -515
    hWnd = UserControl.hWnd
End Property

'Outside functions can call this to force a display of the gradient selection window
Public Sub DisplayGradientSelection()
    RaiseGradientDialog
End Sub

Private Sub ucSupport_ClickCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    UpdateMousePosition x, y
    If m_MouseInsideGradientRect Then RaiseGradientDialog
End Sub

Private Sub ucSupport_CustomMessage(ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturn As Long)
    If (wMsg = WM_PD_COLOR_MANAGEMENT_CHANGE) Then RedrawBackBuffer
End Sub

Private Sub ucSupport_KeyDownSystem(ByVal Shift As ShiftConstants, ByVal whichSysKey As PD_NavigationKey, markEventHandled As Boolean)
    
    'Enter/Esc get reported directly to the system key handler.  Note that we track the return, because TRUE
    ' means the key was successfully forwarded to the relevant handler.  (If FALSE is returned, no control
    ' accepted the keypress, meaning we should forward the event down the line.)
    markEventHandled = NavKey.NotifyNavKeypress(Me, whichSysKey, Shift)
    
End Sub

Private Sub ucSupport_MouseDownCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal timeStamp As Long)
    UpdateMousePosition x, y
    If m_MouseInsideGradientRect Then
        m_MouseDownGradientRect = True
        RedrawBackBuffer
    End If
End Sub

Private Sub ucSupport_MouseEnter(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    UpdateMousePosition x, y
    RedrawBackBuffer
End Sub

Private Sub ucSupport_MouseLeave(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    UpdateMousePosition -100, -100
    RedrawBackBuffer
End Sub

Private Sub ucSupport_MouseMoveCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal timeStamp As Long)
    UpdateMousePosition x, y
    RedrawBackBuffer
End Sub

Private Sub ucSupport_MouseUpCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal clickEventAlsoFiring As Boolean, ByVal timeStamp As Long)
    m_MouseDownGradientRect = False
    RedrawBackBuffer
End Sub

Private Sub UpdateMousePosition(ByVal mouseX As Single, ByVal mouseY As Single)
    m_MouseInsideGradientRect = PDMath.IsPointInRectF(mouseX, mouseY, m_GradientRect)
    If m_MouseInsideGradientRect Then ucSupport.RequestCursor IDC_HAND Else ucSupport.RequestCursor IDC_DEFAULT
End Sub

Private Sub ucSupport_GotFocusAPI()
    RaiseEvent GotFocusAPI
End Sub

Private Sub ucSupport_LostFocusAPI()
    RaiseEvent LostFocusAPI
End Sub

Private Sub ucSupport_RepaintRequired(ByVal updateLayoutToo As Boolean)
    If updateLayoutToo Then UpdateControlLayout
    RedrawBackBuffer
End Sub

Private Sub RaiseGradientDialog()

    isDialogLive = True
    
    'Backup the current gradient; if the dialog is canceled, we want to restore it
    Dim NewGradient As String, oldGradient As String
    oldGradient = Gradient
    
    'Display the gradient dialog, then wait for it to return
    If ShowGradientDialog(NewGradient, oldGradient, Me) Then
        Gradient = NewGradient
    Else
        Gradient = oldGradient
    End If
    
    isDialogLive = False
    
End Sub

Private Sub UserControl_Initialize()
    
    Set m_Brush = New pd2DBrush
    m_Brush.SetBrushMode P2_BM_Gradient
    
    'Initialize a master user control support class
    Set ucSupport = New pdUCSupport
    ucSupport.RegisterControl UserControl.hWnd, True
    ucSupport.RequestExtraFunctionality True
    ucSupport.RequestCaptionSupport
    ucSupport.SubclassCustomMessage WM_PD_COLOR_MANAGEMENT_CHANGE, True
    
    'Prep the color manager and load default colors
    Set m_Colors = New pdThemeColors
    Dim colorCount As PDGS_COLOR_LIST: colorCount = [_Count]
    m_Colors.InitializeColorList "PDGradientSelector", colorCount
    If Not MainModule.IsProgramRunning() Then UpdateColorList
    
    'Update the control size parameters at least once
    UpdateControlLayout
    
End Sub

Private Sub UserControl_InitProperties()
    Caption = vbNullString
    FontSize = 12
    Gradient = vbNullString
End Sub

'At run-time, painting is handled by the support class.  In the IDE, however, we must rely on VB's internal paint event.
Private Sub UserControl_Paint()
    ucSupport.RequestIDERepaint UserControl.hDC
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        Caption = .ReadProperty("Caption", vbNullString)
        Gradient = .ReadProperty("curGradient", vbNullString)
        FontSize = .ReadProperty("FontSize", 12)
    End With
End Sub

Private Sub UserControl_Resize()
    If Not MainModule.IsProgramRunning() Then ucSupport.RequestRepaint True
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "Caption", ucSupport.GetCaptionText, vbNullString
        .WriteProperty "FontSize", ucSupport.GetCaptionFontSize, 12
        .WriteProperty "curGradient", m_curGradient, vbNullString
    End With
End Sub

'Whenever a control property changes that affects control size or layout (including internal changes, like caption adjustments),
' call this function to recalculate the control's internal layout
Private Sub UpdateControlLayout()
    
    'Retrieve DPI-aware control dimensions from the support class
    Dim bWidth As Long, bHeight As Long
    bWidth = ucSupport.GetBackBufferWidth
    bHeight = ucSupport.GetBackBufferHeight
    
    'Next, determine the positioning of the caption, if present.  (ucSupport.GetCaptionBottom tells us where the
    ' caption text ends vertically.)
    If ucSupport.IsCaptionActive Then
        
        'The clickable area is placed relative to the caption
        With m_GradientRect
            .Left = FixDPI(8)
            .Top = ucSupport.GetCaptionBottom + 2
            .Width = (bWidth - 2) - .Left
            .Height = (bHeight - 2) - .Top
        End With
        
    'If there's no caption, allow the clickable portion to fill the entire control
    Else
        
        With m_GradientRect
            .Left = 1
            .Top = 1
            .Width = (bWidth - 2) - .Left
            .Height = (bHeight - 2) - .Top
        End With
        
    End If
            
End Sub

'Primary rendering function.  Note that ucSupport handles a number of rendering duties (like maintaining a back buffer for us).
Private Sub RedrawBackBuffer()
        
    'Request the back buffer DC, and ask the support module to erase any existing rendering for us.
    Dim bufferDC As Long
    bufferDC = ucSupport.GetBackBufferDC(True)
    
    'NOTE: if a caption exists, it has already been drawn.  We just need to draw the clickable brush portion.
    If MainModule.IsProgramRunning() Then
    
        'Render the gradient first.  To do this, we use a temporary gradient object that *only* renders as
        ' a linear gradient.  (This gradient control only allows editing of a gradient's color, not its shape.)
        Dim tmpGradient As pd2DGradient
        Set tmpGradient = New pd2DGradient
        tmpGradient.CreateGradientFromString m_curGradient
        tmpGradient.SetGradientShape P2_GS_Linear
        tmpGradient.SetGradientAngle 0#
        
        m_Brush.SetBoundaryRect m_GradientRect
        m_Brush.SetBrushGradientAllSettings tmpGradient.GetGradientAsString
        
        Dim tmpBrush As Long
        tmpBrush = m_Brush.GetHandle
        
        With m_GradientRect
            GDI_Plus.GDIPlusFillPatternToDC bufferDC, .Left, .Top, .Width, .Height, g_CheckerboardPattern
            GDI_Plus.GDIPlusFillDC_Brush bufferDC, tmpBrush, .Left, .Top, .Width, .Height
        End With
        
        m_Brush.ReleaseBrush
        
        'Before drawing borders around the brush results, ask our parent control to apply color-management to
        ' the brush preview.  (Note that this *will* result in the background checkerboard being color-managed.
        ' This isn't ideal, but we'll live with it for now as the alternative is messy.)
        ucSupport.RequestBufferColorManagement VarPtr(m_GradientRect)
        
        'Draw borders around the brush results.
        Dim outlineColor As Long, outlineWidth As Long
        outlineColor = m_Colors.RetrieveColor(PDGS_Border, Me.Enabled, m_MouseDownGradientRect, m_MouseInsideGradientRect)
        If m_MouseInsideGradientRect Then outlineWidth = 3 Else outlineWidth = 1
        GDI_Plus.GDIPlusDrawRectFOutlineToDC bufferDC, m_GradientRect, outlineColor, , outlineWidth, False, GP_LJ_Miter
       
    End If
    
    'Paint the final result to the screen, as relevant
    ucSupport.RequestRepaint
    
End Sub

'If a gradient selection dialog is active, it will pass gradient updates backward to this function, so that we can let
' our parent form display live updates *while the user is playing with gradients*.
Public Sub NotifyOfLiveGradientChange(ByVal NewGradient As String)
    Gradient = NewGradient
End Sub

'Before this control does any painting, we need to retrieve relevant colors from PD's primary theming class.  Note that this
' step must also be called if/when PD's visual theme settings change.
Private Sub UpdateColorList()
    m_Colors.LoadThemeColor PDGS_Border, "Border", IDE_BLACK
End Sub

'External functions can call this to request a redraw.  This is helpful for live-updating theme settings, as in the Preferences dialog.
Public Sub UpdateAgainstCurrentTheme(Optional ByVal hostFormhWnd As Long = 0)
    If ucSupport.ThemeUpdateRequired Then
        UpdateColorList
        If MainModule.IsProgramRunning() Then NavKey.NotifyControlLoad Me, hostFormhWnd
        If MainModule.IsProgramRunning() Then ucSupport.UpdateAgainstThemeAndLanguage
    End If
End Sub

'By design, PD prefers to not use design-time tooltips.  Apply tooltips at run-time, using this function.
' (IMPORTANT NOTE: translations are handled automatically.  Always pass the original English text!)
Public Sub AssignTooltip(ByRef newTooltip As String, Optional ByRef newTooltipTitle As String = vbNullString, Optional ByVal newTooltipIcon As TT_ICON_TYPE = TTI_NONE)
    ucSupport.AssignTooltip UserControl.ContainerHwnd, newTooltip, newTooltipTitle, newTooltipIcon
End Sub
