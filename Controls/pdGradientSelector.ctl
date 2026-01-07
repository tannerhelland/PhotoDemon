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
'Copyright 2015-2026 by Tanner Helland
'Created: 23/July/15
'Last updated: 04/December/20
'Last update: migrate renderer to pd2D
'
'This thin user control is basically an empty control that when clicked, displays a gradient editor window.
' If a gradient is selected (e.g. Cancel is not pressed), it updates its appearance to match, and raises a
' "GradientChanged" event.
'
'Though simple, this control solves a lot of problems.  It is especially helpful for improving interaction
' with the command bar user control, as it easily supports gradient reset/randomize/preset events.  It is
' also nice to update a single central function for gradient selection, then have the change propagate to
' all tool instances.
'
'The actual gradient functionality of the control comes from the pd2dGradient class - look there for details.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************


Option Explicit

'This control doesn't really do anything interesting, besides allow a gradient to be selected.
Public Event GradientChanged()

'Because VB focus events are wonky, especially when we use CreateWindow within a UC, this control raises its own
' specialized focus events.  If you need to track focus, use these instead of the default VB functions.
Public Event GotFocusAPI()
Public Event LostFocusAPI()
Public Event SetCustomTabTarget(ByVal shiftTabWasPressed As Boolean, ByRef newTargetHwnd As Long)

'The control's current gradient settings
Private m_curGradient As String

'Temporary brush object, used to render the gradient preview
Private m_Brush As pd2DBrush

'When the "select gradient" dialog is live, this will be set to TRUE
Private isDialogLive As Boolean

'The rectangle where the gradient preview is actually rendered, and a boolean to track whether the mouse is inside that rect
Private m_GradientRect As RectF, m_MouseInsideGradientRect As Boolean, m_MouseDownGradientRect As Boolean

'A secondary rect with a clickable button; the user can click this button to reverse the current gradient
Private m_ReverseRect As RectF, m_MouseInsideReverseRect As Boolean, m_MouseDownReverseRect As Boolean

'User control support class.  Historically, many classes (and associated subclassers) were required by each user control,
' but I've since wrapped these into a single central support class.
Private WithEvents ucSupport As pdUCSupport
Attribute ucSupport.VB_VarHelpID = -1

'Local list of themable colors.  This list includes all potential colors used by this class, regardless of state change
' or internal control settings.  The list is updated by calling the UpdateColorList function.
' (Note also that this list does not include variants, e.g. "BorderColor" vs "BorderColor_Hovered".  Variant values are
'  automatically calculated by the color management class, and they are retrieved by passing boolean modifiers to that
'  class, rather than treating every imaginable variant as a separate constant.)
Private Enum PDGS_COLOR_LIST
    [_First] = 0
    PDGS_Arrows = 0
    PDGS_Border = 1
    PDGS_ButtonFill = 2
    [_Last] = 2
    [_Count] = 3
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

Public Property Let Gradient(ByRef NewGradient As String)
    m_curGradient = NewGradient
    RedrawBackBuffer
    RaiseEvent GradientChanged
    PropertyChanged "Gradient"
End Property

Public Property Get hWnd() As Long
Attribute hWnd.VB_UserMemId = -515
    hWnd = UserControl.hWnd
End Property

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

'Outside functions can call this to force a display of the gradient selection window
Public Sub DisplayGradientSelection()
    RaiseGradientDialog
End Sub

Private Sub ucSupport_ClickCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    
    UpdateMousePosition x, y
    
    If m_MouseInsideGradientRect Then
        RaiseGradientDialog
    ElseIf m_MouseInsideReverseRect Then
        
        'Create a temporary gradient object, use it to reverse the current gradient,
        ' then redraw the control and notify any parent object(s)
        Dim tmpGradient As pd2DGradient
        Set tmpGradient = New pd2DGradient
        tmpGradient.CreateGradientFromString m_curGradient
        tmpGradient.ReverseGradient
        m_curGradient = tmpGradient.GetGradientAsString()
        
        RedrawBackBuffer
        RaiseEvent GradientChanged
        
    End If
    
End Sub

Private Sub ucSupport_CustomMessage(ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturn As Long)
    If (wMsg = WM_PD_COLOR_MANAGEMENT_CHANGE) Then RedrawBackBuffer
End Sub

Private Sub ucSupport_KeyDownCustom(ByVal Shift As ShiftConstants, ByVal vkCode As Long, markEventHandled As Boolean)

    markEventHandled = False
    
    If Me.Enabled And (vkCode = VK_SPACE) Then
        RaiseGradientDialog
        markEventHandled = True
    End If
    
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
    ElseIf m_MouseInsideReverseRect Then
        m_MouseDownReverseRect = True
    End If
    If m_MouseInsideGradientRect Or m_MouseInsideReverseRect Then RedrawBackBuffer
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
    m_MouseDownReverseRect = False
    RedrawBackBuffer
End Sub

Private Sub UpdateMousePosition(ByVal mouseX As Single, ByVal mouseY As Single)
    m_MouseInsideGradientRect = PDMath.IsPointInRectF(mouseX, mouseY, m_GradientRect)
    m_MouseInsideReverseRect = PDMath.IsPointInRectF(mouseX, mouseY, m_ReverseRect)
    If m_MouseInsideGradientRect Or m_MouseInsideReverseRect Then ucSupport.RequestCursor IDC_HAND Else ucSupport.RequestCursor IDC_DEFAULT
End Sub

Private Sub ucSupport_GotFocusAPI()
    RedrawBackBuffer
    RaiseEvent GotFocusAPI
End Sub

Private Sub ucSupport_LostFocusAPI()
    RedrawBackBuffer
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

Private Sub ucSupport_SetCustomTabTarget(ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    RaiseEvent SetCustomTabTarget(shiftTabWasPressed, newTargetHwnd)
End Sub

Private Sub UserControl_Initialize()
    
    Set m_Brush = New pd2DBrush
    m_Brush.SetBrushMode P2_BM_Gradient
    
    'Initialize a user control support class
    Set ucSupport = New pdUCSupport
    ucSupport.RegisterControl UserControl.hWnd, True
    ucSupport.RequestExtraFunctionality True, True
    ucSupport.SpecifyRequiredKeys VK_SPACE
    ucSupport.RequestCaptionSupport
    ucSupport.SubclassCustomMessage WM_PD_COLOR_MANAGEMENT_CHANGE, True
    
    'Prep the color manager and load default colors
    Set m_Colors = New pdThemeColors
    Dim colorCount As PDGS_COLOR_LIST: colorCount = [_Count]
    m_Colors.InitializeColorList "PDGradientSelector", colorCount
    If Not PDMain.IsProgramRunning() Then UpdateColorList
    
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
        m_curGradient = .ReadProperty("curGradient", vbNullString)
        FontSize = .ReadProperty("FontSize", 12)
    End With
End Sub

Private Sub UserControl_Resize()
    If Not PDMain.IsProgramRunning() Then ucSupport.RequestRepaint True
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
        With m_ReverseRect
            .Top = ucSupport.GetCaptionBottom + 2
            .Height = (bHeight - 2) - .Top
            .Width = Interface.FixDPI(24)
            .Left = (bWidth - 2) - .Width
        End With
        
        With m_GradientRect
            .Left = Interface.FixDPI(8)
            .Top = ucSupport.GetCaptionBottom + 2
            .Width = m_ReverseRect.Left - .Left
            .Height = (bHeight - 2) - .Top
        End With
        
    'If there's no caption, allow the clickable portion to fill the entire control
    Else
        
        With m_ReverseRect
            .Top = 1
            .Height = (bHeight - 2) - .Top
            .Width = Interface.FixDPI(24)
            .Left = (bWidth - 2) - .Width
        End With
        
        With m_GradientRect
            .Left = 1
            .Top = 1
            .Width = m_ReverseRect.Left - .Left
            .Height = (bHeight - 2) - .Top
        End With
        
    End If
            
End Sub

'Primary rendering function.  Note that ucSupport handles a number of rendering duties (like maintaining a back buffer for us).
Private Sub RedrawBackBuffer()
        
    'Request the back buffer DC, and ask the support module to erase any existing rendering for us.
    Dim bufferDC As Long
    bufferDC = ucSupport.GetBackBufferDC(True)
    If (bufferDC = 0) Then Exit Sub
    
    'NOTE: if a caption exists, it has already been drawn.  We just need to draw the clickable brush portion.
    If PDMain.IsProgramRunning() Then
    
        'Render the gradient first.  To do this, we use a temporary gradient object that *only* renders as
        ' a linear gradient.  (This gradient control only allows editing of a gradient's color, not its shape.)
        Dim tmpGradient As pd2DGradient
        Set tmpGradient = New pd2DGradient
        tmpGradient.CreateGradientFromString m_curGradient
        tmpGradient.SetGradientShape P2_GS_Linear
        tmpGradient.SetGradientAngle 0#
        
        m_Brush.SetBoundaryRect m_GradientRect
        m_Brush.SetBrushGradientAllSettings tmpGradient.GetGradientAsString
        
        'Start by using the universal checkerboard pattern object to paint a checkerboard background
        ' (in case the gradient contains low-opacity stops), followed by the gradient itself
        Dim cSurface As pd2DSurface
        Set cSurface = New pd2DSurface
        cSurface.WrapSurfaceAroundDC bufferDC
        cSurface.SetSurfaceAntialiasing P2_AA_None
        
        PD2D.FillRectangleF_FromRectF cSurface, g_CheckerboardBrush, m_GradientRect
        PD2D.FillRectangleF_FromRectF cSurface, m_Brush, m_GradientRect
        
        m_Brush.ReleaseBrush
        
        'Before drawing borders around the brush results, ask our parent control to apply color-management to
        ' the brush preview.  (Note that this *will* result in the background checkerboard being color-managed.
        ' This isn't ideal, but we'll live with it for now as the alternative is messy.)
        ucSupport.RequestBufferColorManagement VarPtr(m_GradientRect)
        
        'Fill the "reverse" button with the proper color (depending on its mouseover status)
        Dim cBrush As pd2DBrush
        Set cBrush = New pd2DBrush
        cBrush.SetBrushColor m_Colors.RetrieveColor(PDGS_ButtonFill, Me.Enabled, m_MouseDownReverseRect, m_MouseInsideReverseRect)
        PD2D.FillRectangleF_FromRectF cSurface, cBrush, m_ReverseRect
        
        'Always start by drawing inactive borders around both "buttons" on the control
        Dim cPen As pd2DPen
        Set cPen = New pd2DPen
        cPen.SetPenColor m_Colors.RetrieveColor(PDGS_Border, Me.Enabled, ucSupport.DoIHaveFocus, False)
        cPen.SetPenWidth 1!
        cPen.SetPenLineJoin P2_LJ_Miter
        PD2D.DrawRectangleF_FromRectF cSurface, cPen, m_GradientRect
        PD2D.DrawRectangleF_FromRectF cSurface, cPen, m_ReverseRect
        
        'Next, draw highlight borders around either button if they are actively mouse-hovered
        ' (or if the control has keyboard focus)
        cPen.SetPenColor m_Colors.RetrieveColor(PDGS_Border, Me.Enabled, False, True)
        cPen.SetPenWidth 3!
        
        If m_MouseInsideGradientRect Then
            PD2D.DrawRectangleF_FromRectF cSurface, cPen, m_GradientRect
        ElseIf m_MouseInsideReverseRect Then
            PD2D.DrawRectangleF_FromRectF cSurface, cPen, m_ReverseRect
        ElseIf ucSupport.DoIHaveFocus Then
            PD2D.DrawRectangleF_FromRectF cSurface, cPen, m_GradientRect
        End If
        
        'Finally, activate antialiasing and draw arrows to indicate the "reverse" purpose of the button.
        cSurface.SetSurfaceAntialiasing P2_AA_HighQuality
        cSurface.SetSurfacePixelOffset P2_PO_Half
        
        cPen.SetPenWidth 1!
        cPen.SetPenEndCap P2_LC_ArrowAnchor
        cPen.SetPenColor m_Colors.RetrieveColor(PDGS_Border, Me.Enabled, m_MouseDownReverseRect, m_MouseInsideReverseRect)
        
        Dim arrFirstPoint As PointFloat, arrSecondPoint As PointFloat
        arrFirstPoint.x = m_ReverseRect.Left + (m_ReverseRect.Width * 0.225!)
        arrSecondPoint.x = m_ReverseRect.Left + (m_ReverseRect.Width * 0.8!)
        arrFirstPoint.y = m_ReverseRect.Top + (m_ReverseRect.Height \ 2) - 4! + 0.5!
        arrSecondPoint.y = arrFirstPoint.y
        
        PD2D.DrawLineF_FromPtF cSurface, cPen, arrFirstPoint, arrSecondPoint
        arrFirstPoint.x = m_ReverseRect.Left + (m_ReverseRect.Width * 0.2!)
        arrSecondPoint.x = m_ReverseRect.Left + (m_ReverseRect.Width * 0.775!)
        arrFirstPoint.y = m_ReverseRect.Top + (m_ReverseRect.Height \ 2) + 4! + 1.5!
        arrSecondPoint.y = arrFirstPoint.y
        PD2D.DrawLineF_FromPtF cSurface, cPen, arrSecondPoint, arrFirstPoint
        
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
    With m_Colors
        .LoadThemeColor PDGS_Arrows, "Arrows", IDE_BLACK
        .LoadThemeColor PDGS_Border, "Border", IDE_BLACK
        .LoadThemeColor PDGS_ButtonFill, "ButtonFill", IDE_WHITE
    End With
End Sub

'External functions can call this to request a redraw.  This is helpful for live-updating theme settings, as in the Preferences dialog.
Public Sub UpdateAgainstCurrentTheme(Optional ByVal hostFormhWnd As Long = 0)
    If ucSupport.ThemeUpdateRequired Then
        UpdateColorList
        If PDMain.IsProgramRunning() Then NavKey.NotifyControlLoad Me, hostFormhWnd
        If PDMain.IsProgramRunning() Then ucSupport.UpdateAgainstThemeAndLanguage
    End If
End Sub

'By design, PD prefers to not use design-time tooltips.  Apply tooltips at run-time, using this function.
' (IMPORTANT NOTE: translations are handled automatically.  Always pass the original English text!)
Public Sub AssignTooltip(ByRef newTooltip As String, Optional ByRef newTooltipTitle As String = vbNullString, Optional ByVal raiseTipsImmediately As Boolean = False)
    ucSupport.AssignTooltip UserControl.ContainerHwnd, newTooltip, newTooltipTitle, raiseTipsImmediately
End Sub
