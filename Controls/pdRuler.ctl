VERSION 5.00
Begin VB.UserControl pdRuler 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ClipBehavior    =   0  'None
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
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "pdRuler.ctx":0000
End
Attribute VB_Name = "pdRuler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Viewport Ruler UI element
'Copyright 2018-2018 by Tanner Helland
'Created: 03/April/18
'Last updated: 03/April/18
'Last update: initial build
'
'At present, this control is only designed for use on PD's primary canvas.  A few things to note:
'
' 1) High DPI settings are handled automatically.
' 2) Coloration is automatically handled by PD's internal theming engine.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'User control support class.  Historically, many classes (and associated subclassers) were required by each user control,
' but I've since attempted to wrap these into a single master control support class.
Private WithEvents ucSupport As pdUCSupport
Attribute ucSupport.VB_VarHelpID = -1

'Rulers can be horizontal or vertical, obviously
Private m_Orientation As PD_Orientation

'Local list of themable colors.  This list includes all potential colors used by the control, regardless of state change
' or internal control settings.  The list is updated by calling the UpdateColorList function.
' (Note also that this list does not include variants, e.g. "BorderColor" vs "BorderColor_Hovered".  Variant values are
'  automatically calculated by the color management class, and they are retrieved by passing boolean modifiers to that
'  class, rather than treating every imaginable variant as a separate constant.)
Private Enum PDRULER_COLOR_LIST
    [_First] = 0
    PDR_Background = 0
    PDR_Text = 1
    PDR_Notch = 2
    PDR_Mouse = 3
    [_Last] = 3
    [_Count] = 4
End Enum

'Color retrieval and storage is handled by a dedicated class; this allows us to optimize theme interactions,
' without worrying about the details locally.
Private m_Colors As pdThemeColors

Public Function GetControlType() As PD_ControlType
    GetControlType = pdct_Ruler
End Function

Public Function GetControlName() As String
    GetControlName = UserControl.Extender.Name
End Function

'The Enabled property is a bit unique; see http://msdn.microsoft.com/en-us/library/aa261357%28v=vs.60%29.aspx
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_UserMemId = -514
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal newValue As Boolean)
    UserControl.Enabled = newValue
    PropertyChanged "Enabled"
    RedrawBackBuffer
End Property

Public Property Get Orientation() As PD_Orientation
    Orientation = m_Orientation
End Property

Public Property Let Orientation(ByVal newOrientation As PD_Orientation)
    If (newOrientation <> m_Orientation) Then
        m_Orientation = newOrientation
        PropertyChanged "Orientation"
    End If
End Property

'hWnds aren't exposed by default
Public Property Get hWnd() As Long
Attribute hWnd.VB_UserMemId = -515
    hWnd = UserControl.hWnd
End Property

'Container hWnd must be exposed for external tooltip handling
Public Property Get ContainerHwnd() As Long
    ContainerHwnd = UserControl.ContainerHwnd
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

Private Sub ucSupport_RepaintRequired(ByVal updateLayoutToo As Boolean)
    If updateLayoutToo Then UpdateControlLayout Else RedrawBackBuffer
End Sub

Private Sub UserControl_Initialize()
    
    'Initialize a master user control support class
    Set ucSupport = New pdUCSupport
    ucSupport.RegisterControl UserControl.hWnd, False
    
    'Prep the color manager and load default colors
    Set m_Colors = New pdThemeColors
    Dim colorCount As PDRULER_COLOR_LIST: colorCount = [_Count]
    m_Colors.InitializeColorList "PDRuler", colorCount
    If Not pdMain.IsProgramRunning() Then UpdateColorList
    
End Sub

Private Sub UserControl_InitProperties()
    Me.Enabled = True
    Me.Orientation = pdo_Horizontal
End Sub

'At run-time, painting is handled by PD's pdWindowPainter class.  In the IDE, however, we must rely on VB's internal paint event.
Private Sub UserControl_Paint()
    If (Not pdMain.IsProgramRunning()) Then ucSupport.RequestIDERepaint UserControl.hDC
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        Me.Enabled = .ReadProperty("Enabled", True)
        Me.Orientation = .ReadProperty("Orientation", pdo_Horizontal)
    End With
End Sub

Private Sub UserControl_Resize()
    If (Not pdMain.IsProgramRunning()) Then ucSupport.NotifyIDEResize UserControl.Width, UserControl.Height
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "Enabled", Me.Enabled, True
        .WriteProperty "Orientation", m_Orientation, pdo_Horizontal
    End With
End Sub

'Because this control automatically forces all internal buttons to identical sizes, we have to recalculate a number
' of internal sizing metrics whenever the control size changes.
Private Sub UpdateControlLayout()
    
    'Retrieve DPI-aware control dimensions from the support class
    Dim bWidth As Long, bHeight As Long
    bWidth = ucSupport.GetBackBufferWidth
    bHeight = ucSupport.GetBackBufferHeight
    
    'TODO
    
    'No other special preparation is required for this control, so proceed with recreating the back buffer
    RedrawBackBuffer
            
End Sub

'Use this function to completely redraw the back buffer from scratch.  Note that this is computationally expensive compared to just flipping the
' existing buffer to the screen, so only redraw the backbuffer if the control state has somehow changed.
Private Sub RedrawBackBuffer()
    
    'Request the back buffer DC, and ask the support module to erase any existing rendering for us.
    Dim bufferDC As Long
    bufferDC = ucSupport.GetBackBufferDC(True, m_Colors.RetrieveColor(PDR_Background))
    If (bufferDC = 0) Then Exit Sub
    
    Dim bWidth As Long, bHeight As Long
    bWidth = ucSupport.GetBackBufferWidth
    bHeight = ucSupport.GetBackBufferHeight
    
    'Rendering is pretty easy - fill a fraction of the control with the current progress level!
    If pdMain.IsProgramRunning() And ucSupport.AmIVisible() Then
        
        Dim cPainter As pd2DPainter
        Drawing2D.QuickCreatePainter cPainter
        
        Dim cSurface As pd2DSurface
        Drawing2D.QuickCreateSurfaceFromDC cSurface, bufferDC, False
        
        Dim cPen As pd2DPen
        Drawing2D.QuickCreateSolidPen cPen, 1, m_Colors.RetrieveColor(PDR_Mouse, Me.Enabled)
        
        Dim ctlRectF As RectF
        With ctlRectF
            .Top = 0!
            .Left = 0!
            .Width = bWidth - 1
            .Height = bHeight - 1
        End With
        cPainter.DrawRectangleF_FromRectF cSurface, cPen, ctlRectF
        
        Set cPen = Nothing
        
        Set cSurface = Nothing: Set cPainter = Nothing
        
    End If
    
    'Paint the final result to the screen, as relevant
    ucSupport.RequestRepaint True
    If (Not pdMain.IsProgramRunning()) Then UserControl.Refresh
    
End Sub

'Before this control does any painting, we need to retrieve relevant colors from PD's primary theming class.  Note that this
' step must also be called if/when PD's visual theme settings change.
Private Sub UpdateColorList()
    With m_Colors
        .LoadThemeColor PDR_Background, "Background", IDE_WHITE
        .LoadThemeColor PDR_Text, "Text", IDE_BLACK
        .LoadThemeColor PDR_Notch, "Notch", IDE_BLACK
        .LoadThemeColor PDR_Mouse, "Mouse", IDE_BLUE
    End With
End Sub

'External functions can call this to request a redraw.  This is helpful for live-updating theme settings, as in the Preferences dialog.
Public Sub UpdateAgainstCurrentTheme(Optional ByVal hostFormhWnd As Long = 0)
    If ucSupport.ThemeUpdateRequired Then
        UpdateColorList
        If pdMain.IsProgramRunning() Then ucSupport.UpdateAgainstThemeAndLanguage
    End If
End Sub

'By design, PD prefers to not use design-time tooltips.  Apply tooltips at run-time, using this function.
' (IMPORTANT NOTE: translations are handled automatically.  Always pass the original English text!)
Public Sub AssignTooltip(ByRef newTooltip As String, Optional ByRef newTooltipTitle As String = vbNullString, Optional ByVal newTooltipIcon As TT_ICON_TYPE = TTI_NONE, Optional ByVal raiseTipsImmediately As Boolean = False)
    ucSupport.AssignTooltip UserControl.ContainerHwnd, newTooltip, newTooltipTitle, newTooltipIcon, raiseTipsImmediately
End Sub
