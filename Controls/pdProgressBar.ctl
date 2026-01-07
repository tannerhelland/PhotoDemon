VERSION 5.00
Begin VB.UserControl pdProgressBar 
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
   ToolboxBitmap   =   "pdProgressBar.ctx":0000
End
Attribute VB_Name = "pdProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Progress Bar UI element
'Copyright 2014-2026 by Tanner Helland
'Created: 20/March/18
'Last updated: 20/March/18
'Last update: initial build
'
'In a surprise to precisely no one, PhotoDemon has some unique needs when it comes to user controls - needs that
' the intrinsic VB controls can't handle.  These range from the obnoxious (lack of an "autosize" property for
' anything but labels) to the critical (no Unicode support).
'
'As such, I've created many of my own UCs for the program.  All are owner-drawn, with the goal of maintaining
' visual fidelity across the program, while also enabling key features like Unicode support.
'
'A few notes on this generic progress bar control, specifically:
'
' 1) High DPI settings are handled automatically.
' 2) Coloration is automatically handled by PD's internal theming engine.
' 3) Just like a system progress bar, a "marquee" mode is provided for tasks with unpredictable time requirements
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'User control support class.  Historically, many classes (and associated subclassers) were required by each user control,
' but I've since wrapped these into a single central support class.
Private WithEvents ucSupport As pdUCSupport
Attribute ucSupport.VB_VarHelpID = -1

'Local list of themable colors.  This list includes all potential colors used by the control, regardless of state change
' or internal control settings.  The list is updated by calling the UpdateColorList function.
' (Note also that this list does not include variants, e.g. "BorderColor" vs "BorderColor_Hovered".  Variant values are
'  automatically calculated by the color management class, and they are retrieved by passing boolean modifiers to that
'  class, rather than treating every imaginable variant as a separate constant.)
Private Enum PDPROGRESSBAR_COLOR_LIST
    [_First] = 0
    PDPB_Background = 0
    PDPB_Border = 1
    PDPB_MarqueeHighlight = 2
    PDPB_Progress = 3
    [_Last] = 3
    [_Count] = 4
End Enum

'Rects to define rendering regions for the progress bar's border vs the actual progress bar itself
Private m_BorderRect As RectF, m_ProgressRect As RectF

'Current progress bar max and value properties
Private m_ProgBarMax As Double, m_ProgBarValue As Double

'In marquee mode, the control handles all rendering internally; things like "prog bar max" don't matter
Private m_MarqueeMode As Boolean, m_MarqueeOffset As Double, m_LastTimeRendered As Currency
Private WithEvents m_MarqueeTimer As pdTimer
Attribute m_MarqueeTimer.VB_VarHelpID = -1

'Color retrieval and storage is handled by a dedicated class; this allows us to optimize theme interactions,
' without worrying about the details locally.
Private m_Colors As pdThemeColors

Public Function GetControlType() As PD_ControlType
    GetControlType = pdct_ProgressBar
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

'Progress-bar specific properties
Public Property Get MarqueeMode() As Boolean
    MarqueeMode = m_MarqueeMode
End Property

Public Property Let MarqueeMode(ByVal newMode As Boolean)

    If (m_MarqueeMode <> newMode) Then
    
        m_MarqueeMode = newMode
        PropertyChanged "MarqueeMode"
        
        'Start or release the marquee timer, as necessary
        If m_MarqueeMode And PDMain.IsProgramRunning Then
            Set m_MarqueeTimer = New pdTimer
            m_MarqueeTimer.Interval = 16
            m_MarqueeTimer.StartTimer
        Else
            If (Not m_MarqueeTimer Is Nothing) Then
                m_MarqueeTimer.StopTimer
                Set m_MarqueeTimer = Nothing
            End If
        End If
        
    End If
    
End Property

Public Property Get Max() As Double
    Max = m_ProgBarMax
End Property

Public Property Let Max(ByVal newValue As Double)
    If (m_ProgBarMax <> newValue) Then
        m_ProgBarMax = newValue
        RedrawBackBuffer
    End If
End Property

Public Property Get Value() As Double
    Value = m_ProgBarValue
End Property

Public Property Let Value(ByVal newValue As Double)
    If (m_ProgBarValue <> newValue) Then
        m_ProgBarValue = newValue
        RedrawBackBuffer
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

Private Sub m_MarqueeTimer_Timer()
    RedrawBackBuffer
End Sub

Private Sub ucSupport_RepaintRequired(ByVal updateLayoutToo As Boolean)
    If updateLayoutToo Then UpdateControlLayout Else RedrawBackBuffer
End Sub

Private Sub ucSupport_VisibilityChange(ByVal newVisibility As Boolean)
    If (Not newVisibility) And m_MarqueeMode Then
        If (Not m_MarqueeTimer Is Nothing) Then m_MarqueeTimer.StopTimer
        Set m_MarqueeTimer = Nothing
    End If
End Sub

Private Sub UserControl_Initialize()
    
    'Initialize a user control support class
    Set ucSupport = New pdUCSupport
    ucSupport.RegisterControl UserControl.hWnd, False
    
    'Prep the color manager and load default colors
    Set m_Colors = New pdThemeColors
    Dim colorCount As PDPROGRESSBAR_COLOR_LIST: colorCount = [_Count]
    m_Colors.InitializeColorList "PDProgressBar", colorCount
    If Not PDMain.IsProgramRunning() Then UpdateColorList
    
End Sub

Private Sub UserControl_InitProperties()
    Me.Enabled = True
    Me.MarqueeMode = False
End Sub

'At run-time, painting is handled by PD's pdWindowPainter class.  In the IDE, however, we must rely on VB's internal paint event.
Private Sub UserControl_Paint()
    If (Not PDMain.IsProgramRunning()) Then ucSupport.RequestIDERepaint UserControl.hDC
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        Me.Enabled = .ReadProperty("Enabled", True)
        Me.MarqueeMode = .ReadProperty("MarqueeMode", False)
    End With
End Sub

Private Sub UserControl_Resize()
    If (Not PDMain.IsProgramRunning()) Then ucSupport.NotifyIDEResize UserControl.Width, UserControl.Height
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "Enabled", Me.Enabled, True
        .WriteProperty "MarqueeMode", Me.MarqueeMode, False
    End With
End Sub

'Because this control automatically forces all internal buttons to identical sizes, we have to recalculate a number
' of internal sizing metrics whenever the control size changes.
Private Sub UpdateControlLayout()
    
    'Retrieve DPI-aware control dimensions from the support class
    Dim bWidth As Long, bHeight As Long
    bWidth = ucSupport.GetBackBufferWidth
    bHeight = ucSupport.GetBackBufferHeight
    
    'Cache rendering rects now; this improves redraw performance
    With m_BorderRect
        .Left = 0
        .Top = 0
        .Width = bWidth - 1
        .Height = bHeight - 1
    End With
    
    With m_ProgressRect
        .Left = 2!
        .Top = 2!
        .Width = bWidth - 4
        .Height = bHeight - 4
    End With
    
    'No other special preparation is required for this control, so proceed with recreating the back buffer
    RedrawBackBuffer
            
End Sub

'Use this function to completely redraw the back buffer from scratch.  Note that this is computationally expensive compared to just flipping the
' existing buffer to the screen, so only redraw the backbuffer if the control state has somehow changed.
Private Sub RedrawBackBuffer()
    
    'Request the back buffer DC, and ask the support module to erase any existing rendering for us.
    Dim bufferDC As Long
    bufferDC = ucSupport.GetBackBufferDC(True, m_Colors.RetrieveColor(PDPB_Background))
    If (bufferDC = 0) Then Exit Sub
    
    'Rendering is pretty easy - fill a fraction of the control with the current progress level!
    If PDMain.IsProgramRunning() And ucSupport.AmIVisible() Then
        
        Dim cSurface As pd2DSurface
        Drawing2D.QuickCreateSurfaceFromDC cSurface, bufferDC, False
        
        Dim cPen As pd2DPen
        Drawing2D.QuickCreateSolidPen cPen, 1!, m_Colors.RetrieveColor(PDPB_Border, Me.Enabled)
        PD2D.DrawRectangleF_FromRectF cSurface, cPen, m_BorderRect
        Set cPen = Nothing
        
        Dim cBrush As pd2DBrush
        
        'Marquee mode doesn't use the value or max properties; instead, all rendering is handled manually
        If m_MarqueeMode Then
            
            'Figure out how much time has elapsed since the last render, and use that to calculate a marquee offset
            Dim timeElapsed As Double
            If (m_LastTimeRendered <> 0) Then
                timeElapsed = VBHacks.GetTimerDifferenceNow(m_LastTimeRendered)
                Const MARQUEE_SPEED As Double = (8# / 0.016)
                m_MarqueeOffset = m_MarqueeOffset + timeElapsed * MARQUEE_SPEED
            Else
                m_MarqueeOffset = 0#
            End If
            
            'We now want to draw a gradient at the current marquee position
            
            'Start by figuring out the gradient's size.  (It's relative to the progress bar's width.)
            Const GRADIENT_RATIO As Double = 0.25
            Dim gradWidth As Single
            gradWidth = GRADIENT_RATIO * m_ProgressRect.Width
            
            'If the marquee extends past the end of the progress bar, reset it to its leftmost position
            If (m_MarqueeOffset > m_ProgressRect.Width + gradWidth) Then m_MarqueeOffset = (-1# * gradWidth)
            
            'Fill the progress bar with the default accent color
            Drawing2D.QuickCreateSolidBrush cBrush, m_Colors.RetrieveColor(PDPB_Progress, Me.Enabled)
            With m_ProgressRect
                PD2D.FillRectangleF cSurface, cBrush, .Left, .Top, .Width, .Height
            End With
            
            'Activate high-quality rendering
            cSurface.SetSurfaceAntialiasing P2_AA_HighQuality
            cSurface.SetSurfacePixelOffset P2_PO_Half
            
            'Construct a gradient brush
            Dim cGradient As pd2DGradient
            Set cGradient = New pd2DGradient
            
            Dim gradPoints(0 To 2) As GradientPoint
            gradPoints(0).PointPosition = 0!
            gradPoints(0).PointOpacity = 100!
            gradPoints(0).PointRGB = m_Colors.RetrieveColor(PDPB_Progress)
            
            gradPoints(1).PointPosition = 0.5!
            gradPoints(1).PointOpacity = 100!
            gradPoints(1).PointRGB = m_Colors.RetrieveColor(PDPB_MarqueeHighlight)
            
            gradPoints(2).PointPosition = 1!
            gradPoints(2).PointOpacity = 100!
            gradPoints(2).PointRGB = m_Colors.RetrieveColor(PDPB_Progress)
            
            cGradient.CreateGradientFromPointCollection 3, gradPoints
            cGradient.SetGradientAngle 0
            cGradient.SetGradientShape P2_GS_Linear
            
            Set cBrush = New pd2DBrush
            cBrush.SetBrushMode P2_BM_Gradient
            cBrush.SetBrushGradientAllSettings cGradient.GetGradientAsString()
            
            'Determine a boundary rect for the gradient region
            Dim boundsRect As RectF
            boundsRect.Left = m_MarqueeOffset - gradWidth
            boundsRect.Width = gradWidth * 2
            boundsRect.Top = m_ProgressRect.Top
            boundsRect.Height = m_ProgressRect.Height
            cBrush.SetBoundaryRect boundsRect
            
            'Activate clipping to ensure we don't render outside the progress bar borders
            cSurface.SetSurfaceClip_FromRectF m_ProgressRect
            
            'Render the gradient!
            PD2D.FillRectangleF_FromRectF cSurface, cBrush, boundsRect
            Set cBrush = Nothing: Set cGradient = Nothing
            
            'Before we exit, make a note of the current time; we'll use this on subsequent animations to
            ' achieve perfect subpixel offsets.
            VBHacks.GetHighResTime m_LastTimeRendered
            
        Else
        
            If (m_ProgBarMax <> 0#) Then
                
                'Use subpixel positioning for the progress bar.  (For extremely slow-moving bars,
                ' this provides improved visual feedback.)
                cSurface.SetSurfaceAntialiasing P2_AA_HighQuality
                cSurface.SetSurfacePixelOffset P2_PO_Half
                
                Drawing2D.QuickCreateSolidBrush cBrush, m_Colors.RetrieveColor(PDPB_Progress, Me.Enabled)
                
                Dim progBarWidth As Single
                progBarWidth = (m_ProgBarValue / m_ProgBarMax) * m_ProgressRect.Width
                
                With m_ProgressRect
                    PD2D.FillRectangleF cSurface, cBrush, .Left, .Top, progBarWidth, .Height
                End With
                
                Set cBrush = Nothing
                
            End If
            
        End If
        
        Set cSurface = Nothing
        
    End If
    
    'Paint the final result to the screen, as relevant
    ucSupport.RequestRepaint True
    If (Not PDMain.IsProgramRunning()) Then UserControl.Refresh
    
End Sub

'Before this control does any painting, we need to retrieve relevant colors from PD's primary theming class.  Note that this
' step must also be called if/when PD's visual theme settings change.
Private Sub UpdateColorList()
    With m_Colors
        .LoadThemeColor PDPB_Background, "Background", IDE_WHITE
        .LoadThemeColor PDPB_Border, "Border", IDE_BLACK
        .LoadThemeColor PDPB_MarqueeHighlight, "MarqueeHighlight", IDE_WHITE
        .LoadThemeColor PDPB_Progress, "Progress", IDE_BLUE
    End With
End Sub

'External functions can call this to request a redraw.  This is helpful for live-updating theme settings, as in the Preferences dialog.
Public Sub UpdateAgainstCurrentTheme(Optional ByVal hostFormhWnd As Long = 0)
    If ucSupport.ThemeUpdateRequired Then
        UpdateColorList
        If PDMain.IsProgramRunning() Then NavKey.NotifyControlLoad Me, hostFormhWnd, False
        If PDMain.IsProgramRunning() Then ucSupport.UpdateAgainstThemeAndLanguage
    End If
End Sub

'By design, PD prefers to not use design-time tooltips.  Apply tooltips at run-time, using this function.
' (IMPORTANT NOTE: translations are handled automatically.  Always pass the original English text!)
Public Sub AssignTooltip(ByRef newTooltip As String, Optional ByRef newTooltipTitle As String = vbNullString, Optional ByVal raiseTipsImmediately As Boolean = False)
    ucSupport.AssignTooltip UserControl.ContainerHwnd, newTooltip, newTooltipTitle, raiseTipsImmediately
End Sub
