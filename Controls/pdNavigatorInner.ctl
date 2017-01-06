VERSION 5.00
Begin VB.UserControl pdNavigatorInner 
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
   ToolboxBitmap   =   "pdNavigatorInner.ctx":0000
End
Attribute VB_Name = "pdNavigatorInner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Navigation custom control (inner panel)
'Copyright 2015-2017 by Tanner Helland
'Created: 16/October/15
'Last updated: 16/February/16
'Last update: migrate portions of the navigator control into this standalone inner panel; this frees us up to add
'             additional buttons and other features to the main pdNavigator control, without running into VB's
'             inherent focus issues.
'
'For implementation details, please refer to the main pdNavigator control.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'If the control is resized at run-time, it will request a new thumbnail via this function.  The passed DIB will already
' be sized to the
Public Event RequestUpdatedThumbnail(ByRef thumbDIB As pdDIB, ByRef thumbX As Single, ByRef thumbY As Single)

'When the user interacts with the navigation box, the (x, y) coordinates *in image space* will be returned in this event.
Public Event NewViewportLocation(ByVal imgX As Single, ByVal imgY As Single)

'Because VB focus events are wonky, especially when we use CreateWindow within a UC, this control raises its own
' specialized focus events.  If you need to track focus, use these instead of the default VB functions.
Public Event GotFocusAPI()
Public Event LostFocusAPI()

'The image thumbnail is cached independently, so we only request updates when absolutely necessary.
Private m_ImageThumbnail As pdDIB

'This value will be TRUE while the mouse is inside the navigator box
Private m_MouseInsideBox As Boolean

'Padding (in pixels) between the edges of the control and the image thumbnail.  This is automatically adjusted for
' DPI at run-time.
Private Const THUMB_PADDING As Long = 3

'When the control raises a request for a new thumbnail image, that function will supply an (optional?) (x, y) pair detailing
' where the thumb is centered within the navigator.  We use this to know where the image lies inside the thumb.
Private m_ThumbEventX As Single, m_ThumbEventY As Single

'The rect where the image thumbnail has been drawn.  This is calculated by the RedrawBackBuffer function.
Private m_ThumbRect As RECTF, m_ImageRegion As RECTF

'Last mouse (x, y) values.  We track these so we know whether to highlight the region box inside the navigator.
Private m_LastMouseX As Single, m_LastMouseY As Single

'User control support class.  Historically, many classes (and associated subclassers) were required by each user control,
' but I've since attempted to wrap these into a single master control support class.
Private WithEvents ucSupport As pdUCSupport
Attribute ucSupport.VB_VarHelpID = -1

'Local list of themable colors.  This list includes all potential colors used by this class, regardless of state change
' or internal control settings.  The list is updated by calling the UpdateColorList function.
' (Note also that this list does not include variants, e.g. "BorderColor" vs "BorderColor_Hovered".  Variant values are
'  automatically calculated by the color management class, and they are retrieved by passing boolean modifiers to that
'  class, rather than treating every imaginable variant as a separate constant.)
Private Enum PDNAVINNER_COLOR_LIST
    [_First] = 0
    PDNI_Background = 0
    [_Last] = 0
    [_Count] = 1
End Enum

'Color retrieval and storage is handled by a dedicated class; this allows us to optimize theme interactions,
' without worrying about the details locally.
Private m_Colors As pdThemeColors

'The Enabled property is a bit unique; see http://msdn.microsoft.com/en-us/library/aa261357%28v=vs.60%29.aspx
Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal newValue As Boolean)
    UserControl.Enabled = newValue
    RedrawBackBuffer
    PropertyChanged "Enabled"
End Property

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

Public Property Get ContainerHwnd() As Long
    ContainerHwnd = UserControl.ContainerHwnd
End Property

Private Sub ucSupport_CustomMessage(ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturn As Long)
    If (wMsg = WM_PD_COLOR_MANAGEMENT_CHANGE) Then Me.NotifyNewThumbNeeded
End Sub

Private Sub ucSupport_GotFocusAPI()
    RaiseEvent GotFocusAPI
End Sub

Private Sub ucSupport_LostFocusAPI()
    RaiseEvent LostFocusAPI
End Sub

'To support high-DPI settings properly, we expose some specialized move+size functions
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

'If the mouse button is clicked inside the image portion of the navigator, scroll to that (x, y) position
Private Sub ucSupport_MouseDownCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal timeStamp As Long)
    If (Button And pdLeftButton) <> 0 Then
        If Math_Functions.IsPointInRectF(x, y, m_ImageRegion) Then ScrollToXY x, y
    End If
End Sub

Private Sub ucSupport_MouseEnter(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    m_MouseInsideBox = True
End Sub

Private Sub ucSupport_MouseLeave(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    m_MouseInsideBox = False
    m_LastMouseX = -1: m_LastMouseY = -1
    ucSupport.RequestCursor IDC_DEFAULT
    RedrawBackBuffer
End Sub

Private Sub ucSupport_MouseMoveCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal timeStamp As Long)
    
    m_LastMouseX = x: m_LastMouseY = y
    
    'Set the cursor depending on whether the mouse is inside the image portion of the navigator control
    If Math_Functions.IsPointInRectF(x, y, m_ImageRegion) Then
        ucSupport.RequestCursor IDC_HAND
    Else
        ucSupport.RequestCursor IDC_DEFAULT
    End If
    
    'If the mouse button is down, scroll to that (x, y) position.  Note that we don't care if the cursor is in-bounds;
    ' the ScrollToXY function will automatically fix that for us.
    If (Button And pdLeftButton) <> 0 Then
        ScrollToXY x, y
    Else
        RedrawBackBuffer
    End If
    
End Sub

'Given an (x, y) coordinate in the navigator, scroll to the matching (x, y) in the image.
Private Sub ScrollToXY(ByVal x As Single, ByVal y As Single)

    'Make sure the image region has been successfully created, or this is all for naught
    If (g_OpenImageCount > 0) And (m_ImageRegion.Width <> 0) And (m_ImageRegion.Height <> 0) Then
    
        'Convert the (x, y) to the [0, 1] range
        Dim xRatio As Double, yRatio As Double
        xRatio = (x - m_ImageRegion.Left) / m_ImageRegion.Width
        yRatio = (y - m_ImageRegion.Top) / m_ImageRegion.Height
        If xRatio < 0 Then xRatio = 0: If xRatio > 1 Then xRatio = 1
        If yRatio < 0 Then yRatio = 0: If yRatio > 1 Then yRatio = 1
        
        'Next, convert those to the (min, max) scale of the current viewport scrollbars
        Dim hScrollRange As Double, vScrollRange As Double, newHScroll As Double, newVscroll As Double
        hScrollRange = FormMain.mainCanvas(0).GetScrollMax(PD_HORIZONTAL) - FormMain.mainCanvas(0).GetScrollMin(PD_HORIZONTAL)
        vScrollRange = FormMain.mainCanvas(0).GetScrollMax(PD_VERTICAL) - FormMain.mainCanvas(0).GetScrollMin(PD_VERTICAL)
        newHScroll = (xRatio * hScrollRange) + FormMain.mainCanvas(0).GetScrollMin(PD_HORIZONTAL)
        newVscroll = (yRatio * vScrollRange) + FormMain.mainCanvas(0).GetScrollMin(PD_VERTICAL)
        
        'Assign the new scrollbar values, then request a viewport refresh
        FormMain.mainCanvas(0).SetRedrawSuspension True
        FormMain.mainCanvas(0).SetScrollValue PD_HORIZONTAL, newHScroll
        FormMain.mainCanvas(0).SetScrollValue PD_VERTICAL, newVscroll
        FormMain.mainCanvas(0).SetRedrawSuspension False
        
        Viewport_Engine.Stage3_ExtractRelevantRegion pdImages(g_CurrentImage), FormMain.mainCanvas(0)
        
    End If

End Sub

Private Sub ucSupport_RepaintRequired(ByVal updateLayoutToo As Boolean)
    If updateLayoutToo Then UpdateControlLayout
    RedrawBackBuffer
End Sub

Private Sub ucSupport_WindowResize(ByVal newWidth As Long, ByVal newHeight As Long)
    UpdateControlLayout
End Sub

Private Sub UserControl_Initialize()
    
    'Initialize a master user control support class
    Set ucSupport = New pdUCSupport
    ucSupport.RegisterControl UserControl.hWnd
    ucSupport.RequestExtraFunctionality True
    ucSupport.SubclassCustomMessage WM_PD_COLOR_MANAGEMENT_CHANGE, True
    
    'Prep the color manager and load default colors
    Set m_Colors = New pdThemeColors
    Dim colorCount As PDNAVINNER_COLOR_LIST: colorCount = [_Count]
    m_Colors.InitializeColorList "PDNavInner", colorCount
    If Not g_IsProgramRunning Then UpdateColorList
    
    'Update the control size parameters at least once
    UpdateControlLayout
    
End Sub

'At run-time, painting is handled by the support class.  In the IDE, however, we must rely on VB's internal paint event.
Private Sub UserControl_Paint()
    ucSupport.RequestIDERepaint UserControl.hDC
End Sub

Private Sub UserControl_Resize()
    If Not g_IsProgramRunning Then ucSupport.RequestRepaint True
End Sub

'Call this to recreate all buffers against a changed control size.
Private Sub UpdateControlLayout()
    
    'Retrieve DPI-aware control dimensions from the support class
    Dim bWidth As Long, bHeight As Long
    bWidth = ucSupport.GetBackBufferWidth
    bHeight = ucSupport.GetBackBufferHeight
    
    'Whenever the navigator is resized, we must also resize the image thumbnail to match.
    
    'At present, we pad the thumbnail by a few pixels so we have room for a border.
    Dim thumbWidth As Long, thumbHeight As Long
    thumbWidth = bWidth - FixDPIFloat(THUMB_PADDING) * 2
    thumbHeight = bHeight - FixDPIFloat(THUMB_PADDING) * 2
    
    'Try to optimize re-creating the thumbnail, so we only do it when absolutely necessary
    If m_ImageThumbnail Is Nothing Then Set m_ImageThumbnail = New pdDIB
    If (m_ImageThumbnail.GetDIBWidth <> thumbWidth) Or (m_ImageThumbnail.GetDIBHeight <> thumbHeight) Then
        m_ImageThumbnail.CreateBlank thumbWidth, thumbHeight, 32, 0, 0
    Else
        m_ImageThumbnail.ResetDIB 0
    End If
    
    RaiseEvent RequestUpdatedThumbnail(m_ImageThumbnail, m_ThumbEventX, m_ThumbEventY)
    
    'With the backbuffer and image thumbnail successfully created, we can finally redraw the new navigator window
    RedrawBackBuffer
    
End Sub

'Need to redraw the navigator box?  Call this.  Note that it *does not* request a new image thumbnail.  You must handle
' that separately.  This simply uses whatever's been previously cached.
Private Sub RedrawBackBuffer()
    
    'We can improve shutdown performance by ignoring redraw requests
    If g_ProgramShuttingDown Then
        If (g_Themer Is Nothing) Then Exit Sub
    End If
    
    'Request the back buffer DC, and ask the support module to erase any existing rendering for us.
    Dim bufferDC As Long, bWidth As Long, bHeight As Long
    bufferDC = ucSupport.GetBackBufferDC(True, m_Colors.RetrieveColor(PDNI_Background, Me.Enabled))
    bWidth = ucSupport.GetBackBufferWidth
    bHeight = ucSupport.GetBackBufferHeight
    
    If g_IsProgramRunning Then
    
        'If an image has been loaded, determine a centered position for the image's thumbnail
        If g_OpenImageCount <= 0 Then
            With m_ThumbRect
                .Width = 0
                .Height = 0
                .Left = 0
                .Top = 0
            End With
        Else
            
            With m_ThumbRect
                .Width = m_ImageThumbnail.GetDIBWidth
                .Height = m_ImageThumbnail.GetDIBHeight
                .Left = (bWidth - m_ImageThumbnail.GetDIBWidth) / 2
                .Top = (bHeight - m_ImageThumbnail.GetDIBHeight) / 2
            End With
            
            'Offset that top-left corner by the thumbnail's position, and cache it to a module-level rect so we can use
            ' it for hit-detection during mouse events.
            With m_ImageRegion
                .Left = m_ThumbRect.Left + m_ThumbEventX
                .Top = m_ThumbRect.Top + m_ThumbEventY
                .Width = m_ImageThumbnail.GetDIBWidth - (m_ThumbEventX * 2)
                .Height = m_ImageThumbnail.GetDIBHeight - (m_ThumbEventY * 2)
            End With
            
            'Paint a checkerboard background only over the relevant image region
            With m_ImageRegion
                GDI_Plus.GDIPlusFillDIBRect_Pattern Nothing, .Left, .Top, .Width, .Height, g_CheckerboardPattern, bufferDC, True
            End With
            
            'Paint the thumb rect without regard for the image region (as it will always be a square)
            With m_ThumbRect
                GDI_Plus.GDIPlus_StretchBlt Nothing, .Left, .Top, .Width, .Height, m_ImageThumbnail, 0, 0, .Width, .Height, , GP_IM_HighQualityBicubic, bufferDC
                m_ImageThumbnail.FreeFromDC
            End With
                        
            'Query the active image for a copy of the intersection rect of the viewport, and the image itself,
            ' in image coordinate space
            Dim viewportRect As RECTF
            pdImages(g_CurrentImage).imgViewport.GetIntersectRectImage viewportRect
            
            'We now want to convert the viewport rect into our little navigator coordinate space.  Start by converting the
            ' viewport dimensions to a 1-based system, relative to the original image's width and height.
            If (pdImages(g_CurrentImage).Width > 0) And (pdImages(g_CurrentImage).Height > 0) Then
                
                Dim relativeRect As RECTF
                With relativeRect
                    .Left = viewportRect.Left / pdImages(g_CurrentImage).Width
                    .Top = viewportRect.Top / pdImages(g_CurrentImage).Height
                    .Width = viewportRect.Width / pdImages(g_CurrentImage).Width
                    .Height = viewportRect.Height / pdImages(g_CurrentImage).Height
                
                    'Next, scale those 1-based values by the navigator's current size
                    .Left = .Left * m_ImageRegion.Width
                    .Top = .Top * m_ImageRegion.Height
                    .Width = .Width * m_ImageRegion.Width
                    .Height = .Height * m_ImageRegion.Height
                    
                    'Finally, scale the values by the offsets of the image region
                    .Left = .Left + m_ImageRegion.Left
                    .Top = .Top + m_ImageRegion.Top
                End With
                
                'If the mouse is inside the control, figure out if the last mouse coordinates are inside the region box.
                ' If they are, we want to highlight it.
                Dim useHighlightColor As Boolean
                
                If m_MouseInsideBox Then
                    useHighlightColor = Math_Functions.IsPointInRectF(m_LastMouseX, m_LastMouseY, relativeRect)
                Else
                    useHighlightColor = False
                End If
                
                'Draw a canvas-style border around the relevant viewport rect
                GDI_Plus.GDIPlusDrawCanvasRectF bufferDC, relativeRect, , useHighlightColor
            
            End If
            
        End If
    
    End If
    
    'Paint the final result to the screen, as relevant
    ucSupport.RequestRepaint
    
End Sub

'Call this when a new thumbnail needs to be set.  The class will reset its thumb DIB to match its current size, then raise
' a RequestUpdatedThumbnail function.
Public Sub NotifyNewThumbNeeded()
    
    'Wipe the existing thumbnail, and request a new one.
    If (m_ImageThumbnail Is Nothing) Then
        UpdateControlLayout
    Else
        m_ImageThumbnail.ResetDIB 0
        RaiseEvent RequestUpdatedThumbnail(m_ImageThumbnail, m_ThumbEventX, m_ThumbEventY)
        RedrawBackBuffer
    End If
    
End Sub

'Call this when the viewport position has changed.  This function operates independently of the NotifyNewThumbNeeded() function,
' because the viewport and thumbnail are unlikely to change simultaneously.
Public Sub NotifyNewViewportPosition()
    RedrawBackBuffer
End Sub

'Before this control does any painting, we need to retrieve relevant colors from PD's primary theming class.  Note that this
' step must also be called if/when PD's visual theme settings change.
Private Sub UpdateColorList()
    m_Colors.LoadThemeColor PDNI_Background, "Background", IDE_WHITE
End Sub

'External functions can call this to request a redraw.  This is helpful for live-updating theme settings, as in the Preferences dialog,
' and/or retranslating any text against the current language.
Public Sub UpdateAgainstCurrentTheme()
    If ucSupport.ThemeUpdateRequired Then
        UpdateColorList
        If g_IsProgramRunning Then ucSupport.UpdateAgainstThemeAndLanguage
        UpdateControlLayout
    End If
End Sub

'By design, PD prefers to not use design-time tooltips.  Apply tooltips at run-time, using this function.
' (IMPORTANT NOTE: translations are handled automatically.  Always pass the original English text!)
Public Sub AssignTooltip(ByVal newTooltip As String, Optional ByVal newTooltipTitle As String, Optional ByVal newTooltipIcon As TT_ICON_TYPE = TTI_NONE)
    ucSupport.AssignTooltip UserControl.ContainerHwnd, newTooltip, newTooltipTitle, newTooltipIcon
End Sub

