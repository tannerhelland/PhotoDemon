VERSION 5.00
Begin VB.UserControl pdNavigator 
   BackColor       =   &H80000005&
   ClientHeight    =   1185
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1545
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   79
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   103
   ToolboxBitmap   =   "pdNavigator.ctx":0000
   Begin VB.PictureBox picNavigator 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   120
      ScaleHeight     =   57
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   81
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "pdNavigator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Navigation custom control
'Copyright 2015-2015 by Tanner Helland
'Created: 16/October/15
'Last updated: 17/October/15
'Last update: wrap up initial build
'
'In 7.0, a "navigation" panel was added to the right-side toolbar.  This user control provides the actual "navigation"
' behavior, where the user can click anywhere on the image thumbnail to move the viewport over that location.
'
'I've designed this as a UC in case I ever want to add zoom or other buttons, but right now it only consists of the
' navigation box.
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

'A specialized class handles mouse input for this control
Private WithEvents cMouseEvents As pdInputMouse
Attribute cMouseEvents.VB_VarHelpID = -1

'Reliable focus detection requires a specialized subclasser
Private WithEvents cFocusDetector As pdFocusDetector
Attribute cFocusDetector.VB_VarHelpID = -1
Public Event GotFocusAPI()
Public Event LostFocusAPI()

'Flicker-free window painter
Private WithEvents cPainter As pdWindowPainter
Attribute cPainter.VB_VarHelpID = -1

'Additional helper for rendering themed and multiline tooltips
Private toolTipManager As pdToolTip

'The image thumbnail is cached independently, so we only re-request it when absolutely necessary.
Private m_ImageThumbnail As pdDIB

'This back buffer is for the navigator, specifically, including the image thumbnail with borders and a viewport rect
' composited over the top.
Private m_BackBuffer As pdDIB

'This value will be TRUE while the mouse is inside the navigator box
Private m_MouseInsideBox As Boolean

'API technique for drawing a focus rectangle; used only for designer mode (see the Paint method for details)
Private Declare Function DrawFocusRect Lib "user32" (ByVal hDC As Long, lpRect As RECT) As Long

'Padding (in pixels) between the edges of picNavigator and the image thumbnail.  Automatically adjusted for DPI
' at run-time.
Private Const THUMB_PADDING As Long = 3

'When the control raises a request for a new thumbnail image, that function will supply an (optional?) (x, y) pair detailing
' where the thumb is centered within the navigator.  We use this to know where the image lies inside the thumb.
Private m_ThumbEventX As Single, m_ThumbEventY As Single

'The rect where the image thumbnail has been drawn.  This is calculated by the DrawNavigator function.
Private m_ThumbRect As RECTF, m_ImageRegion As RECTF

'Last mouse (x, y) values.  We track these so we know whether to highlight the region box inside the navigator.
Private m_LastMouseX As Single, m_LastMouseY As Single

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

Public Property Get containerHwnd() As Long
    containerHwnd = UserControl.containerHwnd
End Property

'When the control receives focus, relay the event externally
Private Sub cFocusDetector_GotFocusReliable()
    RaiseEvent GotFocusAPI
End Sub

'When the control loses focus, relay the event externally
Private Sub cFocusDetector_LostFocusReliable()
    RaiseEvent LostFocusAPI
End Sub

Private Sub cMouseEvents_MouseDownCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    
    'If the mouse button is clicked inside the image portion of the navigator, scroll to that (x, y) position
    If (Button And pdLeftButton) <> 0 Then
        If Math_Functions.isPointInRectF(x, y, m_ImageRegion) Then ScrollToXY x, y
    End If
    
End Sub

Private Sub cMouseEvents_MouseEnter(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    m_MouseInsideBox = True
End Sub

Private Sub cMouseEvents_MouseLeave(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    m_MouseInsideBox = False
    m_LastMouseX = -1: m_LastMouseY = -1
    cMouseEvents.setSystemCursor IDC_DEFAULT
End Sub

Private Sub cMouseEvents_MouseMoveCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    
    m_LastMouseX = x: m_LastMouseY = y
    
    'Set the cursor depending on whether the mouse is inside the image portion of the navigator control
    If Math_Functions.isPointInRectF(x, y, m_ImageRegion) Then
        cMouseEvents.setSystemCursor IDC_HAND
    Else
        cMouseEvents.setSystemCursor IDC_DEFAULT
    End If
    
    'If the mouse button is down, scroll to that (x, y) position.  Note that we don't care if the cursor is in-bounds;
    ' the ScrollToXY function will automatically fix that for us.
    If (Button And pdLeftButton) <> 0 Then
        ScrollToXY x, y
    Else
        DrawNavigator
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
        hScrollRange = FormMain.mainCanvas(0).getScrollMax(PD_HORIZONTAL) - FormMain.mainCanvas(0).getScrollMin(PD_HORIZONTAL)
        vScrollRange = FormMain.mainCanvas(0).getScrollMax(PD_VERTICAL) - FormMain.mainCanvas(0).getScrollMin(PD_VERTICAL)
        newHScroll = (xRatio * hScrollRange) + FormMain.mainCanvas(0).getScrollMin(PD_HORIZONTAL)
        newVscroll = (yRatio * vScrollRange) + FormMain.mainCanvas(0).getScrollMin(PD_VERTICAL)
        
        'Assign the new scrollbar values, then request a viewport refresh
        FormMain.mainCanvas(0).setRedrawSuspension True
        FormMain.mainCanvas(0).setScrollValue PD_HORIZONTAL, newHScroll
        FormMain.mainCanvas(0).setScrollValue PD_VERTICAL, newVscroll
        FormMain.mainCanvas(0).setRedrawSuspension False
        
        Viewport_Engine.Stage3_ExtractRelevantRegion pdImages(g_CurrentImage), FormMain.mainCanvas(0)
        
    End If

End Sub

'The pdWindowPaint class raises this event when the navigator box needs to be redrawn.  The passed coordinates contain
' the rect returned by GetUpdateRect (but with right/bottom measurements pre-converted to width/height).
Private Sub cPainter_PaintWindow(ByVal winLeft As Long, ByVal winTop As Long, ByVal winWidth As Long, ByVal winHeight As Long)
    
    'Flip the relevant chunk of the buffer to the screen
    BitBlt picNavigator.hDC, winLeft, winTop, winWidth, winHeight, m_BackBuffer.getDIBDC, winLeft, winTop, vbSrcCopy
    
End Sub

Private Sub UserControl_Initialize()
    
    If g_IsProgramRunning Then
        
        'Initialize mouse handling
        Set cMouseEvents = New pdInputMouse
        cMouseEvents.addInputTracker picNavigator.hWnd, True, True, , True, True
        cMouseEvents.setSystemCursor IDC_HAND
        
        'Also start a focus detector
        Set cFocusDetector = New pdFocusDetector
        cFocusDetector.startFocusTracking Me.hWnd
        
        'Also start a flicker-free window painter
        Set cPainter = New pdWindowPainter
        cPainter.startPainter picNavigator.hWnd
        
        'Create a tooltip engine
        Set toolTipManager = New pdToolTip
    
    'In design mode, initialize a base theming class, so our paint function doesn't fail
    Else
        If g_Themer Is Nothing Then Set g_Themer = New pdVisualThemes
    End If
    
    'Draw the control at least once
    UpdateControlSize
    
End Sub

'At run-time, painting is handled by PD's pdWindowPainter class.  In the IDE, however, we must rely on VB's internal paint event.
Private Sub UserControl_Paint()
    
    'Provide minimal painting within the designer
    If Not g_IsProgramRunning Then DrawNavigator
    
End Sub

Private Sub UserControl_Resize()
    UpdateControlSize
End Sub

'Call this to recreate all buffers against a changed control size.
Private Sub UpdateControlSize()
    
    'For now, we simply sync the navigator box to the size of the control
    picNavigator.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
    
    'Resize the back buffer to match the navigator container dimensions.
    If m_BackBuffer Is Nothing Then Set m_BackBuffer = New pdDIB
    If (m_BackBuffer.getDIBWidth <> picNavigator.Width) Or (m_BackBuffer.getDIBHeight <> picNavigator.Height) Then
        m_BackBuffer.createBlank picNavigator.Width, picNavigator.Height, 24
    Else
        m_BackBuffer.resetDIB 0
    End If
    
    'Whenever the navigator is resized, we must also resize the image thumbnail to match.
    
    'At present, we pad the thumbnail by a few pixels so we have room for a border.
    Dim thumbWidth As Long, thumbHeight As Long
    thumbWidth = m_BackBuffer.getDIBWidth - FixDPIFloat(THUMB_PADDING) * 2
    thumbHeight = m_BackBuffer.getDIBHeight - FixDPIFloat(THUMB_PADDING) * 2
    
    'Try to optimize re-creating the thumbnail, so we only do it when absolutely necessary
    If m_ImageThumbnail Is Nothing Then Set m_ImageThumbnail = New pdDIB
    If (m_ImageThumbnail.getDIBWidth <> thumbWidth) Or (m_ImageThumbnail.getDIBHeight <> thumbHeight) Then
        m_ImageThumbnail.createBlank thumbWidth, thumbHeight, 32, 0, 0
    Else
        m_ImageThumbnail.resetDIB 0
    End If
    
    RaiseEvent RequestUpdatedThumbnail(m_ImageThumbnail, m_ThumbEventX, m_ThumbEventY)
    
    'With the backbuffer and image thumbnail successfully created, we can finally redraw the new navigator window
    DrawNavigator
    
End Sub

'Need to redraw the navigator box?  Call this.  Note that it *does not* request a new image thumbnail.  You must handle
' that separately.  This simply uses whatever's been previously cached.
Private Sub DrawNavigator()
    
    'Make sure the back buffer exists.  (This is always problematic in the IDE.)
    If m_BackBuffer Is Nothing Then
        m_BackBuffer.createBlank picNavigator.Width, picNavigator.Height, 24, RGB(255, 255, 255)
    End If
    
    If g_IsProgramRunning Then
    
        'Repaint the background.  Color is still TBD.
        GDI_Plus.GDIPlusFillDIBRect m_BackBuffer, 0, 0, m_BackBuffer.getDIBWidth, m_BackBuffer.getDIBHeight, g_Themer.getThemeColor(PDTC_BACKGROUND_DEFAULT), 255
        
        'If no images are open, do nothing
        If g_OpenImageCount <= 0 Then
            
            With m_ThumbRect
                .Width = 0
                .Height = 0
                .Left = 0
                .Top = 0
            End With
            
        Else
            
            'Determine a position rect for the image thumbnail
            With m_ThumbRect
                .Width = m_ImageThumbnail.getDIBWidth
                .Height = m_ImageThumbnail.getDIBHeight
                .Left = (m_BackBuffer.getDIBWidth - m_ImageThumbnail.getDIBWidth) / 2
                .Top = (m_BackBuffer.getDIBHeight - m_ImageThumbnail.getDIBHeight) / 2
            End With
            
            'Offset that top-left corner by the thumbnail's position, and cache it to a module-level rect so we can use it for
            ' hit-detection during mouse events.
            With m_ImageRegion
                .Left = m_ThumbRect.Left + m_ThumbEventX
                .Top = m_ThumbRect.Top + m_ThumbEventY
                .Width = m_ImageThumbnail.getDIBWidth - (m_ThumbEventX * 2)
                .Height = m_ImageThumbnail.getDIBHeight - (m_ThumbEventY * 2)
            End With
            
            'Paint a checkerboard background only over the relevant image region
            With m_ImageRegion
                GDI_Plus.GDIPlusFillDIBRect_Pattern m_BackBuffer, .Left, .Top, .Width, .Height, g_CheckerboardPattern, , True
            End With
            
            'Paint the thumb rect without regard for the image region (as it will always be a square)
            With m_ThumbRect
                GDI_Plus.GDIPlus_StretchBlt m_BackBuffer, .Left, .Top, .Width, .Height, m_ImageThumbnail, 0, 0, .Width, .Height, , InterpolationModeHighQualityBicubic
            End With
                        
            'Query the active image for a copy of the intersection rect of the viewport, and the image itself, in image coordinate space
            Dim viewportRect As RECTF
            pdImages(g_CurrentImage).imgViewport.getIntersectRectImage viewportRect
            
            'We now want to convert the viewport rect into our little navigator coordinate space.  Start by converting the viewport
            ' dimensions to a 1-based system, relative to the original image's width and height.
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
                    useHighlightColor = Math_Functions.isPointInRectF(m_LastMouseX, m_LastMouseY, relativeRect)
                Else
                    useHighlightColor = False
                End If
                
                'Draw a canvas-style border around the relevant viewport rect
                GDI_Plus.GDIPlusDrawCanvasRectF m_BackBuffer.getDIBDC, relativeRect, , useHighlightColor
            
            End If
            
        End If
    
    'In the designer, draw a focus rect around the control; this is minimal feedback required for positioning
    Else
        
        Dim tmpRect As RECT
        With tmpRect
            .Left = 0
            .Top = 0
            .Right = m_BackBuffer.getDIBWidth
            .Bottom = m_BackBuffer.getDIBHeight
        End With
        
        DrawFocusRect m_BackBuffer.getDIBDC, tmpRect

    End If
    
    'Paint the final result to the screen, as relevant
    If g_IsProgramRunning Then
        cPainter.requestRepaint
    Else
        BitBlt picNavigator.hDC, 0, 0, picNavigator.ScaleWidth, picNavigator.ScaleHeight, m_BackBuffer.getDIBDC, 0, 0, vbSrcCopy
    End If
    
End Sub

'Call this when a new thumbnail needs to be set.  The class will reset its thumb DIB to match its current size, then raise
' a RequestUpdatedThumbnail function.
Public Sub NotifyNewThumbNeeded()
    
    'Wipe the existing thumbnail, and request a new one.
    If m_ImageThumbnail Is Nothing Then
        UpdateControlSize
    Else
        m_ImageThumbnail.resetDIB 0
        RaiseEvent RequestUpdatedThumbnail(m_ImageThumbnail, m_ThumbEventX, m_ThumbEventY)
        DrawNavigator
    End If
    
End Sub

'Call this when the viewport position has changed.  This function operates independently of the NotifyNewThumbNeeded() function,
' because the viewport and thumbnail are unlikely to change simultaneously.
Public Sub NotifyNewViewportPosition()
    DrawNavigator
End Sub

'Due to complex interactions between user controls and PD's translation engine, tooltips require this dedicated function.
' (IMPORTANT NOTE: the tooltip class will handle translations automatically.  Always pass the original English text!)
Public Sub AssignTooltip(ByVal newTooltip As String, Optional ByVal newTooltipTitle As String, Optional ByVal newTooltipIcon As TT_ICON_TYPE = TTI_NONE)
    toolTipManager.setTooltip Me.hWnd, Me.containerHwnd, newTooltip, newTooltipTitle, newTooltipIcon
End Sub

'External functions can call this to request a redraw.  This is helpful for live-updating theme settings, as in the Preferences dialog,
' and/or retranslating any text against the current language.
Public Sub UpdateAgainstCurrentTheme()
    
    'Update the tooltip, if any
    If g_IsProgramRunning Then toolTipManager.UpdateAgainstCurrentTheme
        
    'Redraw the control (in case anything has changed)
    UpdateControlSize
    
End Sub
    
