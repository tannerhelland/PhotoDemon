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
'Last updated: 16/October/15
'Last update: initial build
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
Public Event RequestUpdatedThumbnail(ByRef thumbDIB As pdDIB)

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

'The rect where the image thumbnail has been drawn.  This is calculated by the DrawNavigator function.
Private m_ThumbRect As RECTF

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

Public Property Get ContainerHwnd() As Long
    ContainerHwnd = UserControl.ContainerHwnd
End Property

Private Sub cMouseEvents_MouseEnter(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    m_MouseInsideBox = True
    DrawNavigator
    cMouseEvents.setSystemCursor IDC_HAND
End Sub

Private Sub cMouseEvents_MouseLeave(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    m_MouseInsideBox = False
    DrawNavigator
    cMouseEvents.setSystemCursor IDC_DEFAULT
End Sub

'When the control receives focus, relay the event externally
Private Sub cFocusDetector_GotFocusReliable()
    RaiseEvent GotFocusAPI
End Sub

'When the control loses focus, relay the event externally
Private Sub cFocusDetector_LostFocusReliable()
    RaiseEvent LostFocusAPI
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
    
    'Whenever the navigator is resized, we must resize the thumbnail and back buffer to match.  Start with the thumb.
    Dim thumbSize As Long
    If picNavigator.Width < picNavigator.Height Then
        thumbSize = picNavigator.Width
    Else
        thumbSize = picNavigator.Height
    End If
    
    'Pad the thumb, so we have some empty space around it for borders and the like
    thumbSize = thumbSize - FixDPIFloat(THUMB_PADDING) * 2
    
    'Try to optimize re-creating the thumbnail, so we only do it when absolutely necessary
    If m_ImageThumbnail Is Nothing Then Set m_ImageThumbnail = New pdDIB
    If (m_ImageThumbnail.getDIBWidth <> thumbSize) Or (m_ImageThumbnail.getDIBHeight <> thumbSize) Then
        m_ImageThumbnail.createBlank thumbSize, thumbSize, 32, 0, 0
    Else
        m_ImageThumbnail.resetDIB 0
    End If
    
    RaiseEvent RequestUpdatedThumbnail(m_ImageThumbnail)
    
    'The back buffer itself must also be resized to match the navigator container dimensions.
    If m_BackBuffer Is Nothing Then Set m_BackBuffer = New pdDIB
    If (m_BackBuffer.getDIBWidth <> picNavigator.Width) Or (m_BackBuffer.getDIBHeight <> picNavigator.Height) Then
        m_BackBuffer.createBlank picNavigator.Width, picNavigator.Height, 24
    Else
        m_BackBuffer.resetDIB 0
    End If
    
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
            
            'Paint a checkerboard background, then the thumbnail.
            With m_ThumbRect
                GDI_Plus.GDIPlusFillDIBRect_Pattern m_BackBuffer, .Left, .Top, .Width, .Height, g_CheckerboardPattern, , True
            End With
            
            With m_ThumbRect
                GDI_Plus.GDIPlus_StretchBlt m_BackBuffer, .Left, .Top, .Width, .Height, m_ImageThumbnail, 0, 0, .Width, .Height, , InterpolationModeHighQualityBicubic
            End With
            
            'TODO: rect for the current viewport
            
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
        RaiseEvent RequestUpdatedThumbnail(m_ImageThumbnail)
        DrawNavigator
    End If
    
End Sub

'Due to complex interactions between user controls and PD's translation engine, tooltips require this dedicated function.
' (IMPORTANT NOTE: the tooltip class will handle translations automatically.  Always pass the original English text!)
Public Sub AssignTooltip(ByVal newTooltip As String, Optional ByVal newTooltipTitle As String, Optional ByVal newTooltipIcon As TT_ICON_TYPE = TTI_NONE)
    toolTipManager.setTooltip Me.hWnd, Me.ContainerHwnd, newTooltip, newTooltipTitle, newTooltipIcon
End Sub

'External functions can call this to request a redraw.  This is helpful for live-updating theme settings, as in the Preferences dialog,
' and/or retranslating any text against the current language.
Public Sub UpdateAgainstCurrentTheme()
    
    'Update the tooltip, if any
    If g_IsProgramRunning Then toolTipManager.UpdateAgainstCurrentTheme
        
    'Redraw the control (in case anything has changed)
    UpdateControlSize
    
End Sub
    
