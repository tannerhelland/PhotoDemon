VERSION 5.00
Begin VB.UserControl pdPictureBoxInteractive 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   CanGetFocus     =   0   'False
   ClientHeight    =   795
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3900
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
   ScaleHeight     =   53
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   260
   ToolboxBitmap   =   "pdPictureBoxInteractive.ctx":0000
End
Attribute VB_Name = "pdPictureBoxInteractive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Interactive PictureBox Replacement control
'Copyright 2018-2026 by Tanner Helland
'Created: 21/March/18
'Last updated: 12/November/20
'Last update: new event after a window resize has occurred; the owner may need to perform a special redraw
'
'For interactive UI elements that don't warrant a dedicated user-control, use this control instead.
' It basically acts as a thin operator to a pdUCSupport instance, but you'll need to manually handle
' input (and backbuffer rendering) for the control to do anything useful.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'This control is always owner-drawn
Public Event DrawMe(ByVal targetDC As Long, ByVal ctlWidth As Long, ByVal ctlHeight As Long)

Public Event GotFocusAPI()
Public Event LostFocusAPI()

Public Event MouseDownCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal timeStamp As Long)
Public Event MouseEnter(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
Public Event MouseLeave(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
Public Event MouseMoveCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal timeStamp As Long)
Public Event MouseUpCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal clickEventAlsoFiring As Boolean, ByVal timeStamp As Long)
Public Event Resize(ByVal newWidth As Long, ByVal newHeight As Long)

'After a window resize, this control may require a special redraw (because its dimensions may
' have changed).  A special event exists for notifying this state, as not all dialogs need to
' implement it.
Public Event WindowResizeDetected()

'User control support class.  Historically, many classes (and associated subclassers) were required by each user control,
' but I've since wrapped these into a single central support class.
Private WithEvents ucSupport As pdUCSupport
Attribute ucSupport.VB_VarHelpID = -1

Public Function GetControlType() As PD_ControlType
    GetControlType = pdct_PictureBoxInteractive
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
    RedrawBackBuffer
    PropertyChanged "Enabled"
End Property

'hWnds aren't exposed by default
Public Property Get hWnd() As Long
Attribute hWnd.VB_UserMemId = -515
    hWnd = UserControl.hWnd
End Property

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

Public Sub SetSize(ByVal newWidth As Long, ByVal newHeight As Long)
    ucSupport.RequestNewSize newWidth, newHeight, True
End Sub

Public Sub RequestCursor(ByVal newCursor As SystemCursorConstant)
    ucSupport.RequestCursor newCursor
End Sub

Public Sub RequestRedraw(Optional ByVal paintImmediately As Boolean = False)
    RedrawBackBuffer paintImmediately
End Sub

'If the caller wants to render out of order, they can use StartPaint/EndPaint
Public Sub StartPaint(ByRef dstDC As Long, ByRef dstWidth As Long, ByRef dstHeight As Long, Optional ByVal repaintBackground As Boolean = False, Optional ByVal newBackColor As Long = -1&)
    dstDC = ucSupport.GetBackBufferDC(repaintBackground, newBackColor)
    dstWidth = ucSupport.GetBackBufferWidth()
    dstHeight = ucSupport.GetBackBufferHeight()
End Sub

Public Sub EndPaint(Optional ByVal raiseImmediateDrawEvent As Boolean = False)
    ucSupport.RequestRepaint raiseImmediateDrawEvent
End Sub

'Sometimes (typically failure cases), PD needs to simply throw a warning message onto a picture box.
' This function makes it trivial.
Public Sub PaintText(ByRef srcString As String, Optional ByVal FontSize As Single = 12!, Optional ByVal isBold As Boolean = False)

    Dim dstDC As Long, dstWidth As Long, dstHeight As Long
    dstDC = ucSupport.GetBackBufferDC(True)
    dstWidth = ucSupport.GetBackBufferWidth()
    dstHeight = ucSupport.GetBackBufferHeight()
    
    Dim tmpFont As pdFont
    Set tmpFont = Fonts.GetMatchingUIFont(FontSize, isBold)
    
    Dim tmpRect As RECT
    With tmpRect
        .Left = 0
        .Top = 0
        .Right = dstWidth - 1
        .Bottom = dstHeight - 1
    End With
    
    tmpFont.SetFontColor g_Themer.GetGenericUIColor(UI_TextReadOnly)
    tmpFont.AttachToDC dstDC
    tmpFont.SetTextAlignment vbLeftJustify
    tmpFont.DrawCenteredTextToRect srcString, tmpRect, True
    tmpFont.ReleaseFromDC
    
    Set tmpFont = Nothing
    
End Sub

'For convenience, if you need a DIB painted in a centered position, use this function; we take care of the rest
Public Sub CopyDIB(ByRef srcDIB As pdDIB, Optional ByVal colorManagementMatters As Boolean = True, Optional ByVal doNotStretchIfSmaller As Boolean = False, Optional ByVal suspendTransparencyGrid As Boolean = False, Optional ByVal useNeutralBackground As Boolean = False, Optional ByVal drawBorderAroundImage As Boolean = False, Optional ByVal drawBorderAroundControl As Boolean = True)
    
    If (srcDIB Is Nothing) Then
        PDDebug.LogAction "WARNING!  pdPictureBox.CopyDIB received a null DIB; copy abandoned."
        Exit Sub
    End If
    
    Dim dstWidth As Double, dstHeight As Double
    dstWidth = Me.GetWidth
    dstHeight = Me.GetHeight
    
    Dim dstWidthOrig As Double, dstHeightOrig As Double
    dstWidthOrig = dstWidth
    dstHeightOrig = dstHeight
    
    Dim srcWidth As Double, srcHeight As Double
    srcWidth = srcDIB.GetDIBWidth
    srcHeight = srcDIB.GetDIBHeight
    
    'If the caller expects the source image to be small, they may prevent us from enlarging the image to fit
    Dim fitPrevented As Boolean: fitPrevented = False
    If doNotStretchIfSmaller Then
        If (srcWidth < dstWidth) And (srcHeight < dstHeight) Then
            fitPrevented = True
            dstWidth = srcWidth
            dstHeight = srcHeight
        End If
    End If
    
    'Calculate the aspect ratio of this DIB and the target picture box
    Dim srcAspect As Double, dstAspect As Double
    If (srcHeight > 0) Then srcAspect = srcWidth / srcHeight Else srcAspect = 1#
    If (dstHeight > 0) Then dstAspect = dstWidth / dstHeight Else dstAspect = 1#
    
    Dim dWidth As Long, dHeight As Long, previewX As Long, previewY As Long
    PDMath.ConvertAspectRatio srcWidth, srcHeight, dstWidth, dstHeight, dWidth, dHeight
    
    If fitPrevented Then
        previewX = (dstWidthOrig - srcWidth) * 0.5
        previewY = (dstHeightOrig - srcHeight) * 0.5
    Else
        If (srcAspect > dstAspect) Then
            previewY = Int((dstHeight - dHeight) * 0.5)
            previewX = 0
        Else
            previewX = Int((dstWidth - dWidth) * 0.5)
            previewY = 0
        End If
    End If
    
    'Grab a DC
    Dim dstDC As Long
    dstDC = ucSupport.GetBackBufferDC(True, IIf(useNeutralBackground, RGB(127, 127, 127), -1))
    
    'Paint a transparency grid, as required
    Dim dstSurface As pd2DSurface: Set dstSurface = New pd2DSurface
    dstSurface.WrapSurfaceAroundDC dstDC
    dstSurface.SetSurfaceAntialiasing P2_AA_None
    dstSurface.SetSurfacePixelOffset P2_PO_Normal
    If (Not suspendTransparencyGrid) Then PD2D.FillRectangleI dstSurface, g_CheckerboardBrush, previewX, previewY, dWidth, dHeight
    
    'Finally, paint the image itself
    Dim srcSurface As pd2DSurface: Set srcSurface = New pd2DSurface
    srcSurface.WrapSurfaceAroundPDDIB srcDIB
    dstSurface.SetSurfaceResizeQuality P2_RQ_Bicubic
    PD2D.DrawSurfaceResizedCroppedI dstSurface, previewX, previewY, dWidth, dHeight, srcSurface, 0, 0, srcWidth, srcHeight
    
    'Free the source DIB from its DC, as a convenience
    Set srcSurface = Nothing
    srcDIB.FreeFromDC
    
    'Remaining operations require a RectF; populate one now
    Dim tmpRectF As RectF
    tmpRectF.Left = previewX
    tmpRectF.Top = previewY
    tmpRectF.Width = dWidth - 1
    tmpRectF.Height = dHeight - 1
    
    'Color-management is handled by ucSupport
    If colorManagementMatters Then ucSupport.RequestBufferColorManagement VarPtr(tmpRectF)
    
    'As a convenience, we can draw a border around the image or entire control
    If drawBorderAroundImage Or drawBorderAroundControl Then
        
        Dim cPen As pd2DPen
        Drawing2D.QuickCreateSolidPen cPen, 1!, g_Themer.GetGenericUIColor(UI_CanvasElement)
        
        If drawBorderAroundImage Then PD2D.DrawRectangleF_FromRectF dstSurface, cPen, tmpRectF
        If drawBorderAroundControl Then
            tmpRectF.Left = 0!
            tmpRectF.Top = 0!
            tmpRectF.Width = dstWidthOrig - 1
            tmpRectF.Height = dstHeightOrig - 1
            PD2D.DrawRectangleF_FromRectF dstSurface, cPen, tmpRectF
        End If
        
        Set cPen = Nothing: Set dstSurface = Nothing
        
    End If
    
    'Repaint the screen!
    ucSupport.RequestRepaint True
    
End Sub

Public Sub SetCursorCustom(Optional ByVal standardCursorType As SystemCursorConstant = IDC_DEFAULT)
    ucSupport.RequestCursor standardCursorType
End Sub

Public Sub SetCursorCustom_Resource(ByVal pngResourceName As String, Optional ByVal cursorHotspotX As Long = 0, Optional ByVal cursorHotspotY As Long = 0)
    ucSupport.RequestCursor_Resource pngResourceName, cursorHotspotX, cursorHotspotY
End Sub

Private Sub ucSupport_CustomMessage(ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturn As Long)
    If (wMsg = WM_PD_DIALOG_RESIZE_FINISHED) Then RaiseEvent WindowResizeDetected
End Sub

Private Sub ucSupport_GotFocusAPI()
    RaiseEvent GotFocusAPI
    RedrawBackBuffer
End Sub

Private Sub ucSupport_LostFocusAPI()
    RaiseEvent LostFocusAPI
    RedrawBackBuffer
End Sub

Private Sub ucSupport_MouseDownCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal timeStamp As Long)
    RaiseEvent MouseDownCustom(Button, Shift, x, y, timeStamp)
End Sub

Private Sub ucSupport_MouseEnter(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    ucSupport.RequestCursor IDC_ARROW
    RaiseEvent MouseEnter(Button, Shift, x, y)
End Sub

Private Sub ucSupport_MouseLeave(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    RaiseEvent MouseLeave(Button, Shift, x, y)
End Sub

Private Sub ucSupport_MouseMoveCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal timeStamp As Long)
    RaiseEvent MouseMoveCustom(Button, Shift, x, y, timeStamp)
End Sub

Private Sub ucSupport_MouseUpCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal clickEventAlsoFiring As Boolean, ByVal timeStamp As Long)
    RaiseEvent MouseUpCustom(Button, Shift, x, y, clickEventAlsoFiring, timeStamp)
End Sub

Private Sub ucSupport_RepaintRequired(ByVal updateLayoutToo As Boolean)
    RedrawBackBuffer
End Sub

Private Sub ucSupport_WindowResize(ByVal newWidth As Long, ByVal newHeight As Long)
    RaiseEvent Resize(newWidth, newHeight)
End Sub

Private Sub UserControl_Initialize()
    
    'Initialize a user control support class
    Set ucSupport = New pdUCSupport
    ucSupport.RegisterControl UserControl.hWnd, True, False
    ucSupport.RequestExtraFunctionality True
    ucSupport.SubclassCustomMessage WM_PD_DIALOG_RESIZE_FINISHED, True
    ucSupport.RequestHighPerformanceRendering True
    
End Sub

'At run-time, painting is handled by PD's pdWindowPainter class.  In the IDE, however, we must rely on VB's internal paint event.
Private Sub UserControl_Paint()
    ucSupport.RequestIDERepaint UserControl.hDC
End Sub

Private Sub UserControl_Resize()
    If (Not PDMain.IsProgramRunning()) Then ucSupport.RequestRepaint True
End Sub

'External functions can call this to request a redraw.  This is helpful for live-updating theme settings, as in the Preferences dialog.
Public Sub UpdateAgainstCurrentTheme(Optional ByVal hostFormhWnd As Long = 0)
    If ucSupport.ThemeUpdateRequired And PDMain.IsProgramRunning() Then
        NavKey.NotifyControlLoad Me, hostFormhWnd, False
        ucSupport.UpdateAgainstThemeAndLanguage
        RedrawBackBuffer
    End If
End Sub

'Use this function to completely redraw the back buffer from scratch.  Note that this is computationally expensive compared to just flipping the
' existing buffer to the screen, so only redraw the backbuffer if the control state has somehow changed.
Private Sub RedrawBackBuffer(Optional ByVal paintImmediately As Boolean = False)
    
    'Request the back buffer DC, and ask the support module to erase any existing rendering for us.
    Dim bufferDC As Long
    bufferDC = ucSupport.GetBackBufferDC(True)
    If (bufferDC <> 0) Then RaiseEvent DrawMe(bufferDC, ucSupport.GetBackBufferWidth, ucSupport.GetBackBufferHeight)
    
    'Paint the final result to the screen, as relevant
    ucSupport.RequestRepaint paintImmediately
    
End Sub

'By design, PD prefers to not use design-time tooltips.  Apply tooltips at run-time, using this function.
' (IMPORTANT NOTE: translations are handled automatically.  Always pass the original English text!)
Public Sub AssignTooltip(ByRef newTooltip As String, Optional ByRef newTooltipTitle As String = vbNullString, Optional ByVal raiseTipsImmediately As Boolean = False)
    ucSupport.AssignTooltip UserControl.ContainerHwnd, newTooltip, newTooltipTitle, raiseTipsImmediately
End Sub
