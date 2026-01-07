VERSION 5.00
Begin VB.UserControl pdNewOld 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
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
   ToolboxBitmap   =   "pdNewOld.ctx":0000
End
Attribute VB_Name = "pdNewOld"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon New/Old comparison control
'Copyright 2016-2026 by Tanner Helland
'Created: 14/October/16
'Last updated: 14/October/16
'Last update: initial build
'
'This user control is currently used in the color selection dialog.  It provides a semi-owner-drawn mechanism
' for displaying a "new" and "old" value side-by-side.  The user can click the "old" value to make it the
' "new" value, sort of like a reset option.
'
'This control is best used in places where a side-by-side comparison between two elements is required.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

Public Event OldItemClicked()
Public Event DrawNewItem(ByVal targetDC As Long, ByVal ptrToRectF As Long)
Public Event DrawOldItem(ByVal targetDC As Long, ByVal ptrToRectF As Long)

'Because VB focus events are wonky, especially when we use CreateWindow within a UC, this control raises its own
' specialized focus events.  If you need to track focus, use these instead of the default VB functions.
Public Event GotFocusAPI()
Public Event LostFocusAPI()

'To simplify rendering, we pre-calculate positions for the captions "old" and "new", and rects for the old and
' new sample areas.  These values are calculated by UpdateControlLayout, and they must be recalculated if the
' control size or current language changes (as the new translations may be longer/shorter).
Private m_FontSize As Single
Private m_NewCaptionTranslated As String, m_OldCaptionTranslated As String
Private m_NewCaptionPt As PointFloat, m_OldCaptionPt As PointFloat
Private m_NewItemRect As RectF, m_OldItemRect As RectF

'The only hoverable item in this control is the "old" item rect
Private m_OldItemIsHovered As Boolean

'User control support class.  Historically, many classes (and associated subclassers) were required by each user control,
' but I've since wrapped these into a single central support class.
Private WithEvents ucSupport As pdUCSupport
Attribute ucSupport.VB_VarHelpID = -1

'Local list of themable colors.  This list includes all potential colors used by this class, regardless of state change
' or internal control settings.  The list is updated by calling the UpdateColorList function.
' (Note also that this list does not include variants, e.g. "BorderColor" vs "BorderColor_Hovered".  Variant values are
'  automatically calculated by the color management class, and they are retrieved by passing boolean modifiers to that
'  class, rather than treating every imaginable variant as a separate constant.)
Private Enum PDNEWOLD_COLOR_LIST
    [_First] = 0
    PDNO_Background = 0
    PDNO_Caption = 1
    PDNO_Border = 2
    [_Last] = 2
    [_Count] = 3
End Enum

'Color retrieval and storage is handled by a dedicated class; this allows us to optimize theme interactions,
' without worrying about the details locally.
Private m_Colors As pdThemeColors

Public Function GetControlType() As PD_ControlType
    GetControlType = pdct_NewOld
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

Public Property Get FontSize() As Single
    FontSize = m_FontSize
End Property

Public Property Let FontSize(ByVal newSize As Single)
    m_FontSize = newSize
    PropertyChanged "FontSize"
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

Public Sub RequestRedraw(Optional ByVal paintImmediately As Boolean = False)
    RedrawBackBuffer paintImmediately
End Sub

Private Sub ucSupport_GotFocusAPI()
    RaiseEvent GotFocusAPI
    RedrawBackBuffer
End Sub

Private Sub ucSupport_KeyDownSystem(ByVal Shift As ShiftConstants, ByVal whichSysKey As PD_NavigationKey, markEventHandled As Boolean)
    
    'Enter/Esc get reported directly to the system key handler.  Note that we track the return, because TRUE
    ' means the key was successfully forwarded to the relevant handler.  (If FALSE is returned, no control
    ' accepted the keypress, meaning we should forward the event down the line.)
    markEventHandled = NavKey.NotifyNavKeypress(Me, whichSysKey, Shift)
    
End Sub

Private Sub ucSupport_LostFocusAPI()
    RaiseEvent LostFocusAPI
    RedrawBackBuffer
End Sub

'Only left clicks raise Click() events
Private Sub ucSupport_MouseDownCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal timeStamp As Long)
    
    If Me.Enabled And ((Button And pdLeftButton) <> 0) Then
        
        'We do not raise events for clicking the "new" option.  Only the "old" option.
        If PDMath.IsPointInRectF(x, y, m_OldItemRect) Then
            RaiseEvent OldItemClicked
        End If
        
    End If
    
End Sub

Private Sub ucSupport_MouseEnter(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    ucSupport.RequestCursor IDC_HAND
    RedrawBackBuffer
End Sub

'When the mouse leaves the UC, we must repaint the button (as it's no longer hovered)
Private Sub ucSupport_MouseLeave(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    m_OldItemIsHovered = False
    RedrawBackBuffer
End Sub

Private Sub ucSupport_MouseMoveCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal timeStamp As Long)
    Dim oldHoverCheck As Boolean
    oldHoverCheck = m_OldItemIsHovered
    m_OldItemIsHovered = PDMath.IsPointInRectF(x, y, m_OldItemRect)
    If m_OldItemIsHovered Then ucSupport.RequestCursor IDC_HAND Else ucSupport.RequestCursor IDC_DEFAULT
    If (oldHoverCheck <> m_OldItemIsHovered) Then RedrawBackBuffer
End Sub

Private Sub ucSupport_RepaintRequired(ByVal updateLayoutToo As Boolean)
    If updateLayoutToo Then UpdateControlLayout Else RedrawBackBuffer
End Sub

Private Sub UserControl_Initialize()
    
    'Initialize a user control support class
    Set ucSupport = New pdUCSupport
    ucSupport.RegisterControl UserControl.hWnd, True
    
    'Request any control-specific functionality
    ucSupport.RequestExtraFunctionality True
    
    'Prep the color manager and load default colors
    Set m_Colors = New pdThemeColors
    Dim colorCount As PDNEWOLD_COLOR_LIST: colorCount = [_Count]
    m_Colors.InitializeColorList "PDNewOld", colorCount
    If (Not PDMain.IsProgramRunning()) Then UpdateColorList
    
    'Inside the IDE, use placeholder text
    If (Not PDMain.IsProgramRunning()) Then
        m_NewCaptionTranslated = "new:"
        m_OldCaptionTranslated = "original:"
    End If
    
End Sub

'Set default properties
Private Sub UserControl_InitProperties()
    FontSize = 12
End Sub

'At run-time, painting is handled by PD's pdWindowPainter class.  In the IDE, however, we must rely on VB's internal paint event.
Private Sub UserControl_Paint()
    ucSupport.RequestIDERepaint UserControl.hDC
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        FontSize = .ReadProperty("FontSize", 12)
    End With
End Sub

Private Sub UserControl_Resize()
    If (Not PDMain.IsProgramRunning()) Then ucSupport.RequestRepaint True
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "FontSize", m_FontSize, 12
    End With
End Sub

'Call this layout calculator whenever the control size changes
Private Sub UpdateControlLayout()

    'Retrieve DPI-aware control dimensions from the support class
    Dim bWidth As Long, bHeight As Long
    bWidth = ucSupport.GetBackBufferWidth
    bHeight = ucSupport.GetBackBufferHeight
    
    Const hTextPadding As Long = 4&
    
    'Vertical dimensions are easy to calculate.  Start by dividing the available vertical space in half.
    With m_NewItemRect
        .Top = 1
        .Height = (bHeight \ 2) - 1
    End With
    
    With m_OldItemRect
        .Top = m_NewItemRect.Top + m_NewItemRect.Height
        .Height = (bHeight - .Top) - 2
    End With
    
    'Next, we need to calculate the size of the "new" and "old" captions.  We want to align these within the
    ' same column, so we need to know which is larger (in terms of pixels).
    Dim newCaptionRect As RectF, oldCaptionRect As RectF
    Dim tmpFont As pdFont
    Set tmpFont = Fonts.GetMatchingUIFont(Me.FontSize)
    
    With newCaptionRect
        .Width = tmpFont.GetWidthOfString(m_NewCaptionTranslated)
        .Height = tmpFont.GetHeightOfString(m_NewCaptionTranslated)
    End With
    
    With oldCaptionRect
        .Width = tmpFont.GetWidthOfString(m_OldCaptionTranslated)
        .Height = tmpFont.GetHeightOfString(m_OldCaptionTranslated)
    End With
    
    Set tmpFont = Nothing
    
    Dim largestWidth As Long
    largestWidth = PDMath.Max2Int(newCaptionRect.Width, oldCaptionRect.Width) + hTextPadding * 2
    
    'The captions are right-aligned against the new and old sample boxes (which are stacked atop each other)
    newCaptionRect.Top = m_NewItemRect.Top + (m_NewItemRect.Height - newCaptionRect.Height) \ 2 - 1
    newCaptionRect.Left = largestWidth - hTextPadding - newCaptionRect.Width
    oldCaptionRect.Top = m_OldItemRect.Top + (m_OldItemRect.Height - oldCaptionRect.Height) \ 2 - 1
    oldCaptionRect.Left = largestWidth - hTextPadding - oldCaptionRect.Width
    
    m_NewCaptionPt.x = newCaptionRect.Left
    m_NewCaptionPt.y = newCaptionRect.Top
    m_OldCaptionPt.x = oldCaptionRect.Left
    m_OldCaptionPt.y = oldCaptionRect.Top
    
    'We now have everything we need to calculate the box positions and widths
    m_NewItemRect.Left = largestWidth + 1
    m_NewItemRect.Width = (bWidth - m_NewItemRect.Left) - 2
    m_OldItemRect.Left = m_NewItemRect.Left
    m_OldItemRect.Width = m_NewItemRect.Width
        
    'No other special preparation is required for this control, so proceed with recreating the back buffer
    RedrawBackBuffer
            
End Sub

'Before this control does any painting, we need to retrieve relevant colors from PD's primary theming class.  Note that this
' step must also be called if/when PD's visual theme settings change.
Private Sub UpdateColorList()
    With m_Colors
        .LoadThemeColor PDNO_Background, "Background", IDE_WHITE
        .LoadThemeColor PDNO_Caption, "Caption", IDE_GRAY
        .LoadThemeColor PDNO_Border, "Border", IDE_GRAY
    End With
End Sub

'External functions can call this to request a redraw.  This is helpful for live-updating theme settings, as in the Preferences dialog.
Public Sub UpdateAgainstCurrentTheme(Optional ByVal hostFormhWnd As Long = 0)
    
    If ucSupport.ThemeUpdateRequired Then
        
        UpdateColorList
        
        If PDMain.IsProgramRunning() Then NavKey.NotifyControlLoad Me, hostFormhWnd
        
        If PDMain.IsProgramRunning() Then
            m_NewCaptionTranslated = g_Language.TranslateMessage("new:")
            m_OldCaptionTranslated = g_Language.TranslateMessage("original:")
            NavKey.NotifyControlLoad Me, hostFormhWnd, False
            ucSupport.UpdateAgainstThemeAndLanguage
        Else
            m_NewCaptionTranslated = "new:"
            m_OldCaptionTranslated = "original:"
        End If
        
    End If
    
End Sub

'Use this function to completely redraw the back buffer from scratch.  Note that this is computationally expensive compared to just flipping the
' existing buffer to the screen, so only redraw the backbuffer if the control state has somehow changed.
Private Sub RedrawBackBuffer(Optional ByVal paintImmediately As Boolean = False)
    
    'Request the back buffer DC, and ask the support module to erase any existing rendering for us.
    Dim bufferDC As Long
    bufferDC = ucSupport.GetBackBufferDC(True, m_Colors.RetrieveColor(PDNO_Background, Me.Enabled))
    If (bufferDC = 0) Then Exit Sub
        
    'Before doing anything else, ask our owner to paint the "new" and "old" areas
    RaiseEvent DrawNewItem(bufferDC, VarPtr(m_NewItemRect))
    RaiseEvent DrawOldItem(bufferDC, VarPtr(m_OldItemRect))
    
    'Next, paint the "new" and "old" captions
    Dim tmpFont As pdFont
    Set tmpFont = Fonts.GetMatchingUIFont(Me.FontSize)
    
    With tmpFont
        .SetTextAlignment vbLeftJustify
        .SetFontColor m_Colors.RetrieveColor(PDNO_Caption, Me.Enabled)
    End With
    
    tmpFont.AttachToDC bufferDC
    tmpFont.FastRenderText m_NewCaptionPt.x, m_NewCaptionPt.y, m_NewCaptionTranslated
    tmpFont.FastRenderText m_OldCaptionPt.x, m_OldCaptionPt.y, m_OldCaptionTranslated
    
    tmpFont.ReleaseFromDC
    Set tmpFont = Nothing
    
    'Next, draw borders around the new and old items
    If PDMain.IsProgramRunning() Then
            
        Dim cSurface As pd2DSurface, cPen As pd2DPen
        Drawing2D.QuickCreateSurfaceFromDC cSurface, bufferDC, True
        Drawing2D.QuickCreateSolidPen cPen, 1#, m_Colors.RetrieveColor(PDNO_Border, Me.Enabled), 100#
        PD2D.DrawRectangleF_FromRectF cSurface, cPen, m_NewItemRect
        
        Dim oldItemBorderWidth As Single
        If m_OldItemIsHovered Or ucSupport.DoIHaveFocus Then oldItemBorderWidth = 3# Else oldItemBorderWidth = 1#
        Drawing2D.QuickCreateSolidPen cPen, oldItemBorderWidth, m_Colors.RetrieveColor(PDNO_Border, Me.Enabled, , m_OldItemIsHovered Or ucSupport.DoIHaveFocus), 100#
        PD2D.DrawRectangleF_FromRectF cSurface, cPen, m_OldItemRect
        
        Set cSurface = Nothing: Set cPen = Nothing
        
    End If
    
    'Paint the final result to the screen, as relevant
    ucSupport.RequestRepaint paintImmediately
    
End Sub

'By design, PD prefers to not use design-time tooltips.  Apply tooltips at run-time, using this function.
' (IMPORTANT NOTE: translations are handled automatically.  Always pass the original English text!)
Public Sub AssignTooltip(ByRef newTooltip As String, Optional ByRef newTooltipTitle As String = vbNullString, Optional ByVal raiseTipsImmediately As Boolean = False)
    ucSupport.AssignTooltip UserControl.ContainerHwnd, newTooltip, newTooltipTitle, raiseTipsImmediately
End Sub
