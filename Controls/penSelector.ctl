VERSION 5.00
Begin VB.UserControl penSelector 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ClipBehavior    =   0  'None
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
   ToolboxBitmap   =   "penSelector.ctx":0000
End
Attribute VB_Name = "penSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Pen Selector custom control
'Copyright 2014-2015 by Tanner Helland
'Created: 04/July/15
'Last updated: 04/November/15
'Last update: convert to master UC support class; add caption support; simplify rendering approach
'
'This thin user control is basically an empty control that when clicked, displays a pen selection window.  If a
' pen is selected (e.g. Cancel is not pressed), it updates its appearance to match, and raises a "PenChanged"
' event.
'
'Though simple, this control solves a lot of problems.  It is especially helpful for improving interaction with the
' command bar user control, as it easily supports pen reset/randomize/preset events.  It is also nice to be able
' to update a single master function for pen selection, then have the change propagate to all tool windows.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************


Option Explicit

'This control doesn't really do anything interesting, besides allow a pen to be selected.
Public Event PenChanged()

'Because VB focus events are wonky, especially when we use CreateWindow within a UC, this control raises its own
' specialized focus events.  If you need to track focus, use these instead of the default VB functions.
Public Event GotFocusAPI()
Public Event LostFocusAPI()

'The control's current pen settings
Private m_curPen As String

'A temporary pen object, used to render the pen preview
Private m_PenPreview As pdGraphicsPen

'The path used for the preview window
Private m_PreviewPath As pdGraphicsPath

'When the "select pen" dialog is live, this will be set to TRUE
Private isDialogLive As Boolean

'The rectangle where the pen preview is actually rendered, and a boolean to track whether the mouse is inside that rect
Private m_PenRect As RECTF, m_MouseInsidePenRect As Boolean

'User control support class.  Historically, many classes (and associated subclassers) were required by each user control,
' but I've since attempted to wrap these into a single master control support class.
Private WithEvents ucSupport As pdUCSupport
Attribute ucSupport.VB_VarHelpID = -1

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

Public Property Get hWnd() As Long
Attribute hWnd.VB_UserMemId = -515
    hWnd = UserControl.hWnd
End Property

'You can retrieve the pen param string (not a pdGraphicsPen object!) via this property
Public Property Get Pen() As String
    Pen = m_curPen
End Property

Public Property Let Pen(ByVal newPen As String)
    m_curPen = newPen
    RedrawBackBuffer
    RaiseEvent PenChanged
    PropertyChanged "Pen"
End Property

'Outside functions can call this to force a display of the pen selection window
Public Sub DisplayPenSelection()
    RaisePenDialog
End Sub

Private Sub ucSupport_ClickCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    UpdateMousePosition x, y
    If m_MouseInsidePenRect Then RaisePenDialog
End Sub

Private Sub ucSupport_MouseEnter(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    UpdateMousePosition x, y
    RedrawBackBuffer
End Sub

Private Sub ucSupport_MouseLeave(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    UpdateMousePosition -100, -100
    RedrawBackBuffer
End Sub

Private Sub UpdateMousePosition(ByVal mouseX As Single, ByVal mouseY As Single)
    m_MouseInsidePenRect = Math_Functions.isPointInRectF(mouseX, mouseY, m_PenRect)
    If m_MouseInsidePenRect Then ucSupport.RequestCursor IDC_HAND Else ucSupport.RequestCursor IDC_DEFAULT
End Sub

Private Sub ucSupport_MouseMoveCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    UpdateMousePosition x, y
    RedrawBackBuffer
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

Private Sub ucSupport_WindowResize(ByVal newWidth As Long, ByVal newHeight As Long)
    UpdateControlLayout
End Sub

Private Sub RaisePenDialog()

    isDialogLive = True
    
    'Backup the current pen; if the dialog is canceled, we want to restore it
    Dim newPen As String, oldPen As String
    oldPen = Pen
    
    'Use the brush dialog to select a new color
    If showPenDialog(newPen, oldPen, Me) Then
        Pen = newPen
    Else
        Pen = oldPen
    End If
    
    isDialogLive = False
    
End Sub

Private Sub UserControl_Initialize()
    
    Set m_PenPreview = New pdGraphicsPen
    Set m_PreviewPath = New pdGraphicsPath
    
    'Initialize a master user control support class
    Set ucSupport = New pdUCSupport
    ucSupport.RegisterControl UserControl.hWnd
    
    'Request some additional input functionality (custom mouse events)
    ucSupport.RequestExtraFunctionality True
    
    'Enable caption support, so we don't need an attached label
    ucSupport.RequestCaptionSupport
        
    'In design mode, initialize a base theming class, so our paint functions don't fail
    If g_Themer Is Nothing Then Set g_Themer = New pdVisualThemes
    
    'Update the control size parameters at least once
    UpdateControlLayout
    
End Sub

Private Sub UserControl_InitProperties()
    Caption = ""
    FontSize = 12
    Pen = ""
End Sub

'At run-time, painting is handled by the support class.  In the IDE, however, we must rely on VB's internal paint event.
Private Sub UserControl_Paint()
    ucSupport.RequestIDERepaint UserControl.hDC
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        Caption = .ReadProperty("Caption", "")
        FontSize = .ReadProperty("FontSize", 12)
        Pen = .ReadProperty("curPen", "")
    End With
End Sub

Private Sub UserControl_Resize()
    If Not g_IsProgramRunning Then ucSupport.RequestRepaint True
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "Caption", ucSupport.GetCaptionText, ""
        .WriteProperty "FontSize", ucSupport.GetCaptionFontSize, 12
        .WriteProperty "curPen", m_curPen, ""
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
        With m_PenRect
            .Left = FixDPI(8)
            .Top = ucSupport.GetCaptionBottom + 2
            .Width = (bWidth - 2) - .Left
            .Height = (bHeight - 2) - .Top
        End With
        
    'If there's no caption, allow the clickable portion to fill the entire control
    Else
        
        With m_PenRect
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
    If g_IsProgramRunning Then
        
        'Paint a checkerboard background first
        With m_PenRect
            GDI_Plus.GDIPlusFillPatternToDC bufferDC, .Left, .Top, .Width, .Height, g_CheckerboardPattern
        End With
        
        'Next, create a matching GDI+ pen
        m_PenPreview.createPenFromString Me.Pen
        
        Dim tmpPen As Long
        tmpPen = m_PenPreview.getPenHandle
        
        'Prep the preview path.  Note that we manually pad it to make the preview look a little prettier.
        Dim hPadding As Single, vPadding As Single
        hPadding = m_PenPreview.getPenProperty(pgps_PenWidth) * 2
        If hPadding > FixDPIFloat(12) Then hPadding = FixDPIFloat(12)
        vPadding = hPadding
        
        m_PreviewPath.resetPath
        m_PreviewPath.createSamplePathForRect m_PenRect, hPadding, vPadding
        
        m_PreviewPath.strokePathToDIB_BarePen tmpPen, , bufferDC, True
        m_PenPreview.releasePenHandle tmpPen
        
        'Draw borders around the pen results.
        Dim outlineColor As Long, outlineWidth As Long, outlineOffset As Long
        
        If g_IsProgramRunning And m_MouseInsidePenRect Then
            outlineColor = g_Themer.GetThemeColor(PDTC_ACCENT_DEFAULT)
            outlineWidth = 3
        Else
            outlineColor = vbBlack
            outlineWidth = 1
        End If
        
        GDI_Plus.GDIPlusDrawRectFOutlineToDC bufferDC, m_PenRect, outlineColor, , outlineWidth, False, LineJoinMiter
        
    End If
    
    'Paint the final result to the screen, as relevant
    ucSupport.RequestRepaint
    
End Sub

'If a pen selection dialog is active, it will pass pen updates backward to this function, so that we can let
' our parent form display live updates *while the user is playing with pens* - very cool!
Public Sub NotifyOfLivePenChange(ByVal newPen As String)
    Pen = newPen
End Sub

'External functions can call this to request a redraw.  This is helpful for live-updating theme settings, as in the Preferences dialog.
Public Sub UpdateAgainstCurrentTheme()
    If g_IsProgramRunning Then ucSupport.UpdateAgainstThemeAndLanguage
End Sub

'By design, PD prefers to not use design-time tooltips.  Apply tooltips at run-time, using this function.
' (IMPORTANT NOTE: translations are handled automatically.  Always pass the original English text!)
Public Sub AssignTooltip(ByVal newTooltip As String, Optional ByVal newTooltipTitle As String, Optional ByVal newTooltipIcon As TT_ICON_TYPE = TTI_NONE)
    ucSupport.AssignTooltip UserControl.ContainerHwnd, newTooltip, newTooltipTitle, newTooltipIcon
End Sub

