VERSION 5.00
Begin VB.UserControl brushSelector 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   1710
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   ClipControls    =   0   'False
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
   HitBehavior     =   0  'None
   ScaleHeight     =   114
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   ToolboxBitmap   =   "brushSelector.ctx":0000
End
Attribute VB_Name = "brushSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Brush Selector custom control
'Copyright 2013-2015 by Tanner Helland
'Created: 30/June/15
'Last updated: 04/November/15
'Last update: convert to master UC support class; add caption support; simplify rendering approach
'
'This thin user control is basically an empty control that when clicked, displays a brush selection window.  If a
' brush is selected (e.g. Cancel is not pressed), it updates its appearance to match, and raises a "BrushChanged"
' event.
'
'Though simple, this control solves a lot of problems.  It is especially helpful for improving interaction with the
' command bar user control, as it easily supports brush reset/randomize/preset events.  It is also nice to be able
' to update a single master function for brush selection, then have the change propagate to all tool windows.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************


Option Explicit

'This control doesn't really do anything interesting, besides allow a brush to be selected.
Public Event BrushChanged()

'Because VB focus events are wonky, especially when we use CreateWindow within a UC, this control raises its own
' specialized focus events.  If you need to track focus, use these instead of the default VB functions.
Public Event GotFocusAPI()
Public Event LostFocusAPI()

'The control's current brush settings; this string is how we actually create the brush
Private m_curBrush As String

'A temporary filler object, used to render the brush preview
Private m_Filler As pdGraphicsBrush

'When the "select brush" dialog is live, this will be set to TRUE
Private isDialogLive As Boolean

'The rectangle where the brush preview is actually rendered, and a boolean to track whether the mouse is inside that rect
Private m_BrushRect As RECTF, m_MouseInsideBrushRect As Boolean

'User control support class.  Historically, many classes (and associated subclassers) were required by each user control,
' but I've since attempted to wrap these into a single master control support class.
Private WithEvents ucSupport As pdUCSupport
Attribute ucSupport.VB_VarHelpID = -1

'At present, all this control does is store a brush XML string.  This string defines all brush settings.
Public Property Get Brush() As String
    Brush = m_curBrush
End Property

Public Property Let Brush(ByVal newBrush As String)
    m_curBrush = newBrush
    RedrawBackBuffer
    RaiseEvent BrushChanged
    PropertyChanged "Brush"
End Property

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

'Outside functions can call this to force a display of the brush selection window
Public Sub DisplayBrushSelection()
    RaiseBrushDialog
End Sub

Private Sub ucSupport_ClickCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    UpdateMousePosition x, y
    If m_MouseInsideBrushRect Then RaiseBrushDialog
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
    m_MouseInsideBrushRect = Math_Functions.isPointInRectF(mouseX, mouseY, m_BrushRect)
    If m_MouseInsideBrushRect Then ucSupport.RequestCursor IDC_HAND Else ucSupport.RequestCursor IDC_DEFAULT
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

Private Sub RaiseBrushDialog()

    isDialogLive = True
    
    'Store the current brush
    Dim newBrush As String, oldBrush As String
    oldBrush = Brush
    
    'Use the brush dialog to select a new color
    If showBrushDialog(newBrush, oldBrush, Me) Then
        Brush = newBrush
    Else
        Brush = oldBrush
    End If
    
    isDialogLive = False
    
End Sub

Private Sub UserControl_Initialize()

    Set m_Filler = New pdGraphicsBrush
    
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
    Brush = ""
    Caption = ""
    FontSize = 12
End Sub

'At run-time, painting is handled by the support class.  In the IDE, however, we must rely on VB's internal paint event.
Private Sub UserControl_Paint()
    ucSupport.RequestIDERepaint UserControl.hDC
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        Brush = .ReadProperty("curBrush", "")
        Caption = .ReadProperty("Caption", "")
        FontSize = .ReadProperty("FontSize", 12)
    End With
End Sub

Private Sub UserControl_Resize()
    If Not g_IsProgramRunning Then ucSupport.RequestRepaint True
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "curBrush", m_curBrush, ""
        .WriteProperty "Caption", ucSupport.GetCaptionText, ""
        .WriteProperty "FontSize", ucSupport.GetCaptionFontSize, 12
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
        
        'The brush area is placed relative to the caption
        With m_BrushRect
            .Left = FixDPI(8)
            .Top = ucSupport.GetCaptionBottom + 2
            .Width = (bWidth - 2) - .Left
            .Height = (bHeight - 2) - .Top
        End With
        
    'If there's no caption, allow the clickable portion to fill the entire control
    Else
        
        With m_BrushRect
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
        
        'Render the brush first
        m_Filler.setBoundaryRect m_BrushRect
        m_Filler.createBrushFromString Me.Brush
        
        Dim tmpBrush As Long
        tmpBrush = m_Filler.getBrushHandle
        
        With m_BrushRect
            GDI_Plus.GDIPlusFillPatternToDC bufferDC, .Left, .Top, .Width, .Height, g_CheckerboardPattern
            GDI_Plus.GDIPlusFillDC_Brush bufferDC, tmpBrush, .Left, .Top, .Width, .Height
        End With
        
        m_Filler.releaseBrushHandle tmpBrush
        
        'Draw borders around the brush results.
        Dim outlineColor As Long, outlineWidth As Long, outlineOffset As Long
        
        If g_IsProgramRunning And m_MouseInsideBrushRect Then
            outlineColor = g_Themer.GetThemeColor(PDTC_ACCENT_DEFAULT)
            outlineWidth = 3
        Else
            outlineColor = vbBlack
            outlineWidth = 1
        End If
        
        GDI_Plus.GDIPlusDrawRectFOutlineToDC bufferDC, m_BrushRect, outlineColor, , outlineWidth, False, LineJoinMiter
        
    End If
    
    'Paint the final result to the screen, as relevant
    ucSupport.RequestRepaint
    
End Sub

'If a brush selection dialog is active, it will pass brush updates backward to this function, so that we can let
' our parent form display live updates *while the user is playing with brushes* - very cool!
Public Sub NotifyOfLiveBrushChange(ByVal newBrush As String)
    Brush = newBrush
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

