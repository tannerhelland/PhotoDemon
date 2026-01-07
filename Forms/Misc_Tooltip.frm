VERSION 5.00
Begin VB.Form tool_Tooltip 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   3015
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4560
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawStyle       =   5  'Transparent
   Enabled         =   0   'False
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
   Icon            =   "Misc_Tooltip.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   201
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   304
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
End
Attribute VB_Name = "tool_Tooltip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Tooltip replacement window
'Copyright 2016-2026 by Tanner Helland
'Created: 21/September/16
'Last updated: 21/September/16
'Last update: initial build
'
'Normal Windows tooltips leak user and GDI objects when intermixed with safe-self-subclassing techniques.  They are
' also cumbersome to theme and position, so it's easier to just write our own replacement.
'
'Note that this form contains relatively little code.  External functions are responsible for raising (and hiding)
' this window, as necessary.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'When a tooltip is first invoked, this class figures out the best place to position the tooltip window.  As part of
' that effort, it fills a matching rect where various bits of the tooltip need to be rendered (like the border, title,
' caption, and others, as necessary).
Private m_Caption As String, m_Title As String
Private m_InternalPadding As Long, m_TitlePadding As Long

'User control support class.  Historically, many classes (and associated subclassers) were required by each user control,
' but I've since wrapped these into a single central support class.
Private WithEvents ucSupport As pdUCSupport
Attribute ucSupport.VB_VarHelpID = -1

'Local list of themable colors.  This list includes all potential colors used by this class, regardless of state change
' or internal control settings.  The list is updated by calling the UpdateColorList function.
' (Note also that this list does not include variants, e.g. "BorderColor" vs "BorderColor_Hovered".  Variant values are
'  automatically calculated by the color management class, and they are retrieved by passing boolean modifiers to that
'  class, rather than treating every imaginable variant as a separate constant.)
Private Enum PDTT_COLOR_LIST
    [_First] = 0
    PDTT_Background = 0
    PDTT_Border = 1
    PDTT_Caption = 2
    [_Last] = 2
    [_Count] = 3
End Enum

'Color retrieval and storage is handled by a dedicated class; this allows us to optimize theme interactions,
' without worrying about the details locally.
Private m_Colors As pdThemeColors

Private Sub ucSupport_MouseMoveCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal timeStamp As Long)
    UserControls.HideUCTooltip True
End Sub

Private Sub ucSupport_RepaintRequired(ByVal updateLayoutToo As Boolean)
    If updateLayoutToo Then UpdateControlLayout
    RedrawBackBuffer
End Sub

Private Sub Form_Load()
    
    'Initialize a user control support class
    Set ucSupport = New pdUCSupport
    ucSupport.RegisterControl Me.hWnd, False
    ucSupport.RequestExtraFunctionality True, , , False
    
    'Prep the color manager and load default colors
    Set m_Colors = New pdThemeColors
    Dim colorCount As PDTT_COLOR_LIST: colorCount = [_Count]
    m_Colors.InitializeColorList "PDTooltip", colorCount
    UpdateAgainstCurrentTheme
    
End Sub

Private Sub Form_Resize()
    If (Not PDMain.IsProgramRunning()) Then ucSupport.RequestRepaint False
End Sub

Public Sub NotifyTooltipSettings(ByRef ttCaption As String, ByRef ttTitle As String, ByVal internalPadding As Single, ByVal titlePadding As Single)
    m_Caption = ttCaption
    m_Title = ttTitle
    m_InternalPadding = internalPadding
    m_TitlePadding = titlePadding
    RedrawBackBuffer
End Sub

'Whenever a control property changes that affects control size or layout (including internal changes, like caption adjustments),
' call this function to recalculate the control's internal layout
Private Sub UpdateControlLayout()
    
    'Retrieve DPI-aware control dimensions from the support class
    Dim bWidth As Long, bHeight As Long
    bWidth = ucSupport.GetBackBufferWidth
    bHeight = ucSupport.GetBackBufferHeight
    
    'At present, this control doesn't make any of its own rendering decisions.  Instead, it works with the size
    ' of the form as set by the UserControls module.
            
End Sub

'Primary rendering function.  Note that ucSupport handles a number of rendering duties (like maintaining a back buffer for us).
Private Sub RedrawBackBuffer()
    
    'NOTE: if a caption exists, it has already been drawn.  We just need to draw the clickable brush portion.
    If PDMain.IsProgramRunning() Then
    
        'Request the back buffer DC, and ask the support module to erase any existing rendering for us.
        Dim bufferDC As Long
        bufferDC = ucSupport.GetBackBufferDC(True, m_Colors.RetrieveColor(PDTT_Background, True))
        If (bufferDC = 0) Then Exit Sub
        
        'Start by rendering a border around the outside of the form
        Dim cSurface As pd2DSurface, cPen As pd2DPen, cBrush As pd2DBrush
        Drawing2D.QuickCreateSurfaceFromDC cSurface, bufferDC, True
        
        Drawing2D.QuickCreateSolidPen cPen, 1, m_Colors.RetrieveColor(PDTT_Border)
        PD2D.DrawRectangleF cSurface, cPen, 0, 0, ucSupport.GetBackBufferWidth - 1, ucSupport.GetBackBufferHeight - 1
        
        Set cSurface = Nothing: Set cPen = Nothing: Set cBrush = Nothing
        
        'Next, paint the title (if any)
        Dim yOffset As Long
        yOffset = m_InternalPadding
        
        Dim availableTextWidth As Long
        availableTextWidth = ucSupport.GetBackBufferWidth - m_InternalPadding * 2 + 1
        
        Dim ttFont As pdFont
        
        If (LenB(m_Title) > 0) Then
            Set ttFont = Fonts.GetMatchingUIFont(10, True)
            ttFont.AttachToDC bufferDC
            ttFont.SetFontColor m_Colors.RetrieveColor(PDTT_Caption)
            ttFont.SetTextAlignment vbLeftJustify
            ttFont.FastRenderMultilineTextWithClipping m_InternalPadding, yOffset, availableTextWidth, ucSupport.GetBackBufferHeight, m_Title, , False
            yOffset = yOffset + ttFont.GetHeightOfWordwrapString(m_Title, availableTextWidth) + m_TitlePadding
            ttFont.ReleaseFromDC
        End If
        
        'Finally, paint the tooltip itself
        If (LenB(m_Caption) > 0) Then
            Set ttFont = Fonts.GetMatchingUIFont(10, False)
            ttFont.AttachToDC bufferDC
            ttFont.SetFontColor m_Colors.RetrieveColor(PDTT_Caption)
            ttFont.SetTextAlignment vbLeftJustify
            ttFont.FastRenderMultilineTextWithClipping m_InternalPadding, yOffset, availableTextWidth, ucSupport.GetBackBufferHeight - yOffset, m_Caption, , False
            ttFont.ReleaseFromDC
        End If
        
        Set ttFont = Nothing
        
        'Paint the final result to the screen, as relevant
        ucSupport.RequestRepaint
        
    End If
    
End Sub

'Before this control does any painting, we need to retrieve relevant colors from PD's primary theming class.  Note that this
' step must also be called if/when PD's visual theme settings change.
Private Sub UpdateColorList()
    With m_Colors
        .LoadThemeColor PDTT_Background, "Background", IDE_WHITE
        .LoadThemeColor PDTT_Border, "Border", IDE_BLACK
        .LoadThemeColor PDTT_Caption, "Caption", IDE_BLACK
    End With
End Sub

'External functions can call this to request a redraw.  This is helpful for live-updating theme settings, as in the Preferences dialog.
Public Sub UpdateAgainstCurrentTheme()
    UpdateColorList
    If PDMain.IsProgramRunning() Then ucSupport.UpdateAgainstThemeAndLanguage
End Sub
