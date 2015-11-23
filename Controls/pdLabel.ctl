VERSION 5.00
Begin VB.UserControl pdLabel 
   BackColor       =   &H80000005&
   CanGetFocus     =   0   'False
   ClientHeight    =   690
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3945
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
   ScaleHeight     =   46
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   263
   ToolboxBitmap   =   "pdLabel.ctx":0000
End
Attribute VB_Name = "pdLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Unicode Label control
'Copyright 2014-2015 by Tanner Helland
'Created: 28/October/14
'Last updated: 02/November/15
'Last update: convert to ucSupport.  This control was a messy one, but it has the most to gain from program-level
'             font caching (vs each control maintaining its own font copy).
'
'In a surprise to precisely no one, PhotoDemon has some unique needs when it comes to user controls - needs that
' the intrinsic VB controls can't handle.  These range from the obnoxious (lack of an "autosize" property for
' anything but labels) to the critical (no Unicode support).
'
'As such, I've created many of my own UCs for the program.  All are owner-drawn, with the goal of maintaining
' visual fidelity across the program, while also enabling key features like Unicode support.
'
'A few notes on this label control, specifically:
'
' 1) This label uses an either/or system for its size: either the control is auto-sized based on caption length, or the
'    caption font is automatically shrunk until the caption can fit within the control border region.
' 2) High DPI settings are handled automatically.
' 3) By design, this control does not accept focus, and it does not raise any input-related events.
' 4) Coloration is automatically handled by PD's internal theming engine.
' 5) RTL language support is a work in progress.  I've designed the control so that RTL support can be added simply by
'    fixing some layout issues in this control, without the need to modify any control instances throughout PD.
'    However, working out any bugs is difficult without an RTL language to test, so further work has been postponed
'    for now.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'This control raises no events, by design.

'Rather than handle autosize and wordwrap separately, this control combines them into a single "Layout" property.
' All four possible layout approaches are covered by this enum.
Public Enum PD_LABEL_LAYOUT
    AutoFitCaption = 0
    AutoFitCaptionPlusWordWrap = 1
    AutoSizeControl = 2
    AutoSizeControlPlusWordWrap = 3
End Enum

#If False Then
    Private Const AutoFitCaption = 0, AutoFitCaptionPlusWordWrap = 1, AutoSizeControl = 2, AutoSizeControlPlusWordWrap = 3
#End If

'Control (and caption) layout
Private m_Layout As PD_LABEL_LAYOUT

'Normally, we let this control automatically determine its colors according to the current theme.  However, in some rare cases
' (like the pdCanvas status bar), we may want to override the automatic BackColor with a custom one.  Two variables are used
' for this: a BackColor/ForeColor property (which is normally ignored), and a boolean flag property "UseCustomBack/ForeColor".
Private m_BackColor As OLE_COLOR
Private m_UseCustomBackColor As Boolean

Private m_ForeColor As OLE_COLOR
Private m_UseCustomForeColor As Boolean

'On certain layouts, this control will try to shrink the caption to fit within the control.  If it cannot physically do it
' (because we run out of font sizes), this failure state will be set to TRUE.  When that happens, ellipses will be added to
' the control caption.
Private m_FitFailure As Boolean

'User control support class.  Historically, many classes (and associated subclassers) were required by each user control,
' but I've since attempted to wrap these into a single master control support class.
Private WithEvents ucSupport As pdUCSupport
Attribute ucSupport.VB_VarHelpID = -1

'Alignment is handled just like VB's internal label alignment property.
Public Property Get Alignment() As AlignmentConstants
    Alignment = ucSupport.GetCaptionAlignment()
End Property

Public Property Let Alignment(ByVal newAlignment As AlignmentConstants)
    ucSupport.SetCaptionAlignment newAlignment
    If (Not g_IsProgramRunning) Then updateControlLayout
    PropertyChanged "Alignment"
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal newColor As OLE_COLOR)
    If m_BackColor <> newColor Then
        m_BackColor = newColor
        redrawBackBuffer
    End If
End Property

'Caption is handled just like the common control label's caption property.  It is valid at design-time, and any translation,
' if present, will not be processed until run-time.
' IMPORTANT NOTE: only the ENGLISH caption is returned.  I don't have a reason for returning a translated caption (if any),
'                  but I can revisit in the future if that ever becomes relevant.
Public Property Get Caption() As String
Attribute Caption.VB_UserMemId = -518
    Caption = ucSupport.GetCaptionText
End Property

Public Property Let Caption(ByRef newCaption As String)
    
    ucSupport.SetCaptionText newCaption
    
    'Normally we would rely on the ucSupport class to raise redraw events for us, but this label control is a weird one,
    ' since we may need to resize the entire control when the caption changes.  As such, force an immediate layout update.
    If (Not g_IsProgramRunning) Then
        updateControlLayout
    Else
        If (m_Layout = AutoSizeControl) Or (m_Layout = AutoSizeControlPlusWordWrap) Then updateControlLayout
    End If
    
    PropertyChanged "Caption"
    
End Property

'The Enabled property is a bit unique; see http://msdn.microsoft.com/en-us/library/aa261357%28v=vs.60%29.aspx
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_UserMemId = -514
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal newValue As Boolean)
    
    UserControl.Enabled = newValue
    PropertyChanged "Enabled"
    
    'Redraw the control
    redrawBackBuffer
    
End Property

Public Property Get FontBold() As Boolean
    FontBold = ucSupport.GetCaptionFontBold
End Property

Public Property Let FontBold(ByVal newBoldSetting As Boolean)
    ucSupport.SetCaptionFontBold newBoldSetting
    PropertyChanged "FontBold"
End Property

Public Property Get FontItalic() As Boolean
    FontItalic = ucSupport.GetCaptionFontItalic
End Property

Public Property Let FontItalic(ByVal newItalicSetting As Boolean)
    ucSupport.SetCaptionFontItalic newItalicSetting
    PropertyChanged "FontItalic"
End Property

Public Property Get FontSize() As Single
    FontSize = ucSupport.GetCaptionFontSize
End Property

Public Property Let FontSize(ByVal newSize As Single)
    ucSupport.SetCaptionFontSize newSize
    PropertyChanged "FontSize"
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal newColor As OLE_COLOR)
    If m_ForeColor <> newColor Then
        m_ForeColor = newColor
        redrawBackBuffer
    End If
End Property

Public Property Get InternalWidth() As Long
    InternalWidth = ucSupport.GetBackBufferWidth
End Property

Public Property Get InternalHeight() As Long
    InternalHeight = ucSupport.GetBackBufferHeight
End Property

Public Property Get Layout() As PD_LABEL_LAYOUT
    Layout = m_Layout
End Property

Public Property Let Layout(ByVal newLayout As PD_LABEL_LAYOUT)
    m_Layout = newLayout
    updateControlLayout
End Property

'Because there can be a delay between window resize events and VB processing the related message (and updating its internal properties),
' owner windows may wish to access these read-only properties, which will return the actual control size at any given time.
Public Property Get PixelWidth() As Long
    PixelWidth = ucSupport.GetBackBufferWidth
End Property

Public Property Get PixelHeight() As Long
    PixelHeight = ucSupport.GetBackBufferHeight
End Property

Public Property Get UseCustomBackColor() As Boolean
    UseCustomBackColor = m_UseCustomBackColor
End Property

Public Property Let UseCustomBackColor(ByVal newSetting As Boolean)
    If newSetting <> m_UseCustomBackColor Then
        m_UseCustomBackColor = newSetting
        redrawBackBuffer
    End If
End Property

Public Property Get UseCustomForeColor() As Boolean
    UseCustomForeColor = m_UseCustomForeColor
End Property

Public Property Let UseCustomForeColor(ByVal newSetting As Boolean)
    If newSetting <> m_UseCustomForeColor Then
        m_UseCustomForeColor = newSetting
        redrawBackBuffer
    End If
End Property

Private Sub ucSupport_RepaintRequired(ByVal updateLayoutToo As Boolean)
    If updateLayoutToo Then updateControlLayout
    redrawBackBuffer
End Sub

'hWnds aren't exposed by default
Public Property Get hWnd() As Long
Attribute hWnd.VB_UserMemId = -515
    hWnd = UserControl.hWnd
End Property

'Container hWnd must be exposed for external tooltip handling
Public Property Get containerHwnd() As Long
    containerHwnd = UserControl.containerHwnd
End Property

Private Sub ucSupport_WindowResize(ByVal newWidth As Long, ByVal newHeight As Long)
    updateControlLayout
End Sub

'Because we sometimes do run-timre rearranging of label controls, we wrap a couple helper functions to ensure proper high-DPI support
Public Function GetLeft() As Long
    Dim controlRect As RECTL
    ucSupport.GetControlRect controlRect
    GetLeft = controlRect.Left
End Function

Public Sub SetLeft(ByVal newLeft As Long)
    ucSupport.RequestNewPosition newLeft
End Sub

Public Function GetWidth() As Long
    GetWidth = ucSupport.GetBackBufferWidth
End Function

Public Sub SetWidth(ByVal newWidth As Long)
    ucSupport.RequestNewSize newWidth, , True
End Sub

'INITIALIZE control
Private Sub UserControl_Initialize()
    
    'Initialize a master user control support class
    Set ucSupport = New pdUCSupport
    ucSupport.RegisterControl UserControl.hWnd
    ucSupport.RequestCaptionSupport False
    ucSupport.SetCaptionAutomaticPainting False
    
    'In design mode, initialize a base theming class, so our paint functions don't fail
    If g_Themer Is Nothing Then Set g_Themer = New pdVisualThemes
    
End Sub

'Set default properties
Private Sub UserControl_InitProperties()
        
    Alignment = vbLeftJustify
    Caption = "caption"
    Layout = AutoFitCaption
    
    BackColor = vbWindowBackground
    UseCustomBackColor = False
    
    ForeColor = RGB(96, 96, 96)
    UseCustomForeColor = False
    
    FontBold = False
    FontItalic = False
    FontSize = 10
    
End Sub

'At run-time, painting is handled by PD's pdWindowPainter class.  In the IDE, however, we must rely on VB's internal paint event.
Private Sub UserControl_Paint()
    ucSupport.RequestIDERepaint UserControl.hDC
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    With PropBag
        Alignment = .ReadProperty("Alignment", vbLeftJustify)
        BackColor = .ReadProperty("BackColor", vbWindowBackground)
        Caption = .ReadProperty("Caption", "caption")
        FontBold = .ReadProperty("FontBold", False)
        FontItalic = .ReadProperty("FontItalic", False)
        FontSize = .ReadProperty("FontSize", 10)
        ForeColor = .ReadProperty("ForeColor", RGB(96, 96, 96))
        Layout = .ReadProperty("Layout", AutoFitCaption)
        UseCustomBackColor = .ReadProperty("UseCustomBackColor", False)
        UseCustomForeColor = .ReadProperty("UseCustomForeColor", False)
    End With

End Sub

Private Sub UserControl_Resize()
    If Not g_IsProgramRunning Then ucSupport.RequestRepaint True
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    'Store all associated properties
    With PropBag
        .WriteProperty "Alignment", Alignment, vbLeftJustify
        .WriteProperty "BackColor", m_BackColor, vbWindowBackground
        .WriteProperty "Caption", Caption, "caption"
        .WriteProperty "FontBold", FontBold, False
        .WriteProperty "FontItalic", FontItalic, False
        .WriteProperty "FontSize", FontSize, 10
        .WriteProperty "ForeColor", m_ForeColor, RGB(96, 96, 96)
        .WriteProperty "Layout", m_Layout, AutoFitCaption
        .WriteProperty "UseCustomBackColor", m_UseCustomBackColor, False
        .WriteProperty "UseCustomForeColor", m_UseCustomForeColor, False
    End With
    
End Sub

'Because this control automatically forces all internal buttons to identical sizes, we have to recalculate a number
' of internal sizing metrics whenever the control size changes.
Private Sub updateControlLayout()
    
    'Retrieve DPI-aware control dimensions from the support class
    Dim bWidth As Long, bHeight As Long
    bWidth = ucSupport.GetBackBufferWidth
    bHeight = ucSupport.GetBackBufferHeight
    
    'Depending on the layout in use (e.g. autosize vs non-autosize), we may need to reposition the user control.
    ' Right-aligned labels in particular must have their .Left property modified, any time the .Width property is modified.
    ' To facilitate this behavior, we'll store the original label's width and height; this will let us know how far we
    ' need to move the label, if any.
    Dim controlRect As RECTL, controlWidth As Long, controlHeight As Long
    ucSupport.GetControlRect controlRect
    controlWidth = controlRect.Right - controlRect.Left
    controlHeight = controlRect.Bottom - controlRect.Top
    
    'Different layout styles will modify the control's behavior based on the width (normal labels) or height
    ' (wordwrap labels) of the current caption
    Dim stringWidth As Long, stringHeight As Long
    
    'The end goal of this process is to end up with an appropriate control size.  When auto-fitting text, this process is
    ' fairly simple; we simply want to make sure the label is tall enough for the selected font.  For autosized labels,
    ' the process is significantly more convoluted.
    
    'Each caption layout has its own considerations.  We'll handle all four possibilities in turn.
    Select Case m_Layout
    
        'Auto-fit caption requires the control caption to fit entirely within the control's boundaries, with no provision
        ' for word-wrapping.  Most of the nasty work of this case is handled by ucSupport (which wraps pdCaption).
        Case AutoFitCaption
            
            'Measure the font relative to the current control size
            ucSupport.SetCaptionWordWrap False
            stringWidth = ucSupport.GetCaptionWidth(True)
            stringHeight = ucSupport.GetCaptionHeight()
            
            'If the font is at its normal size (e.g. autofit is not required), there is a small chance that the label will
            ' still not be tall enough (vertically) to hold it.  This is due to rendering differences between system fonts
            ' on different OSes.  As such, we always perform a failsafe check on the label's height, and increase it if necessary.
            If (stringHeight > controlHeight) Then ucSupport.RequestNewSize controlWidth, stringHeight
            
            'If the caption still does not fit within the available area (because it's so damn large that we can't physically
            ' shrink the font enough to compensate), set the failure state to TRUE.
            If stringWidth > controlWidth Then
                m_FitFailure = True
            Else
                m_FitFailure = False
            End If
            
        'Identical to the auto-fit steps above, but instead of fitting the caption horizontally, we fit it vertically.
        Case AutoFitCaptionPlusWordWrap
            
            'We don't actually need to do anything here; just set the caption wordwrap state to match
            ucSupport.SetCaptionWordWrap True
            
        'Resize the control horizontally to fit the caption, with no changes made to current font size.
        Case AutoSizeControl
            
            'Measure the current caption, without autofit behavior active
            ucSupport.SetCaptionWordWrap False
            stringWidth = ucSupport.GetCaptionWidth(False)
            stringHeight = ucSupport.GetCaptionHeight(False)
            
            If stringWidth = 0 Then stringWidth = 1
            If stringHeight = 0 Then stringHeight = 1
            
            'Request a matching size from the support class.
            ucSupport.RequestNewSize stringWidth, stringHeight
            
        'Resize the control vertically to fit the caption, with no changes made to current font size.
        Case AutoSizeControlPlusWordWrap
            
            'Measure the current caption, without autofit behavior active
            ucSupport.SetCaptionWordWrap True
            stringWidth = controlWidth
            stringHeight = ucSupport.GetCaptionHeight(False)
            
            If stringWidth = 0 Then stringWidth = 1
            If stringHeight = 0 Then stringHeight = 1
            
            'Request a matching size from the support class.
            ucSupport.RequestNewSize stringWidth, stringHeight
            
    End Select
    
    'If the label's caption alignment is RIGHT, and AUTOSIZE is active, we must move the LEFT property by a proportional amount
    ' to any size changes.
    If (ucSupport.GetCaptionAlignment = vbRightJustify) And (controlWidth <> ucSupport.GetBackBufferWidth) And (m_Layout = AutoSizeControl) Then
        ucSupport.RequestNewPosition controlRect.Left + (ucSupport.GetBackBufferWidth - controlWidth), controlRect.Top
    End If
    
    'With all size metrics handled, we can now paint the back buffer
    redrawBackBuffer
            
End Sub

'Use this function to completely redraw the back buffer from scratch.  Note that this is computationally expensive compared to just flipping the
' existing buffer to the screen, so only redraw the backbuffer if the control state has somehow changed.
Private Sub redrawBackBuffer()
    
    'Because labels are so prevalent throughout the program, this function may end up being called when PD is going down.
    ' As such, we need to perform a failsafe check on the theming class.
    If g_ProgramShuttingDown Then
        If (g_Themer Is Nothing) Then Exit Sub
    End If
    
    'Retrieve DPI-aware control dimensions from the support class
    Dim bWidth As Long, bHeight As Long
    bWidth = ucSupport.GetBackBufferWidth
    bHeight = ucSupport.GetBackBufferHeight
    
    'Figure out which back color to use.  This is normally determined by theme, but individual labels also allow a custom
    ' .BackColor property.
    Dim targetColor As Long
    If m_UseCustomBackColor Then
        targetColor = m_BackColor
    Else
        If g_IsProgramRunning Then
            targetColor = g_Themer.GetThemeColor(PDTC_BACKGROUND_DEFAULT)
        Else
            targetColor = vbWhite
        End If
    End If
    
    'Request the back buffer DC, and ask the support module to erase any existing rendering for us.
    Dim bufferDC As Long
    bufferDC = ucSupport.GetBackBufferDC(True, targetColor)
    
    'Text color also varies by theme, and possibly control enablement
    If Me.Enabled Then
        If m_UseCustomForeColor Then
            targetColor = m_ForeColor
        Else
            targetColor = g_Themer.GetThemeColor(PDTC_TEXT_DEFAULT)
        End If
    Else
        targetColor = g_Themer.GetThemeColor(PDTC_DISABLED)
    End If
    
    'Paint the caption manually
    Select Case m_Layout
    
        Case AutoFitCaption
            If m_FitFailure Then
                ucSupport.PaintCaptionManually_Clipped 0, 0, ucSupport.GetBackBufferWidth, ucSupport.GetBackBufferHeight, targetColor, True, False
            Else
                ucSupport.PaintCaptionManually_Clipped 0, 0, ucSupport.GetBackBufferWidth, ucSupport.GetBackBufferHeight, targetColor, False, False
            End If
        
        Case AutoSizeControl
            ucSupport.PaintCaptionManually_Clipped 0, 0, ucSupport.GetBackBufferWidth, ucSupport.GetBackBufferHeight, targetColor, False, True
            
        Case AutoFitCaptionPlusWordWrap
            ucSupport.PaintCaptionManually_Clipped 0, 0, ucSupport.GetBackBufferWidth, ucSupport.GetBackBufferHeight, targetColor, False, False
        
        Case AutoSizeControlPlusWordWrap
            ucSupport.PaintCaptionManually_Clipped 0, 0, ucSupport.GetBackBufferWidth, ucSupport.GetBackBufferHeight, targetColor, False, True
            
    End Select
    
    'Paint the final result to the screen, as relevant
    ucSupport.RequestRepaint
    If (Not g_IsProgramRunning) Then UserControl.Refresh
    
End Sub

'External functions can call this to request a redraw.  This is helpful for live-updating theme settings, as in the Preferences dialog.
Public Sub UpdateAgainstCurrentTheme()
    
    'The support class handles most of this for us
    If g_IsProgramRunning Then ucSupport.UpdateAgainstThemeAndLanguage
    
    'If theme changes require us to redraw our control, the support class will raise additional paint events for us.
    
End Sub

'Post-translation, we can request an immediate refresh
Public Sub RequestRefresh()
    ucSupport.RequestRepaint
End Sub

'By design, PD prefers to not use design-time tooltips.  Apply tooltips at run-time, using this function.
' (IMPORTANT NOTE: translations are handled automatically.  Always pass the original English text!)
Public Sub AssignTooltip(ByVal newTooltip As String, Optional ByVal newTooltipTitle As String, Optional ByVal newTooltipIcon As TT_ICON_TYPE = TTI_NONE)
    ucSupport.AssignTooltip UserControl.containerHwnd, newTooltip, newTooltipTitle, newTooltipIcon
End Sub
