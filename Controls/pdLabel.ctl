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
'Last updated: 12/January/15
'Last update: rewrite control to handle its own caption and tooltip translations
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

'Flicker-free window painter
Private WithEvents cPainter As pdWindowPainter
Attribute cPainter.VB_VarHelpID = -1

'DPI-aware window resizer
Private WithEvents cResize As pdWindowSize
Attribute cResize.VB_VarHelpID = -1

'pdFont handles all text rendering duties.
Private curFont As pdFont

'Rather than use an StdFont container (which requires VB to create redundant font objects), we track font properties manually,
' via dedicated properties.
Private m_FontBold As Boolean
Private m_FontItalic As Boolean
Private m_FontSize As Single

'If a label caption is too long, and auto-fit layout has been specified, we must dynamically shrink the label's font size
' until an acceptable value is reached.  This variable represents the *currently in-use font size*, not the font size property.
Private m_CurFontSize As Long

'Current caption string (persistent within the IDE, but must be set at run-time for Unicode languages).  Note that m_CaptionEn
' is the ENGLISH CAPTION ONLY.  A translated caption will be stored in m_CaptionTranslated; the translated copy will be updated
' by any caption change, or by a call to UpdateAgainstCurrentTheme.
Private m_CaptionEn As String
Private m_CaptionTranslated As String

'Caption alignment
Private m_Alignment As AlignmentConstants

'Control (and caption) layout
Private m_Layout As PD_LABEL_LAYOUT

'Persistent back buffer, which we manage internally
Private m_BackBuffer As pdDIB

'If the user resizes a label, the control's back buffer needs to be redrawn.  If we resize the label as part of an internal
' AutoSize calculation, however, we will already be in the midst of resizing the backbuffer - so we override the behavior
' of the UserControl_Resize event, using this variable.
Private m_InternalResizeState As Boolean

'To further improve performance of this control, we cache back buffer repaint requests, and do not actually process them
' until a paint event is requested.  This particularly improves performance at control intialization time, because a whole bunch
' of buffer repaint requests will be generated as various properties are initialized and set, events we can coalesce until a
' paint event is actually required.
Private m_BufferDirty As Boolean

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

'Because there is sometimes a delay between updating a VB extender property (e.g. width, height) and VB actually reporting
' that property when queried, this control will manually cache its size calculations.  These values can be retrieved when
' manually aligning controls, to guarantee that any size calculations are accurate, even if VB fails to report them correctly.
Private m_ControlWidth As Long, m_ControlHeight As Long

'Additional helper for rendering themed and multiline tooltips
Private toolTipManager As pdToolTip

'Alignment is handled just like VB's internal label alignment property.
Public Property Get Alignment() As AlignmentConstants
    Alignment = m_Alignment
End Property

Public Property Let Alignment(ByVal newAlignment As AlignmentConstants)
    m_Alignment = newAlignment
    If g_IsProgramRunning Then m_BufferDirty = True Else UpdateControlSize
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal newColor As OLE_COLOR)
    If m_BackColor <> newColor Then
        m_BackColor = newColor
        If m_UseCustomBackColor Then m_BufferDirty = True
    End If
End Property

'Caption is handled just like the common control label's caption property.  It is valid at design-time, and any translation,
' if present, will not be processed until run-time.
' IMPORTANT NOTE: only the ENGLISH caption is returned.  I don't have a reason for returning a translated caption (if any),
'                  but I can revisit in the future if that ever becomes relevant.
Public Property Get Caption() As String
Attribute Caption.VB_UserMemId = -518
    Caption = m_CaptionEn
End Property

Public Property Let Caption(ByRef newCaption As String)

    If StrComp(newCaption, m_CaptionEn, vbBinaryCompare) <> 0 Then
        
        m_CaptionEn = newCaption
        
        'During run-time, apply translations as necessary
        If g_IsProgramRunning Then
        
            'See if translations are necessary.
            Dim isTranslationActive As Boolean
                
            If Not (g_Language Is Nothing) Then
                If g_Language.translationActive Then
                    isTranslationActive = True
                Else
                    isTranslationActive = False
                End If
            Else
                isTranslationActive = False
            End If
            
            'Update the translated caption accordingly
            If isTranslationActive Then
                m_CaptionTranslated = g_Language.TranslateMessage(m_CaptionEn)
            Else
                m_CaptionTranslated = m_CaptionEn
            End If
        
        Else
            m_CaptionTranslated = m_CaptionEn
        End If
        
        PropertyChanged "Caption"
        
        UpdateControlSize
        
    End If
        
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
    FontBold = m_FontBold
End Property

Public Property Let FontBold(ByVal newBoldSetting As Boolean)
    If newBoldSetting <> m_FontBold Then
        m_FontBold = newBoldSetting
        refreshFont
    End If
End Property

Public Property Get FontItalic() As Boolean
    FontItalic = m_FontItalic
End Property

Public Property Let FontItalic(ByVal newItalicSetting As Boolean)
    If newItalicSetting <> m_FontItalic Then
        m_FontItalic = newItalicSetting
        refreshFont
    End If
End Property

Public Property Get FontSize() As Single
    FontSize = m_FontSize
End Property

Public Property Let FontSize(ByVal newSize As Single)
    If newSize <> m_FontSize Then
        m_FontSize = newSize
        refreshFont
    End If
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal newColor As OLE_COLOR)
    If m_ForeColor <> newColor Then
        m_ForeColor = newColor
        If m_UseCustomForeColor Then m_BufferDirty = True
    End If
End Property

Public Property Get InternalWidth() As Long
    InternalWidth = m_ControlWidth
End Property

Public Property Get InternalHeight() As Long
    InternalHeight = m_ControlHeight
End Property

Public Property Get Layout() As PD_LABEL_LAYOUT
    Layout = m_Layout
End Property

Public Property Let Layout(ByVal newLayout As PD_LABEL_LAYOUT)
    m_Layout = newLayout
    If g_IsProgramRunning Then m_BufferDirty = True Else UpdateControlSize
End Property

'Because there can be a delay between window resize events and VB processing the related message (and updating its internal properties),
' owner windows may wish to access these read-only properties, which will return the actual control size at any given time.
Public Property Get PixelWidth() As Long
    If Not (m_BackBuffer Is Nothing) Then PixelWidth = m_BackBuffer.getDIBWidth Else PixelWidth = 0
End Property

Public Property Get PixelHeight() As Long
    If Not (m_BackBuffer Is Nothing) Then PixelHeight = m_BackBuffer.getDIBHeight Else PixelHeight = 0
End Property

Public Property Get UseCustomBackColor() As Boolean
    UseCustomBackColor = m_UseCustomBackColor
End Property

Public Property Let UseCustomBackColor(ByVal newSetting As Boolean)
    If newSetting <> m_UseCustomBackColor Then
        m_UseCustomBackColor = newSetting
        m_BufferDirty = True
    End If
End Property

Public Property Get UseCustomForeColor() As Boolean
    UseCustomForeColor = m_UseCustomForeColor
End Property

Public Property Let UseCustomForeColor(ByVal newSetting As Boolean)
    If newSetting <> m_UseCustomForeColor Then
        m_UseCustomForeColor = newSetting
        m_BufferDirty = True
    End If
End Property

'The pdWindowPaint class raises this event when the control needs to be redrawn.  The passed coordinates contain the
' rect returned by GetUpdateRect (but with right/bottom measurements pre-converted to width/height).
Private Sub cPainter_PaintWindow(ByVal winLeft As Long, ByVal winTop As Long, ByVal winWidth As Long, ByVal winHeight As Long)

    'Recreate the buffer as necessary
    If (Not m_InternalResizeState) Then
        
        If m_BufferDirty Then UpdateControlSize
        
        'Flip the relevant chunk of the buffer to the screen
        BitBlt UserControl.hDC, winLeft, winTop, winWidth, winHeight, m_BackBuffer.getDIBDC, winLeft, winTop, vbSrcCopy
        
    End If
        
End Sub

'When the font used for the label changes in some way, it can be recreated (refreshed) using this function.  Note that font
' creation is expensive, so it's worthwhile to avoid this step as much as possible.
Private Sub refreshFont()
    
    Dim fontRefreshRequired As Boolean
    fontRefreshRequired = curFont.HasFontBeenCreated
    
    'Update each font parameter in turn.  If one (or more) requires a new font object, the font will be recreated as the final step.
    
    'Font face is always set automatically, to match the current program-wide font
    If (Len(g_InterfaceFont) <> 0) And (StrComp(curFont.GetFontFace, g_InterfaceFont, vbBinaryCompare) <> 0) Then
        fontRefreshRequired = True
        curFont.SetFontFace g_InterfaceFont
    End If
    
    'In the future, I may switch to GDI+ for font rendering, as it supports floating-point font sizes.  In the meantime, we check
    ' parity using an Int() conversion, as GDI only supports integer font sizes.
    If Int(m_FontSize) <> Int(curFont.GetFontSize) Then
        fontRefreshRequired = True
        curFont.SetFontSize m_FontSize
    End If
    
    'Bold and italic are the simplest settings to handle
    If m_FontBold <> curFont.GetFontBold Then
        fontRefreshRequired = True
        curFont.SetFontBold m_FontBold
    End If
    
    If m_FontItalic <> curFont.GetFontItalic Then
        fontRefreshRequired = True
        curFont.SetFontItalic m_FontItalic
    End If
    
    'Request a new font, if one or more settings have changed
    If fontRefreshRequired Then curFont.CreateFontObject
    
    'Also, the back buffer needs to be rebuilt to reflect the new font metrics
    UpdateControlSize

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

'When the API detects a resize, update ourselves accordingly.
Private Sub cResize_WindowResize(ByVal newWidth As Long, ByVal newHeight As Long)
    If g_IsProgramRunning Then
        If (Not m_InternalResizeState) Then UpdateControlSize
    End If
End Sub

'************************************************
' START cResize BLOCK OF COPIED FUNCTIONS
'************************************************

'Note: this set of helper functions is provided to make life easier for external callers.  Callers should use these functions
'      (instead of VB's internal size/move functions) to ensure accurate size and movement changes under any screen DPI.
Public Function GetLeft() As Long
    If g_IsProgramRunning Then GetLeft = cResize.GetLeft
End Function

Public Function GetTop() As Long
    If g_IsProgramRunning Then GetTop = cResize.GetTop
End Function

Public Function GetWidth() As Long
    If g_IsProgramRunning Then GetWidth = cResize.GetWidth
End Function

Public Function GetHeight() As Long
    If g_IsProgramRunning Then GetHeight = cResize.GetHeight
End Function

Public Sub SetLeft(ByVal newLeft As Long)
    If g_IsProgramRunning Then cResize.SetPosition newLeft, cResize.GetTop
End Sub

Public Sub SetTop(ByVal newTop As Long)
    If g_IsProgramRunning Then cResize.SetPosition cResize.GetLeft, newTop
End Sub

Public Sub SetWidth(ByVal newWidth As Long)
    If g_IsProgramRunning Then cResize.SetSize newWidth, cResize.GetHeight
End Sub

Public Sub SetHeight(ByVal newHeight As Long)
    If g_IsProgramRunning Then cResize.SetSize cResize.GetWidth, newHeight
End Sub

Public Sub SetSize(ByVal newWidth As Long, ByVal newHeight As Long)
    If g_IsProgramRunning Then cResize.SetSize newWidth, newHeight
End Sub

Public Sub SetPosition(ByVal newLeft As Long, ByVal newTop As Long)
    If g_IsProgramRunning Then cResize.SetPosition newLeft, newTop
End Sub

'************************************************
' END cResize BLOCK OF COPIED FUNCTIONS
'************************************************

'INITIALIZE control
Private Sub UserControl_Initialize()
    
    'Initialize the internal font object
    Set curFont = New pdFont
    curFont.SetTextAlignment vbLeftJustify
    
    'Initialize a DPI-aware window resizer.  (Window messages won't be subclassed in the IDE, FYI)
    Set cResize = New pdWindowSize
    cResize.AttachToHWnd Me.hWnd, g_IsProgramRunning
    
    'When not in design mode, initialize a tracker for mouse events
    If g_IsProgramRunning Then
        
        'Start a flicker-free window painter
        Set cPainter = New pdWindowPainter
        cPainter.StartPainter Me.hWnd
        
        'Create a tooltip engine
        Set toolTipManager = New pdToolTip
                
    'In design mode, initialize a base theming class, so our paint function doesn't fail
    Else
        Set g_Themer = New pdVisualThemes
    End If
        
    'Note that we are not currently responsible for any resize events
    m_InternalResizeState = False
    
    'Mark the back buffer as dirty
    m_BufferDirty = True
                    
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
    
    m_FontBold = False
    m_FontItalic = False
    m_FontSize = 10
    
End Sub

'At run-time, painting is handled by PD's pdWindowPainter class.  In the IDE, however, we must rely on VB's internal paint event.
Private Sub UserControl_Paint()
    
    'Provide minimal painting within the designer
    If Not g_IsProgramRunning Then
        If m_BufferDirty Then UpdateControlSize Else redrawBackBuffer
    End If
    
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

'The control dynamically resizes each button to match the dimensions of their relative captions.
Private Sub UserControl_Resize()
    If Not g_IsProgramRunning Then
        If (Not m_InternalResizeState) Then UpdateControlSize
    End If
End Sub

'Because this control automatically forces all internal buttons to identical sizes, we have to recalculate a number
' of internal sizing metrics whenever the control size changes.
Private Sub UpdateControlSize()
    
    'Because we will be recreating the back buffer now, it is no longer dirty!
    m_BufferDirty = False
    
    'Remove our font object from the buffer DC, because we are about to recreate it
    curFont.ReleaseFromDC
    
    'Reset our back buffer, and reassign the font to it.
    If (m_BackBuffer Is Nothing) Then Set m_BackBuffer = New pdDIB
    
    'If the label caption was previously blank, and the label is set to "autosize", the user control may have dimensions (0, 0).
    ' If this happens, creating the back buffer will fail, so we must manually create a (1, 1) buffer, which will then become
    ' properly sized in subsequent render steps.
    If (cResize.GetWidth() = 0) Or (cResize.GetHeight() = 0) Or (m_BackBuffer.getDIBWidth = 0) Then
        m_BackBuffer.createBlank 1, 1, 24
    End If
    
    'Depending on the layout in use (e.g. autosize vs non-autosize), we may need to reposition the user control.
    ' Right-aligned labels in particular must have their .Left property modified, any time the .Width property is modified.
    ' To facilitate this behavior, we'll store the original label's width and height; this will let us know how far we
    ' need to move the label, if any.
    Dim origWidth As Long, origHeight As Long
    origWidth = cResize.GetWidth()
    origHeight = cResize.GetHeight()
    
    'You might think that it makes sense to wait to create our back buffer until we know what dimensions are
    ' required for an AutoSize label - and you'd be right!  However, we can't measure a string against a GDI font
    ' without first selecting the GDI font into a DC, so it's a catch-22.  Thus we create the back buffer automatically,
    ' and resize as necessary.
    
    'Start by setting the current font size to match the font size property value.
    m_CurFontSize = m_FontSize
    If m_CurFontSize <> Int(curFont.GetFontSize) Then
        curFont.SetFontSize m_CurFontSize
        curFont.CreateFontObject
    End If
    curFont.AttachToDC m_BackBuffer.getDIBDC
    
    'Different layout styles will modify the control's behavior based on the width (normal labels) or height
    ' (wordwrap labels) of the current caption
    Dim stringWidth As Long, stringHeight As Long
    
    'Each caption layout has its own considerations.  We'll handle all four possibilities in turn.
    Select Case m_Layout
    
        'Auto-fit caption requires the control caption to fit entirely within the control's boundaries, with no provision
        ' for word-wrapping.  Thus we need to find the largest possible font size that allows the caption to still fit
        ' within the current control boundaries.
        Case AutoFitCaption
            
            'Measure the font relative to the current control size
            stringWidth = curFont.GetWidthOfString(m_CaptionTranslated)
            
            'If the string does not fit within the control size, shrink the font accordingly.
            Do While (stringWidth > origWidth) And (m_CurFontSize >= 8)
                
                'Shrink the font size
                m_CurFontSize = m_CurFontSize - 1
                
                'Recreate the font
                curFont.ReleaseFromDC
                curFont.SetFontSize m_CurFontSize
                curFont.CreateFontObject
                curFont.AttachToDC m_BackBuffer.getDIBDC
                
                'Measure the new size
                stringWidth = curFont.GetWidthOfString(m_CaptionTranslated)
                
            Loop
            
            'If the font is at normal size, there is a small chance that the label will not be tall enough (vertically)
            ' to hold it.  This is due to rendering differences between Tahoma (on XP) and Segoe UI (on Vista+).  As such,
            ' perform a failsafe check on the label's height, and increase it as necessary.
            stringHeight = curFont.GetHeightOfString(m_CaptionTranslated)
            
            If (stringHeight > origHeight) Then
                
                m_InternalResizeState = True
                
                'Resize the user control.  For inexplicable reasons, setting the .Width and .Height properties works for .Width,
                ' but not for .Height (aaarrrggghhh).  Fortunately, we can work around this rather easily by using MoveWindow and
                ' forcing a repaint at run-time, and reverting to the problematic internal methods only in the IDE.
                If g_IsProgramRunning Then
                    cResize.SetSize origWidth, stringHeight
                Else
                    UserControl.Size PXToTwipsX(origWidth), PXToTwipsY(stringHeight)
                End If
                
                'Recreate the backbuffer
                If (cResize.GetWidth() <> m_BackBuffer.getDIBWidth) Or (cResize.GetHeight() <> m_BackBuffer.getDIBHeight) Then
                    curFont.ReleaseFromDC
                    m_BackBuffer.createBlank cResize.GetWidth(), cResize.GetHeight(), 24
                    curFont.AttachToDC m_BackBuffer.getDIBDC
                End If
                
                'Restore normal resize behavior
                m_InternalResizeState = False
                
            Else
            
                'Create the backbuffer if it hasn't been created before
                If (cResize.GetWidth() <> m_BackBuffer.getDIBWidth) Or (cResize.GetHeight() > m_BackBuffer.getDIBHeight) Then
                    curFont.ReleaseFromDC
                    m_BackBuffer.createBlank cResize.GetWidth(), cResize.GetHeight(), 24
                    curFont.AttachToDC m_BackBuffer.getDIBDC
                End If
                
            End If
            
            'If the caption still does not fit within the available area, set the failure state to TRUE.
            If stringWidth > cResize.GetWidth() Then
                m_FitFailure = True
            Else
                m_FitFailure = False
            End If
            
            'm_FontSize will now contain the final size of the control's font, and curFont has been updated accordingly.
            ' Proceed with rendering the control.
            
        'Same as auto-fit above, but instead of fitting the caption horizontally, we fit it vertically.
        Case AutoFitCaptionPlusWordWrap
            
            'Measure the font relative to the current control size
            stringHeight = curFont.GetHeightOfWordwrapString(m_CaptionTranslated, origWidth)
            
            'If the string does not fit within the control size, shrink the font accordingly.
            Do While (stringHeight > origHeight) And (m_CurFontSize >= 8)
                
                'Shrink the font size
                m_CurFontSize = m_CurFontSize - 1
                
                'Recreate the font
                curFont.ReleaseFromDC
                curFont.SetFontSize m_CurFontSize
                curFont.CreateFontObject
                curFont.AttachToDC m_BackBuffer.getDIBDC
                
                'Measure the new size
                stringHeight = curFont.GetHeightOfWordwrapString(m_CaptionTranslated, origWidth)
                
            Loop
            
            'Create the backbuffer if it hasn't been created before
            If (cResize.GetWidth() <> m_BackBuffer.getDIBWidth) Or (cResize.GetHeight() > m_BackBuffer.getDIBHeight) Then
                curFont.ReleaseFromDC
                m_BackBuffer.createBlank cResize.GetWidth(), cResize.GetHeight(), 24
                curFont.AttachToDC m_BackBuffer.getDIBDC
            End If
            
        'Resize the control horizontally to fit the caption, with no changes made to current font size.
        Case AutoSizeControl
        
            'Because we will likely be resizing the control as part of our calculation, we must disable the
            ' resize event's default behavior (calling this UpdateControlSize function).
            m_InternalResizeState = True
        
            'Measure the font relative to the current control size
            stringWidth = curFont.GetWidthOfString(m_CaptionTranslated)
            stringHeight = curFont.GetHeightOfString(m_CaptionTranslated)
            
            'We must make the back buffer fit the control's caption precisely.  stringWidth should be accurate;
            ' however, antialiasing may require us to add an additional pixel to the caption, in the event
            ' that ClearType is in use.
            If (stringWidth <> m_BackBuffer.getDIBWidth) Or (stringHeight <> m_BackBuffer.getDIBHeight) Then
                
                'Resize the user control.  For inexplicable reasons, setting the .Width and .Height properties works for .Width,
                ' but not for .Height (aaarrrggghhh).  Fortunately, we can work around this rather easily by using MoveWindow and
                ' forcing a repaint at run-time, and reverting to the problematic internal methods only in the IDE.
                If g_IsProgramRunning Then
                    cResize.SetSize stringWidth, stringHeight
                Else
                    With UserControl
                        .Size PXToTwipsX(stringWidth), PXToTwipsY(stringHeight)
                    End With
                End If
                
                'Recreate the backbuffer
                curFont.ReleaseFromDC
                m_BackBuffer.createBlank stringWidth, stringHeight, 24
                curFont.AttachToDC m_BackBuffer.getDIBDC
                
            End If
            
            'Restore normal resize behavior
            m_InternalResizeState = False
        
        'Resize the control vertically to fit the caption, with no changes made to current font size.
        Case AutoSizeControlPlusWordWrap
        
            'Because we will likely be resizing the control as part of our calculation, we must disable the
            ' resize event's default behavior (calling this UpdateControlSize function).
            m_InternalResizeState = True
        
            'Measure the font relative to the current control size
            stringHeight = curFont.GetHeightOfWordwrapString(m_CaptionTranslated, m_BackBuffer.getDIBWidth)
            
            'We must make the back buffer fit the control's caption precisely.  stringWidth should be accurate;
            ' however, antialiasing may require us to add an additional pixel to the caption, in the event
            ' that ClearType is in use.
            If (stringHeight <> m_BackBuffer.getDIBHeight) Then
                
                'Resize the user control.  For inexplicable reasons, setting the .Width and .Height properties works for .Width,
                ' but not for .Height (aaarrrggghhh).  Fortunately, we can work around this rather easily by using MoveWindow and
                ' forcing a repaint at run-time, and reverting to the problematic internal methods only in the IDE.
                If g_IsProgramRunning Then
                    cResize.SetSize origWidth, stringHeight
                Else
                    UserControl.Height = PXToTwipsY(stringHeight)
                End If
                
                'Recreate the backbuffer
                curFont.ReleaseFromDC
                m_BackBuffer.createBlank cResize.GetWidth(), cResize.GetHeight(), 24
                curFont.AttachToDC m_BackBuffer.getDIBDC
                
            End If
            
            'Restore normal resize behavior
            m_InternalResizeState = False
            
    End Select
    
    'If the label's caption alignment is RIGHT, and AUTOSIZE is active, we must move the LEFT property by a proportional amount
    ' to any size changes.
    If (m_Alignment = vbRightJustify) And (origWidth <> m_BackBuffer.getDIBWidth) And (m_Layout = AutoSizeControl) Then
        m_InternalResizeState = True
        If g_IsProgramRunning Then
            cResize.SetPosition UserControl.Extender.Left + (m_BackBuffer.getDIBWidth - origWidth), cResize.GetTop
        Else
            UserControl.Extender.Left = UserControl.Extender.Left + (m_BackBuffer.getDIBWidth - origWidth)
        End If
        m_InternalResizeState = False
    End If
    
    'With all size metrics handled, we can now paint the back buffer
    redrawBackBuffer
            
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    'Store all associated properties
    With PropBag
        .WriteProperty "Alignment", m_Alignment, vbLeftJustify
        .WriteProperty "BackColor", m_BackColor, vbWindowBackground
        .WriteProperty "Caption", m_CaptionEn, "caption"
        .WriteProperty "FontBold", m_FontBold, False
        .WriteProperty "FontItalic", m_FontItalic, False
        .WriteProperty "FontSize", m_FontSize, 10
        .WriteProperty "ForeColor", m_ForeColor, RGB(96, 96, 96)
        .WriteProperty "Layout", m_Layout, AutoFitCaption
        .WriteProperty "UseCustomBackColor", m_UseCustomBackColor, False
        .WriteProperty "UseCustomForeColor", m_UseCustomForeColor, False
    End With
    
End Sub

'External functions can call this to request a redraw.  This is helpful for live-updating theme settings, as in the Preferences dialog.
Public Sub UpdateAgainstCurrentTheme()
    
    If g_IsProgramRunning Then
        
        'See if translations are necessary.
        Dim isTranslationActive As Boolean
            
        If Not (g_Language Is Nothing) Then
            If g_Language.translationActive Then
                isTranslationActive = True
            Else
                isTranslationActive = False
            End If
        Else
            isTranslationActive = False
        End If
        
        'Update the translated caption accordingly
        If isTranslationActive Then
            m_CaptionTranslated = g_Language.TranslateMessage(m_CaptionEn)
        Else
            m_CaptionTranslated = m_CaptionEn
        End If
        
        'Update the current font, as necessary
        refreshFont
        
        'Cache the calculated size value
        m_ControlWidth = m_BackBuffer.getDIBWidth
        m_ControlHeight = m_BackBuffer.getDIBHeight
        
        'Force an immediate repaint.  (I think we can avoid this manual size request, because refreshFont will prompt one already
        ' if the font settings have changed in any way.)
        ' UpdateControlSize
                
    End If
    
End Sub

'Use this function to completely redraw the back buffer from scratch.  Note that this is computationally expensive compared to just flipping the
' existing buffer to the screen, so only redraw the backbuffer if the control state has somehow changed.
Private Sub redrawBackBuffer()
    
    'During initialization, this function may be called, but various needed drawing elements may not yet exist.
    ' If this happens, ignore repaint requests, obviously.
    If (m_BackBuffer Is Nothing) Or (g_Themer Is Nothing) Then
        m_BufferDirty = True
        Exit Sub
    End If
    
    'Start by erasing the back buffer
    If g_IsProgramRunning Then
        If m_UseCustomBackColor Then
            GDI_Plus.GDIPlusFillDIBRect m_BackBuffer, 0, 0, m_BackBuffer.getDIBWidth, m_BackBuffer.getDIBHeight, m_BackColor, 255
        Else
            GDI_Plus.GDIPlusFillDIBRect m_BackBuffer, 0, 0, m_BackBuffer.getDIBWidth, m_BackBuffer.getDIBHeight, g_Themer.GetThemeColor(PDTC_BACKGROUND_DEFAULT), 255
        End If
    Else
        curFont.ReleaseFromDC
        m_BackBuffer.createBlank m_BackBuffer.getDIBWidth, m_BackBuffer.getDIBHeight, 24, RGB(255, 255, 255)
        curFont.AttachToDC m_BackBuffer.getDIBDC
    End If
    
    'Colors used throughout the label's paint function are simple, and vary only by theme and control enablement
    Dim fontColor As Long
    
    If Me.Enabled Then
        If m_UseCustomForeColor Then
            fontColor = m_ForeColor
        Else
            fontColor = g_Themer.GetThemeColor(PDTC_TEXT_DEFAULT)
        End If
    Else
        fontColor = g_Themer.GetThemeColor(PDTC_DISABLED)
    End If
    
    'Pass all font settings to the font renderer
    curFont.SetFontColor fontColor
    curFont.SetTextAlignment m_Alignment
    
    'Paint the caption
    Select Case m_Layout
    
        Case AutoFitCaption, AutoSizeControl
            
            If (m_Layout = AutoFitCaption) And m_FitFailure Then
                curFont.FastRenderTextWithClipping 0, 0, m_BackBuffer.getDIBWidth, m_BackBuffer.getDIBHeight, m_CaptionTranslated, True
            Else
                curFont.FastRenderTextWithClipping 0, 0, m_BackBuffer.getDIBWidth, m_BackBuffer.getDIBHeight, m_CaptionTranslated, False
            End If
            
        Case AutoFitCaptionPlusWordWrap, AutoSizeControlPlusWordWrap
            curFont.FastRenderMultilineTextWithClipping 0, 0, m_BackBuffer.getDIBWidth, m_BackBuffer.getDIBHeight, m_CaptionTranslated
        
    End Select
    
    'Paint the buffer to the screen
    If g_IsProgramRunning Then cPainter.RequestRepaint Else BitBlt UserControl.hDC, 0, 0, cResize.GetWidth(), cResize.GetHeight(), m_BackBuffer.getDIBDC, 0, 0, vbSrcCopy

End Sub

'Post-translation, we can request an immediate refresh
Public Sub requestRefresh()
    cPainter.RequestRepaint
End Sub

'Due to complex interactions between user controls and PD's translation engine, tooltips require this dedicated function.
' (IMPORTANT NOTE: the tooltip class will handle translations automatically.  Always pass the original English text!)
Public Sub AssignTooltip(ByVal newTooltip As String, Optional ByVal newTooltipTitle As String, Optional ByVal newTooltipIcon As TT_ICON_TYPE = TTI_NONE)
    toolTipManager.setTooltip Me.hWnd, Me.containerHwnd, newTooltip, newTooltipTitle, newTooltipIcon
End Sub
