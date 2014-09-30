VERSION 5.00
Begin VB.UserControl smartOptionButton 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   555
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3735
   ClipBehavior    =   0  'None
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00404040&
   MousePointer    =   99  'Custom
   ScaleHeight     =   37
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   249
   ToolboxBitmap   =   "smartOptionButton.ctx":0000
End
Attribute VB_Name = "smartOptionButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Radio Button control
'Copyright ©2013-2014 by Tanner Helland
'Created: 28/January/13
'Last updated: 30/July/14
'Last update: rework and optimize render code
'
'In a surprise to precisely no one, PhotoDemon has some unique needs when it comes to user controls - needs that
' the intrinsic VB controls can't handle.  These range from the obnoxious (lack of an "autosize" property for
' anything but labels) to the critical (no Unicode support).
'
'As such, I've created many of my own UCs for the program.  All are owner-drawn, with the goal of maintaining
' visual fidelity across the program, while also enabling key features like Unicode support.
'
'A few notes on this radio button replacement, specifically:
'
' 1) The control is autosized based on the current font and caption.
' 2) High DPI settings are handled automatically, so do not attempt to handle this manually.
' 3) A hand cursor is automatically applied, and clicks on both the button and label are registered properly.
' 4) Coloration is automatically handled by PD's internal theming engine.
' 5) When the control receives focus via keyboard, a special focus rect is drawn.  Focus via mouse is conveyed via text glow.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'This function really only needs one event raised - Click
Public Event Click()

'Subclassing is used to better optimize the control's painting; this also requires manual validation of the control rect.
Private Const WM_PAINT As Long = &HF
Private Const WM_ERASEBKGND As Long = &H14
Private Declare Function ValidateRect Lib "user32" (ByVal targetHWnd As Long, ByRef lpRect As Any) As Long

'Retrieve the width and height of a string
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32W" (ByVal hDC As Long, ByVal lpStrPointer As Long, ByVal cbString As Long, ByRef lpSize As POINTAPI) As Long

'Retrieve specific metrics on a font (in our case, crucial for aligning the radio button against the font baseline and ascender)
Private Declare Function GetTextMetrics Lib "gdi32" Alias "GetTextMetricsA" (ByVal hDC As Long, ByRef lpMetrics As TEXTMETRIC) As Long
Private Type TEXTMETRIC
    tmHeight As Long
    tmAscent As Long
    tmDescent As Long
    tmInternalLeading As Long
    tmExternalLeading As Long
    tmAveCharWidth As Long
    tmMaxCharWidth As Long
    tmWeight As Long
    tmOverhang As Long
    tmDigitizedAspectX As Long
    tmDigitizedAspectY As Long
    tmFirstChar As Byte
    tmLastChar As Byte
    tmDefaultChar As Byte
    tmBreakChar As Byte
    tmItalic As Byte
    tmUnderlined As Byte
    tmStruckOut As Byte
    tmPitchAndFamily As Byte
    tmCharSet As Byte
End Type

'API technique for drawing a focus rectangle; used only for designer mode (see the Paint method for details)
Private Declare Function DrawFocusRect Lib "user32" (ByVal hDC As Long, lpRect As RECT) As Long

'Previously, we used VB's internal label control to render the text caption.  This is now handled dynamically,
' via a pdFont object.
Private curFont As pdFont

'Mouse input handler
Private WithEvents cMouseEvents As pdInputMouse
Attribute cMouseEvents.VB_VarHelpID = -1

'Subclasser for handling window messages
Private cSubclass As cSelfSubHookCallback

'An StdFont object is used to make IDE font choices persistent; note that we also need it to raise events,
' so we can track when it changes.
Private WithEvents mFont As StdFont
Attribute mFont.VB_VarHelpID = -1

'Current caption string (persistent within the IDE, but must be set at run-time for Unicode languages).  Note that m_Caption
' is the ENGLISH CAPTION ONLY.  A translated caption, if one exists, will be stored in m_TranslatedCaption, after PD's
' central themer invokes the translateCaption function.
Private m_Caption As String
Private m_TranslatedCaption As String

'Current control value
Private m_Value As Boolean

'Persistent back buffer, which we manage internally
Private m_BackBuffer As pdDIB

'If the mouse is currently INSIDE the control, this will be set to TRUE
Private m_MouseInsideUC As Boolean

'When the option button receives focus via keyboard (e.g. NOT by mouse events), we draw a focus rect to help orient the user.
Private m_FocusRectActive As Boolean

'Whenever the control is repainted, the clickable rect will be updated to reflect the relevant portion of the control's interior
Private clickableRect As RECT

'Additional helpers for rendering themed and multiline tooltips
Private m_ToolTip As clsToolTip
Private m_ToolString As String

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

'Font handling is a bit specialized for user controls; see http://msdn.microsoft.com/en-us/library/aa261313%28v=vs.60%29.aspx
Public Property Get Font() As StdFont
    Set Font = mFont
End Property

Public Property Set Font(mNewFont As StdFont)
    
    With mFont
        .Bold = mNewFont.Bold
        .Italic = mNewFont.Italic
        .Name = mNewFont.Name
        .Size = mNewFont.Size
    End With
    
    'Mirror all settings to our internal curFont object, then recreate it
    If Not curFont Is Nothing Then
        curFont.setFontBold mFont.Bold
        curFont.setFontFace mFont.Name
        curFont.setFontItalic mFont.Italic
        curFont.setFontSize mFont.Size
        curFont.createFontObject
    End If
    
    PropertyChanged "Font"
    
    'Redraw the control to match
    updateControlSize
    
End Property

Private Sub mFont_FontChanged(ByVal PropertyName As String)
    Set UserControl.Font = mFont
End Sub

'To improve responsiveness, MouseDown is used instead of Click
Private Sub cMouseEvents_MouseDownCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)

    'Only apply mouse events if the control is enabled, the click is in a relevant location, and the control value is not already TRUE
    If Me.Enabled And isMouseOverClickArea(x, y) And (Not Me.Value) Then
        Me.Value = True
    End If

End Sub

'When the mouse leaves the UC, we must repaint the caption (as it's no longer hovered)
Private Sub cMouseEvents_MouseLeave(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    
    If m_MouseInsideUC Then
        m_MouseInsideUC = False
        redrawBackBuffer
    End If
    
    'Reset the cursor
    cMouseEvents.setSystemCursor IDC_ARROW
    
End Sub

'When the mouse enters the clickable portion of the UC, we must repaint the caption (to reflect its hovered state)
Private Sub cMouseEvents_MouseMoveCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)

    'If the mouse is over the relevant portion of the user control, display the cursor as clickable
    If isMouseOverClickArea(x, y) Then
        
        cMouseEvents.setSystemCursor IDC_HAND
        
        'Repaint the control as necessary
        If Not m_MouseInsideUC Then
            m_MouseInsideUC = True
            redrawBackBuffer
        End If
    
    Else
    
        cMouseEvents.setSystemCursor IDC_ARROW
        
        'Repaint the control as necessary
        If m_MouseInsideUC Then
            m_MouseInsideUC = False
            redrawBackBuffer
        End If
        
    End If

End Sub

'See if the mouse is over the clickable portion of the control
Private Function isMouseOverClickArea(ByVal mouseX As Single, ByVal mouseY As Single) As Boolean
    
    If Math_Functions.isPointInRect(mouseX, mouseY, clickableRect) Then
        isMouseOverClickArea = True
    Else
        isMouseOverClickArea = False
    End If

End Function

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

'Container hWnd must be exposed for external tooltip handling
Public Property Get containerHwnd() As Long
    containerHwnd = UserControl.containerHwnd
End Property

Public Property Get Value() As Boolean
Attribute Value.VB_UserMemId = 0
    Value = m_Value
End Property

Public Property Let Value(ByVal newValue As Boolean)
    
    'Update our internal value tracker
    If m_Value <> newValue Then
    
        m_Value = newValue
        PropertyChanged "Value"
        
        'Redraw the control; it's important to do this *before* raising the associated event, to maintain an impression of max responsiveness
        redrawBackBuffer
        
        'Set all other option buttons to FALSE
        If newValue Then updateValue
        
        'If the value is being newly set to TRUE, notify the user by raising the CLICK event
        If newValue Then RaiseEvent Click
        
    End If
    
End Property

Public Property Get Caption() As String
Attribute Caption.VB_UserMemId = -518
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal newCaption As String)
    
    m_Caption = newCaption
    PropertyChanged "Caption"
    
    'Captions are a bit strange; because the control is auto-sized, changing the caption requires a full redraw
    updateControlSize
    
End Property

Private Sub UserControl_GotFocus()

    'If the mouse is *not* over the user control, assume focus was set via keyboard
    If Not m_MouseInsideUC Then
        m_FocusRectActive = True
        redrawBackBuffer
    End If

End Sub

Private Sub UserControl_Initialize()
    
    'Initialize the internal font object
    Set curFont = New pdFont
    curFont.setTextAlignment vbLeftJustify
    
    'When not in design mode, initialize a tracker for mouse events
    If g_UserModeFix Then
    
        Set cMouseEvents = New pdInputMouse
        cMouseEvents.addInputTracker Me.hWnd, True, True, , True
        cMouseEvents.setSystemCursor IDC_HAND
        
        Set cSubclass = New cSelfSubHookCallback
        cSubclass.ssc_Subclass Me.hWnd, , , Me
        cSubclass.ssc_AddMsg Me.hWnd, MSG_BEFORE, WM_PAINT, WM_ERASEBKGND
        
    'In design mode, initialize a base theming class, so our paint function doesn't fail
    Else
        Set g_Themer = New pdVisualThemes
    End If
    
    m_MouseInsideUC = False
    m_FocusRectActive = False
    
    'Prepare a font object for use
    Set mFont = New StdFont
    Set UserControl.Font = mFont
    
    'Update the control size parameters at least once
    updateControlSize
    
End Sub

'Set default properties
Private Sub UserControl_InitProperties()
    
    Caption = "caption"
    
    Set mFont = UserControl.Font
    mFont_FontChanged ("")
    
    Value = False
    
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)

    If (KeyAscii = vbKeySpace) And (Not Me.Value) Then
        Me.Value = True
    End If

End Sub

Private Sub UserControl_LostFocus()

    'If a focus rect has been drawn, remove it now
    If (Not m_MouseInsideUC) And m_FocusRectActive Then
        m_FocusRectActive = False
        redrawBackBuffer
    End If

End Sub

'Note: all drawing is done to a buffer DIB, which is flipped to the screen as the final rendering step.
' Because I don't trust VB to forward all messages correctly, WM_PAINT is manually subclassed.  See the end of this module
' for the actual painting function.
Private Sub UserControl_Paint()
    
    'Provide minimal painting within the designer
    If Not g_UserModeFix Then PaintUC
    
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    With PropBag
        Caption = .ReadProperty("Caption", "")
        Set Font = .ReadProperty("Font", Ambient.Font)
        Value = .ReadProperty("Value", False)
    End With

End Sub

'The control dynamically resizes to match the dimensions of the caption.  The size cannot be set by the user.
Private Sub UserControl_Resize()
    updateControlSize
End Sub

Private Sub UserControl_Show()

    'When the control is first made visible, remove the control's tooltip property and reassign it to the checkbox
    ' using a custom solution (which allows for linebreaks and theming).  Note that this has the ugly side-effect of
    ' permanently erasing the extender's tooltip, so FOR THIS CONTROL, TOOLTIPS MUST BE SET AT RUN-TIME!
    m_ToolString = Extender.ToolTipText

    If m_ToolString <> "" Then

        Set m_ToolTip = New clsToolTip
        With m_ToolTip

            .Create Me
            .MaxTipWidth = PD_MAX_TOOLTIP_WIDTH
            .AddTool Me, m_ToolString
            Extender.ToolTipText = ""

        End With

    End If
    
End Sub

'Whenever the size of the control changes (because the control is auto-sized, this is typically from font or caption changes),
' we must recalculate some internal rendering metrics.
Private Sub updateControlSize()

    'By adjusting this fontY parameter, we can control the auto-height of a created check box
    Dim fontY As Long
    fontY = 1
    
    'Calculate a precise size for the requested caption.
    Dim captionWidth As Long, captionHeight As Long, txtSize As POINTAPI
    
    If Not m_BackBuffer Is Nothing Then
        
        GetTextExtentPoint32 m_BackBuffer.getDIBDC, StrPtr(m_Caption), Len(m_Caption), txtSize
        captionHeight = txtSize.y
    
    'Failsafe if a Resize event is fired before we've initialized our back buffer DC
    Else
        captionHeight = fixDPI(32)
    End If
    
    'The control's size is pretty simple: an x-offset (for the selection circle), plus the size of the caption itself,
    ' and a one-pixel border around the edges.
    UserControl.Height = (fontY * 4 + captionHeight + fixDPI(2)) * TwipsPerPixelYFix
    
    'Remove our font object from the buffer DC, because we are about to recreate it
    curFont.releaseFromDC
    
    'Reset our back buffer, and reassign the font to it
    Set m_BackBuffer = New pdDIB
    m_BackBuffer.createBlank UserControl.ScaleWidth, UserControl.ScaleHeight, 24
    curFont.attachToDC m_BackBuffer.getDIBDC
    
    'Redraw the control
    redrawBackBuffer
            
End Sub

'Because this is an option control (not a checkbox), other option controls need to be turned off when it is clicked
Private Sub updateValue()

    'If the option button is set to TRUE, turn off all other option buttons on a form
    If m_Value Then

        'Enumerate through each control on the form; if it's another option button whose value is TRUE, set it to FALSE
        Dim eControl As Object
        For Each eControl In Parent.Controls
            If TypeOf eControl Is smartOptionButton Then
                If eControl.Container.hWnd = UserControl.containerHwnd Then
                    If Not (eControl.hWnd = UserControl.hWnd) Then
                        If eControl.Value Then eControl.Value = False
                    End If
                End If
            End If
        Next eControl
    
    End If
    
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    'Store all associated properties
    With PropBag
        .WriteProperty "Caption", Caption, "caption"
        .WriteProperty "Value", Value, False
        .WriteProperty "Font", mFont, "Tahoma"
    End With
    
End Sub

'External functions can call this to request a redraw.  This is helpful for live-updating theme settings, as in the Preferences dialog.
Public Sub updateAgainstCurrentTheme()
    
    Me.Font.Name = g_InterfaceFont
    curFont.setFontFace g_InterfaceFont
    curFont.createFontObject
    
    'Redraw the control to match
    updateControlSize
    
End Sub

'External functions must call this if a caption translation is required.
Public Sub translateCaption()
    
    Dim newCaption As String
    
    'Translations are active.  Retrieve a translated caption, and make sure it fits within the control.
    If g_Language.translationActive Then
    
        'Only proceed if our caption requires translation (e.g. it's non-null and non-numeric)
        If (Len(Trim(m_Caption)) > 0) And (Not IsNumeric(m_Caption)) Then
    
            'Retrieve the translated text
            newCaption = g_Language.TranslateMessage(m_Caption)
            
            'Check the size of the translated text, using the current font settings
            Dim fullControlWidth As Long
            fullControlWidth = getRadioButtonPlusCaptionWidth(newCaption)
            
            Dim curFontSize As Single
            curFontSize = mFont.Size
            
            'If the size of the caption is wider than the control itself, repeatedly shrink the font size until we
            ' find a size that fits the entire caption.
            Do While (fullControlWidth > UserControl.ScaleWidth - fixDPI(2)) And (curFontSize >= 8)
                
                'Shrink the font size
                curFontSize = curFontSize - 0.25
                curFont.setFontSize curFontSize
                
                'Recreate the font object
                curFont.releaseFromDC
                curFont.createFontObject
                curFont.attachToDC m_BackBuffer.getDIBDC
                
                'Calculate a new width
                fullControlWidth = getRadioButtonPlusCaptionWidth(newCaption)
            
            Loop
            
        Else
            newCaption = ""
        End If
    
    'If translations are not active, skip this step entirely
    Else
        newCaption = ""
    End If
    
    'Redraw the control if the caption has changed
    If StrComp(newCaption, m_TranslatedCaption, vbBinaryCompare) <> 0 Then
        
        m_TranslatedCaption = newCaption
        redrawBackBuffer
        
    End If
    
End Sub

'Use this function to completely redraw the back buffer from scratch.  Note that this is computationally expensive compared to just flipping the
' existing buffer to the screen, so only redraw the backbuffer if the control state has somehow changed.
Private Sub redrawBackBuffer()
    
    'Start by erasing the back buffer
    If g_UserModeFix Then
        GDI_Plus.GDIPlusFillDIBRect m_BackBuffer, 0, 0, m_BackBuffer.getDIBWidth, m_BackBuffer.getDIBHeight, g_Themer.getThemeColor(PDTC_BACKGROUND_DEFAULT), 255
    Else
        m_BackBuffer.createBlank m_BackBuffer.getDIBWidth, m_BackBuffer.getDIBHeight, 24, RGB(255, 255, 255)
        curFont.attachToDC m_BackBuffer.getDIBDC
    End If
    
    'Colors used throughout this paint function are determined primarily control enablement
    ' TODO: tie this into PD's central themer, instead of using custom values for this control!
    Dim optButtonColorBorder As Long, optButtonColorFill As Long
    If Me.Enabled Then
        optButtonColorBorder = g_Themer.getThemeColor(PDTC_GRAY_DEFAULT)
        optButtonColorFill = g_Themer.getThemeColor(PDTC_ACCENT_SHADOW)
    Else
        optButtonColorBorder = g_Themer.getThemeColor(PDTC_DISABLED)
        optButtonColorFill = g_Themer.getThemeColor(PDTC_DISABLED)
    End If
        
    'Next, determine the precise size of our caption, including all internal metrics.  (We need those so we can properly
    ' align the radio button with the baseline of the font and the caps (not ascender!) height.
    Dim captionWidth As Long, captionHeight As Long
    captionWidth = curFont.getWidthOfString(m_Caption)
    captionHeight = curFont.getHeightOfString(m_Caption)
    
    'Retrieve the descent of the current font.
    Dim fontDescent As Long, fontMetrics As TEXTMETRIC
    GetTextMetrics m_BackBuffer.getDIBDC, fontMetrics
    fontDescent = fontMetrics.tmDescent
    
    'From the precise font metrics, determine a radio button offset X and Y, and a radio button size.  Note that 1px is manually
    ' added as part of maintaining a 1px border around the user control as a whole.
    Dim offsetX As Long, offsetY As Long, optCircleSize As Long
    offsetX = 1 + fixDPI(2)
    offsetY = fontMetrics.tmInternalLeading + 1
    optCircleSize = captionHeight - fontDescent
    optCircleSize = optCircleSize - fontMetrics.tmInternalLeading
    optCircleSize = optCircleSize + 1
    
    'Because GDI+ is finicky with antialiasing on odd-number circle sizes, force the circle size to the nearest even number
    If (optCircleSize Mod 2) = 1 Then
        optCircleSize = optCircleSize + 1
        offsetY = offsetY - 1
    End If
    
    'Draw a border circle regardless of option button value
    GDI_Plus.GDIPlusDrawCircleToDC m_BackBuffer.getDIBDC, offsetX + optCircleSize \ 2, offsetY + optCircleSize \ 2, optCircleSize \ 2, optButtonColorBorder, 255, 1.5, True
    
    'If the option button is TRUE, draw a smaller, filled circle inside the border
    If m_Value Then
        GDI_Plus.GDIPlusDrawEllipseToDC m_BackBuffer.getDIBDC, offsetX + 3, offsetY + 3, optCircleSize - 6, optCircleSize - 6, optButtonColorFill, True
    End If
    
    'Set the text color according to the mouse position, e.g. highlight the text if the mouse is over it
    If Me.Enabled Then
    
        If m_MouseInsideUC Then
            curFont.setFontColor g_Themer.getThemeColor(PDTC_TEXT_HYPERLINK)
        Else
            curFont.setFontColor g_Themer.getThemeColor(PDTC_TEXT_DEFAULT)
        End If
        
    Else
        curFont.setFontColor g_Themer.getThemeColor(PDTC_DISABLED)
    End If
    
    'Failsafe check for designer mode
    If Not g_UserModeFix Then
        curFont.setFontColor RGB(0, 0, 0)
    End If
    
    'Render the text
    If Len(m_TranslatedCaption) > 0 Then
        curFont.fastRenderText offsetX * 2 + optCircleSize + fixDPI(6), 1, m_TranslatedCaption
    Else
        curFont.fastRenderText offsetX * 2 + optCircleSize + fixDPI(6), 1, m_Caption
    End If
    
    'Update the clickable rect using the measurements from the final render
    With clickableRect
        .Left = 0
        .Top = 0
        If Len(m_TranslatedCaption) > 0 Then
            .Right = offsetX * 2 + optCircleSize + fixDPI(6) + curFont.getWidthOfString(m_TranslatedCaption) + fixDPI(6)
        Else
            .Right = offsetX * 2 + optCircleSize + fixDPI(6) + curFont.getWidthOfString(m_Caption) + fixDPI(6)
        End If
        .Bottom = m_BackBuffer.getDIBHeight
    End With
    
    'If a focus rect is required (because focus was set via keyboard, not mouse), render it now.
    If m_FocusRectActive And m_MouseInsideUC Then m_FocusRectActive = False
    
    If m_FocusRectActive And Me.Enabled Then
        GDI_Plus.GDIPlusDrawRoundRect m_BackBuffer, 0, 0, clickableRect.Right, m_BackBuffer.getDIBHeight, 3, optButtonColorFill, True, False
    End If
    
    'In the designer, draw a focus rect around the control; this is minimal feedback required for positioning
    If Not g_UserModeFix Then
        
        Dim tmpRect As RECT
        With tmpRect
            .Left = 0
            .Top = 0
            .Right = m_BackBuffer.getDIBWidth
            .Bottom = m_BackBuffer.getDIBHeight
        End With
        
        DrawFocusRect m_BackBuffer.getDIBDC, tmpRect

    End If
        
    'Paint the buffer to the screen
    PaintUC

End Sub

Private Sub PaintUC()
    
    'Flip the buffer to the user control
    BitBlt UserControl.hDC, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, m_BackBuffer.getDIBDC, 0, 0, vbSrcCopy
    
    'Validate the rect to prevent further WM_PAINT messages
    ValidateRect Me.hWnd, ByVal 0&
    
End Sub

'Estimate the size and offset of the radio button and caption chunk of the control.  The function allows you to pass an
' arbitrary caption, which it uses to determine auto-shrinking of font size for lengthy translated captions.
Private Function getRadioButtonPlusCaptionWidth(Optional ByVal relevantCaption As String = "") As Long

    If Len(relevantCaption) = 0 Then relevantCaption = m_Caption

    'Start by retrieving caption width and height.  (Checkbox size is proportional to these values.)
    Dim captionWidth As Long, captionHeight As Long
    captionWidth = curFont.getWidthOfString(relevantCaption)
    captionHeight = curFont.getHeightOfString(relevantCaption)
    
    'Retrieve exact size metrics of the caption, as rendered in the current font
    Dim fontDescent As Long, fontMetrics As TEXTMETRIC
    GetTextMetrics m_BackBuffer.getDIBDC, fontMetrics
    fontDescent = fontMetrics.tmDescent
    
    'Using the font metrics, determine a check box offset and size.  Note that 1px is manually added as part of maintaining a
    ' 1px border around the user control as a whole (which is used for a focus rect).
    Dim offsetX As Long, offsetY As Long, optCircleSize As Long
    offsetX = 1 + fixDPI(2)
    offsetY = fontMetrics.tmInternalLeading + 1
    optCircleSize = captionHeight - fontDescent
    optCircleSize = optCircleSize - fontMetrics.tmInternalLeading
    optCircleSize = optCircleSize + 1
    
    'Because GDI+ is finicky with antialiasing on odd-number circle sizes, force the circle size to the nearest even number
    If optCircleSize Mod 2 = 1 Then
        optCircleSize = optCircleSize + 1
        offsetY = offsetY - 1
    End If
    
    'Return the determined check box size, plus a 6px extender to separate it from the caption.
    getRadioButtonPlusCaptionWidth = offsetX * 2 + optCircleSize + fixDPI(6) + captionWidth

End Function

'All events subclassed by this window are processed here.
Private Sub myWndProc(ByVal bBefore As Boolean, _
                      ByRef bHandled As Boolean, _
                      ByRef lReturn As Long, _
                      ByVal lng_hWnd As Long, _
                      ByVal uMsg As Long, _
                      ByVal wParam As Long, _
                      ByVal lParam As Long, _
                      ByRef lParamUser As Long)
'*************************************************************************************************
'* bBefore    - Indicates whether the callback is before or after the original WndProc. Usually
'*              you will know unless the callback for the uMsg value is specified as
'*              MSG_BEFORE_AFTER (both before and after the original WndProc).
'* bHandled   - In a before original WndProc callback, setting bHandled to True will prevent the
'*              message being passed to the original WndProc and (if set to do so) the after
'*              original WndProc callback.
'* lReturn    - WndProc return value. Set as per the MSDN documentation for the message value,
'*              and/or, in an after the original WndProc callback, act on the return value as set
'*              by the original WndProc.
'* lng_hWnd   - Window handle.
'* uMsg       - Message value.
'* wParam     - Message related data.
'* lParam     - Message related data.
'* lParamUser - User-defined callback parameter. Change vartype as needed (i.e., Object, UDT, etc)
'*************************************************************************************************


    If uMsg = WM_PAINT Then
        
        PaintUC
        
        'Mark the message as handled and exit
        bHandled = True
        lReturn = 0
        
    ElseIf uMsg = WM_ERASEBKGND Then
        
        bHandled = True
        lReturn = 1
        
    End If



' *************************************************************
' C A U T I O N   C A U T I O N   C A U T I O N   C A U T I O N
' -------------------------------------------------------------
' DO NOT ADD ANY OTHER CODE BELOW THE "END SUB" STATEMENT BELOW
'   add this warning banner to the last routine in your class
' *************************************************************
End Sub

