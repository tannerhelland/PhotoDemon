VERSION 5.00
Begin VB.UserControl smartCheckBox 
   BackColor       =   &H80000005&
   ClientHeight    =   375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2520
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
   MousePointer    =   99  'Custom
   ScaleHeight     =   25
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   168
   ToolboxBitmap   =   "smartCheckBox.ctx":0000
End
Attribute VB_Name = "smartCheckBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Checkbox control
'Copyright ©2013-2014 by Tanner Helland
'Created: 28/January/13
'Last updated: 30/July/14
'Last update: mirror changes from completed option button
'
'In a surprise to precisely no one, PhotoDemon has some unique needs when it comes to user controls - needs that
' the intrinsic VB controls can't handle.  These range from the obnoxious (lack of an "autosize" property for
' anything but labels) to the critical (no Unicode support).
'
'As such, I've created many of my own UCs for the program.  All are owner-drawn, with the goal of maintaining
' visual fidelity across the program, while also enabling key features like Unicode support.
'
'A few notes on this checkbox replacement, specifically:
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
Private Const WM_PAINT = &HF
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

'API technique for drawing a focus rectangle; USED ONLY FOR DEBUGGING AT PRESENT (see the Paint method for details)
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function DrawFocusRect Lib "user32" (ByVal hDC As Long, lpRect As RECT) As Long

'Previously, we used VB's internal label control to render the text caption.  This is now handled dynamically,
' via a pdFont object.
Private curFont As pdFont

'Mouse input handler
Private WithEvents cMouseEvents As pdInput
Attribute cMouseEvents.VB_VarHelpID = -1

'Subclasser for handling window messages
Private cSubclass As cSelfSubHookCallback

'An StdFont object is used to make IDE font choices persistent; note that we also need it to raise events,
' so we can track when it changes.
Private WithEvents mFont As StdFont
Attribute mFont.VB_VarHelpID = -1

'Current caption string (persistent within the IDE, but must be set at run-time for Unicode languages)
Private m_Caption As String

'Current control value
Private m_Value As CheckBoxConstants

'Persistent back buffer, which we manage internally
Private m_BackBuffer As pdDIB

'If the mouse is currently INSIDE the control, this will be set to TRUE
Private m_MouseInsideUC As Boolean

'When the option button receives focus via keyboard (e.g. NOT by mouse events), we draw a focus rect to help orient the user.
Private m_FocusRectActive As Boolean

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
    Refresh
    
End Property

'Font handling is a bit specialized for user controls; see http://msdn.microsoft.com/en-us/library/aa261313%28v=vs.60%29.aspx
Public Property Get Font() As StdFont
Attribute Font.VB_UserMemId = -512
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

Private Sub cMouseEvents_MouseEnter(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    
    If (Not m_MouseInsideUC) Or m_FocusRectActive Then
        m_MouseInsideUC = True
        Refresh
    End If
    
End Sub

Private Sub cMouseEvents_MouseLeave(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    
    If m_MouseInsideUC Then
        m_MouseInsideUC = False
        Refresh
    End If
    
End Sub

Private Sub mFont_FontChanged(ByVal PropertyName As String)
    Set UserControl.Font = mFont
End Sub

Public Property Get hWnd() As Long
Attribute hWnd.VB_UserMemId = -515
    hWnd = UserControl.hWnd
End Property

'Container hWnd must be exposed for external tooltip handling
Public Property Get containerHwnd() As Long
    containerHwnd = UserControl.containerHwnd
End Property

Public Property Get Value() As CheckBoxConstants
Attribute Value.VB_UserMemId = 0
    Value = m_Value
End Property

Public Property Let Value(ByVal newValue As CheckBoxConstants)
    
    'Update our internal value tracker
    If m_Value <> newValue Then
    
        m_Value = newValue
        PropertyChanged "Value"
        
        'Redraw the control; it's important to do this *before* raising the associated event, to maintain an impression of max responsiveness
        Refresh
        
        'Notify the user of the change by raising the CLICK event
        RaiseEvent Click
        
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
        Refresh
    End If

End Sub

Private Sub UserControl_Initialize()
    
    'Initialize the internal font object
    Set curFont = New pdFont
    curFont.setTextAlignment vbLeftJustify
    
    'When not in design mode, initialize a tracker for mouse events
    If g_UserModeFix Then
    
        Set cMouseEvents = New pdInput
        cMouseEvents.addInputTracker Me.hWnd, True, True, , True
        cMouseEvents.setSystemCursor IDC_HAND
        'cMouseEvents.requestKeyTracking Me.hWnd
        'cMouseEvents.setKeyTrackers Me.hWnd, True
        
        Set cSubclass = New cSelfSubHookCallback
        cSubclass.ssc_Subclass Me.hWnd, , , Me
        cSubclass.ssc_AddMsg Me.hWnd, MSG_BEFORE, WM_PAINT
        
    'In design mode, initialize a base theming class, so our paint function doesn't fail
    Else
        Set g_Themer = New pdVisualThemes
    End If
    
    m_MouseInsideUC = False
    m_FocusRectActive = False
    
    'Prepare a font object for use
    Set mFont = New StdFont
    Set UserControl.Font = mFont
                
End Sub

'Set default properties
Private Sub UserControl_InitProperties()
    
    Caption = "caption"
    
    Set mFont = UserControl.Font
    mFont_FontChanged ("")
    
    Value = vbChecked
    
End Sub

'Toggle the control's value upon space keypress
Private Sub UserControl_KeyPress(KeyAscii As Integer)

    If (KeyAscii = vbKeySpace) Then
        If CBool(Me.Value) Then Me.Value = vbUnchecked Else Me.Value = vbChecked
    End If

End Sub

Private Sub UserControl_LostFocus()

    'If the mouse is *not* over the user control, assume focus was set via keyboard
    If (Not m_MouseInsideUC) And m_FocusRectActive Then
        m_FocusRectActive = False
        Refresh
    End If

End Sub

'For responsiveness, MouseDown is used instead of Click
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Me.Enabled Then
        If CBool(Me.Value) Then Me.Value = vbUnchecked Else Me.Value = vbChecked
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
        Value = .ReadProperty("Value", vbChecked)
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

    Dim fontX As Long, fontY As Long
    fontX = fixDPI(32)
    fontY = 1
    
    'Calculate a precise size for the requested caption.
    Dim captionWidth As Long, captionHeight As Long, txtSize As POINTAPI
    If Not m_BackBuffer Is Nothing Then
        GetTextExtentPoint32 m_BackBuffer.getDIBDC, StrPtr(m_Caption), Len(m_Caption), txtSize
        captionWidth = txtSize.x
        captionHeight = txtSize.y
    
    'Failsafe if a Resize event is fired before we've initialized our back buffer DC
    Else
        captionWidth = fixDPI(32)
        captionHeight = fixDPI(32)
    End If
    
    'The control's size is pretty simple: an x-offset (for the selection circle), plus the size of the caption itself,
    ' and a one-pixel border around the edges.
    UserControl.Height = (fontY * 2 + captionHeight + 2) * TwipsPerPixelYFix
    UserControl.Width = (fontX + captionWidth + 2) * TwipsPerPixelXFix
    
    'Remove our font object from the buffer DC, because we are about to recreate it
    curFont.releaseFromDC
    
    'Reset our back buffer, and reassign the font to it
    Set m_BackBuffer = New pdDIB
    m_BackBuffer.createBlank UserControl.ScaleWidth, UserControl.ScaleHeight, 24
    curFont.attachToDC m_BackBuffer.getDIBDC
    
    'Redraw the control
    Refresh
            
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    'Store all associated properties
    With PropBag
        .WriteProperty "Caption", Caption, "caption"
        .WriteProperty "Value", Value, vbChecked
        .WriteProperty "Font", mFont, "Tahoma"
    End With
    
End Sub

Private Sub PaintUC()
    
    'Start by erasing the back buffer
    If g_UserModeFix Then
        GDI_Plus.GDIPlusFillDIBRect m_BackBuffer, 0, 0, m_BackBuffer.getDIBWidth, m_BackBuffer.getDIBHeight, g_Themer.getThemeColor(PDTC_BACKGROUND_DEFAULT), 255
    Else
        m_BackBuffer.createBlank m_BackBuffer.getDIBWidth, m_BackBuffer.getDIBHeight, 24, RGB(255, 255, 255)
        curFont.attachToDC m_BackBuffer.getDIBDC
    End If
    
    'Colors used throughout this paint function are determined primarily control enablement
    ' TODO: tie this into PD's central themer, instead of using custom values for this control!
    Dim chkBoxColorBorder, chkBoxColorFill As Long
    If Me.Enabled Then
        chkBoxColorBorder = g_Themer.getThemeColor(PDTC_ACCENT_NONINTERACTIVE, PDTCV_NORMAL)
        chkBoxColorFill = g_Themer.getThemeColor(PDTC_ACCENT_INTERACTIVE, PDTCV_NORMAL)
    Else
        chkBoxColorBorder = g_Themer.getThemeColor(PDTC_ACCENT_INTERACTIVE, PDTCV_DISABLED)
        chkBoxColorFill = g_Themer.getThemeColor(PDTC_ACCENT_INTERACTIVE, PDTCV_DISABLED)
    End If
    
    'If a focus rect is required (because focus was set via keyboard, not mouse), render it now.
    If m_FocusRectActive And m_MouseInsideUC Then m_FocusRectActive = False
    
    If m_FocusRectActive And Me.Enabled Then
        GDI_Plus.GDIPlusDrawRoundRect m_BackBuffer, 0, 0, m_BackBuffer.getDIBWidth, m_BackBuffer.getDIBHeight, 3, chkBoxColorFill, True, False
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
    Dim offsetX As Long, offsetY As Long, chkBoxSize As Long
    offsetX = 1 + fixDPI(2)
    offsetY = fontMetrics.tmInternalLeading + 1
    chkBoxSize = captionHeight - fontDescent
    chkBoxSize = chkBoxSize - fontMetrics.tmInternalLeading
    chkBoxSize = chkBoxSize + 1
    
    'Because GDI+ is finicky with antialiasing on odd-number circle sizes, force the circle size to the nearest even number
    If chkBoxSize Mod 2 = 1 Then
        chkBoxSize = chkBoxSize + 1
        offsetY = offsetY - 1
    End If
    
    'Draw a border for the checkbox regardless of value state
    GDI_Plus.GDIPlusDrawRectOutlineToDC m_BackBuffer.getDIBDC, offsetX, offsetY, offsetX + chkBoxSize, offsetY + chkBoxSize, chkBoxColorBorder, 255, 1
    
    'If the check box button is checked, draw a checkmark inside the border
    If CBool(m_Value) Then
        GDI_Plus.GDIPlusDrawLineToDC m_BackBuffer.getDIBDC, offsetX + 2, offsetY + (chkBoxSize \ 2), offsetX + (chkBoxSize \ 2) - 1.5, offsetY + chkBoxSize - 2.5, chkBoxColorFill, 255, fixDPI(2), True, LineCapRound
        GDI_Plus.GDIPlusDrawLineToDC m_BackBuffer.getDIBDC, offsetX + (chkBoxSize \ 2) - 1, (offsetY + chkBoxSize) - 3, (offsetX + chkBoxSize) - 2, offsetY + 2, chkBoxColorFill, 255, fixDPI(2), True, LineCapRound
    End If
    
    'Set the text color according to the mouse position, e.g. highlight the text if the mouse is over it
    If Me.Enabled Then
    
        If m_MouseInsideUC Then
            curFont.setFontColor g_Themer.getThemeColor(PDTC_TEXT_DEFAULT, PDTCV_HIGHLIGHT)
        Else
            curFont.setFontColor g_Themer.getThemeColor(PDTC_TEXT_DEFAULT, PDTCV_NORMAL)
        End If
        
    Else
        curFont.setFontColor g_Themer.getThemeColor(PDTC_TEXT_DEFAULT, PDTCV_DISABLED)
    End If
    
    'Failsafe check for designer mode
    If Not g_UserModeFix Then
        curFont.setFontColor RGB(0, 0, 0)
    End If
    
    'Render the text
    curFont.fastRenderText offsetX * 2 + chkBoxSize + fixDPI(6), 1, m_Caption
    
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
    
    'Flip the buffer to the user control
    BitBlt UserControl.hDC, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, m_BackBuffer.getDIBDC, 0, 0, vbSrcCopy
    
    'Validate the rect to prevent further WM_PAINT messages
    ValidateRect Me.hWnd, ByVal 0&
    
End Sub

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
        
    End If



' *************************************************************
' C A U T I O N   C A U T I O N   C A U T I O N   C A U T I O N
' -------------------------------------------------------------
' DO NOT ADD ANY OTHER CODE BELOW THE "END SUB" STATEMENT BELOW
'   add this warning banner to the last routine in your class
' *************************************************************
End Sub



