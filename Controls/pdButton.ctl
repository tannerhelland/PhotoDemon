VERSION 5.00
Begin VB.UserControl pdButton 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ClipBehavior    =   0  'None
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
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "pdButton.ctx":0000
End
Attribute VB_Name = "pdButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Generic Button control
'Copyright 2014-2026 by Tanner Helland
'Created: 19/October/14
'Last updated: 01/April/24
'Last update: raise custom drag/drop events (that the owner can respond to as they wish)
'
'In a surprise to precisely no one, PhotoDemon has some unique needs when it comes to user controls - needs that
' the intrinsic VB controls can't handle.  These range from the obnoxious (lack of an "autosize" property for
' anything but labels) to the critical (no Unicode support).
'
'As such, I've created many of my own UCs for the program.  All are owner-drawn, with the goal of maintaining
' visual fidelity across the program, while also enabling key features like Unicode support.
'
'A few notes on this generic button control, specifically:
'
' 1) Captioning is (mostly) handled by the pdCaption class, so autosizing of overlong text is supported.
' 2) High DPI settings are handled automatically.
' 3) A hand cursor is automatically applied, and clicks are returned via the Click event.
' 4) Coloration is automatically handled by PD's internal theming engine.
' 5) This button cannot be used as a checkbox, e.g. it does not set a "Value" property when clicked.  It simply raises
'     a Click() event.  This is by design.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'This control really only needs one event raised - Click
Public Event Click()

'Because VB focus events are wonky, especially when we use CreateWindow within a UC, this control raises its own
' specialized focus events.  If you need to track focus, use these instead of the default VB functions.
Public Event GotFocusAPI()
Public Event LostFocusAPI()
Public Event SetCustomTabTarget(ByVal shiftTabWasPressed As Boolean, ByRef newTargetHwnd As Long)

'As of Feb 2018, pdButton now offers an "owner-drawn" rendering mode.  This was required for the color selector
' panel on the main window, as the settings button is too short to support text display.
Public Event DrawButton(ByVal bufferDC As Long, ByVal buttonIsHovered As Boolean, ByVal ptrToRectF As Long)

'In April 2024, I added DragDrop relays (to enable custom drag/drop behavior on individual buttons).
' (Despite the name, these relays are for the underlying OLE-prefixed events, which are the only drag/drop
' events PD uses.)
Public Event CustomDragDrop(ByRef Data As DataObject, ByRef Effect As Long, ByRef Button As Integer, ByRef Shift As Integer, ByRef x As Single, ByRef y As Single)
Public Event CustomDragOver(ByRef Data As DataObject, ByRef Effect As Long, ByRef Button As Integer, ByRef Shift As Integer, ByRef x As Single, ByRef y As Single, ByRef State As Integer)
Private m_CustomDragDropEnabled As Boolean

Public Enum PDButton_RenderMode
    BRM_Normal = 0
    BRM_OwnerDrawn = 1
End Enum

#If False Then
    Private Const BRM_Normal As Long = 0, BRM_OwnerDrawn As Long = 1
#End If

Private m_RenderMode As PDButton_RenderMode

'Rect where the caption is rendered.  This is calculated by UpdateControlLayout, and it needs to be revisited if either the caption
' or button images change.
Private m_CaptionRect As RECT

'Optional button image spritesheet.  Sprites are stored vertically, in base/hover/disabled order
Private m_ImageWidth As Long, m_ImageHeight As Long, m_Images As pdDIB

'(x, y) position of the button image.  This is auto-calculated by the control.
Private btImageCoords As PointAPI

'Mouse state trackers
Private m_MouseInsideUC As Boolean, m_ButtonStateDown As Boolean

'When the control receives focus via keyboard (e.g. NOT by mouse events), we draw a focus rect to help orient the user.
Private m_FocusRectActive As Boolean

'Current back color and background color; back color is for the button, background color is for the 1px border around the button
Private m_UseCustomBackColor As Boolean, m_UseCustomBackgroundColor As Boolean
Private m_BackColor As OLE_COLOR, m_BackgroundColor As OLE_COLOR

'User control support class.  Historically, many classes (and associated subclassers) were required by each user control,
' but I've since wrapped these into a single central support class.
Private WithEvents ucSupport As pdUCSupport
Attribute ucSupport.VB_VarHelpID = -1

'Local list of themable colors.  This list includes all potential colors used by the control, regardless of state change
' or internal control settings.  The list is updated by calling the UpdateColorList function.
' (Note also that this list does not include variants, e.g. "BorderColor" vs "BorderColor_Hovered".  Variant values are
'  automatically calculated by the color management class, and they are retrieved by passing boolean modifiers to that
'  class, rather than treating every imaginable variant as a separate constant.)
Private Enum PDBUTTON_COLOR_LIST
    [_First] = 0
    PDB_Background = 0
    PDB_ButtonFill = 1
    PDB_Border = 2
    PDB_Caption = 3
    [_Last] = 3
    [_Count] = 4
End Enum

'Color retrieval and storage is handled by a dedicated class; this allows us to optimize theme interactions,
' without worrying about the details locally.
Private m_Colors As pdThemeColors

Public Function GetControlType() As PD_ControlType
    GetControlType = pdct_Button
End Function

Public Function GetControlName() As String
    GetControlName = UserControl.Extender.Name
End Function

'BackgroundColor and BackColor are different properties.  BackgroundColor should always match the color of the parent control,
' while BackColor controls the actual button fill (and can be anything you want).
Public Property Get BackgroundColor() As OLE_COLOR
    BackgroundColor = m_BackgroundColor
End Property

Public Property Let BackgroundColor(ByVal newBackColor As OLE_COLOR)
    If (m_BackgroundColor <> newBackColor) Then
        m_BackgroundColor = newBackColor
        RedrawBackBuffer
    End If
End Property

'BackColor is an important property for this control, as it may sit on other controls whose backcolor is not guaranteed in advance.
' So we can't rely on theming alone to determine this value.
Public Property Get BackColor() As OLE_COLOR
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal newBackColor As OLE_COLOR)
    If (newBackColor <> m_BackColor) Then
        m_BackColor = newBackColor
        RedrawBackBuffer
    End If
End Property

Public Property Get CustomDragDropEnabled() As Boolean
    CustomDragDropEnabled = m_CustomDragDropEnabled
End Property

Public Property Let CustomDragDropEnabled(ByVal newValue As Boolean)
    m_CustomDragDropEnabled = newValue
    If newValue Then UserControl.OLEDropMode = 1 Else UserControl.OLEDropMode = 0
End Property

Public Property Get RenderMode() As PDButton_RenderMode
    RenderMode = m_RenderMode
End Property

Public Property Let RenderMode(ByVal newRenderMode As PDButton_RenderMode)
    m_RenderMode = newRenderMode
End Property

Public Property Get UseCustomBackColor() As Boolean
    UseCustomBackColor = m_UseCustomBackColor
End Property

Public Property Let UseCustomBackColor(ByVal newSetting As Boolean)
    If (newSetting <> m_UseCustomBackColor) Then
        m_UseCustomBackColor = newSetting
        RedrawBackBuffer
    End If
End Property

Public Property Get UseCustomBackgroundColor() As Boolean
    UseCustomBackgroundColor = m_UseCustomBackgroundColor
End Property

Public Property Let UseCustomBackgroundColor(ByVal newSetting As Boolean)
    If (newSetting <> m_UseCustomBackgroundColor) Then
        m_UseCustomBackgroundColor = newSetting
        RedrawBackBuffer
    End If
End Property

'Caption is handled just like the common control label's caption property.  It is valid at design-time, and any translation,
' if present, will not be processed until run-time.
' IMPORTANT NOTE: only the ENGLISH caption is returned.  I don't have a reason for returning a translated caption (if any),
'                  but I can revisit in the future if it ever becomes relevant.
Public Property Get Caption() As String
Attribute Caption.VB_UserMemId = -518
    Caption = ucSupport.GetCaptionText
End Property

Public Property Let Caption(ByRef newCaption As String)
    
    ucSupport.SetCaptionText newCaption
    PropertyChanged "Caption"
    
    'Access keys must be handled manually.
    Dim ampPos As Long
    ampPos = InStr(1, newCaption, "&", vbBinaryCompare)
    
    If (ampPos > 0) And (ampPos < Len(newCaption)) Then
    
        'Get the character immediately following the ampersand, and dynamically assign it
        Dim accessKeyChar As String
        accessKeyChar = Mid$(newCaption, ampPos + 1, 1)
        UserControl.AccessKeys = accessKeyChar
    
    Else
        UserControl.AccessKeys = vbNullString
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
    If PDMain.IsProgramRunning() Then RedrawBackBuffer
End Property

Public Property Get FontSize() As Single
    FontSize = ucSupport.GetCaptionFontSize
End Property

Public Property Let FontSize(ByVal newSize As Single)
    ucSupport.SetCaptionFontSize newSize
    PropertyChanged "FontSize"
End Property

'hWnds aren't exposed by default
Public Property Get hWnd() As Long
Attribute hWnd.VB_UserMemId = -515
    hWnd = UserControl.hWnd
End Property

'Container hWnd must be exposed for external tooltip handling
Public Property Get ContainerHwnd() As Long
    ContainerHwnd = UserControl.ContainerHwnd
End Property

'To support high-DPI settings properly, we expose specialized move+size functions
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

'When creating owner-drawn buttons, I need to know the current caption color.  (Other color retrieval
' functions could be added in the future, if useful.)
Public Function GetCurrentCaptionColor() As Long
    GetCurrentCaptionColor = m_Colors.RetrieveColor(PDB_Caption, Me.Enabled, m_ButtonStateDown, m_MouseInsideUC)
End Function

'When the control receives focus, if the focus isn't received via mouse click, display a focus rect around the active button
Private Sub ucSupport_GotFocusAPI()
    m_FocusRectActive = True
    RedrawBackBuffer
    RaiseEvent GotFocusAPI
End Sub

Private Sub ucSupport_KeyDownSystem(ByVal Shift As ShiftConstants, ByVal whichSysKey As PD_NavigationKey, markEventHandled As Boolean)
    
    'Enter/Esc get reported directly to the system key handler.  Note that we track the return, because TRUE
    ' means the key was successfully forwarded to the relevant handler.  (If FALSE is returned, no control
    ' accepted the keypress, meaning we should forward the event down the line.)
    markEventHandled = NavKey.NotifyNavKeypress(Me, whichSysKey, Shift)
    
End Sub

'When the control loses focus, erase any focus rects it may have active
Private Sub ucSupport_LostFocusAPI()
    MakeLostFocusUIChanges
    RaiseEvent LostFocusAPI
End Sub

Private Sub MakeLostFocusUIChanges()
    
    'If a focus rect has been drawn, remove it now
    If (m_FocusRectActive Or m_ButtonStateDown Or m_MouseInsideUC) Then
        m_FocusRectActive = False
        m_ButtonStateDown = False
        m_MouseInsideUC = False
        RedrawBackBuffer
    End If
    
End Sub

'A few key events are also handled
Private Sub ucSupport_KeyDownCustom(ByVal Shift As ShiftConstants, ByVal vkCode As Long, markEventHandled As Boolean)
    
    markEventHandled = False
    
    'When space is pressed, raise a click event.
    If (vkCode = VK_SPACE) Or (vkCode = VK_RETURN) Then
        
        If (m_FocusRectActive And Me.Enabled) Then
            m_ButtonStateDown = True
            RedrawBackBuffer
            RaiseEvent Click
            markEventHandled = True
        End If
        
    End If
    
End Sub

Private Sub ucSupport_KeyUpCustom(ByVal Shift As ShiftConstants, ByVal vkCode As Long, markEventHandled As Boolean)
    
    markEventHandled = False
    
    'When space is released, redraw the button to match
    If (vkCode = VK_SPACE) Or (vkCode = VK_RETURN) Then

        If Me.Enabled Then
            m_ButtonStateDown = False
            RedrawBackBuffer
            markEventHandled = True
        End If
        
    End If

End Sub

'Note that drawing flags are handled by MouseDown/Up.  Click() is only used for raising a matching Click() event.
Private Sub ucSupport_ClickCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    If Me.Enabled And (Button = pdLeftButton) Then RaiseEvent Click
End Sub

Private Sub ucSupport_MouseDownCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal timeStamp As Long)
    If Me.Enabled And ((Button And pdLeftButton) <> 0) Then
        m_ButtonStateDown = True
        RedrawBackBuffer
    End If
End Sub

Private Sub ucSupport_MouseEnter(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    m_MouseInsideUC = True
    ucSupport.RequestCursor IDC_HAND
    RedrawBackBuffer
End Sub

'When the mouse leaves the UC, we must repaint the button (as it's no longer hovered)
Private Sub ucSupport_MouseLeave(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    If m_MouseInsideUC Then
        m_MouseInsideUC = False
        RedrawBackBuffer
    End If
    ucSupport.RequestCursor IDC_DEFAULT
End Sub

'When the mouse enters the button, we must initiate a repaint (to reflect its hovered state)
Private Sub ucSupport_MouseMoveCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal timeStamp As Long)
    If (Not m_MouseInsideUC) Then
        m_MouseInsideUC = True
        RedrawBackBuffer
    End If
End Sub

Private Sub ucSupport_MouseUpCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal clickEventAlsoFiring As Boolean, ByVal timeStamp As Long)
    If m_ButtonStateDown Then
        m_ButtonStateDown = False
        RedrawBackBuffer
    End If
End Sub

'Assign a DIB to this button.  Matching disabled and hover state DIBs are automatically generated.
' Note that you can supply an existing DIB, or a resource name.  You must supply one or the other (obviously).
' No preprocessing is currently applied to DIBs loaded as a resource.
Public Sub AssignImage(Optional ByRef resName As String = vbNullString, Optional ByRef srcDIB As pdDIB, Optional ByVal imgWidth As Long = 0, Optional ByVal imgHeight As Long = 0, Optional ByVal useCustomColor As Long = -1, Optional ByVal resampleAlgorithm As GP_InterpolationMode = GP_IM_HighQualityBicubic, Optional ByVal usePDResamplerInstead As PD_ResamplingFilter = rf_Automatic)
    
    'Load the requested resource DIB, as necessary
    If (LenB(resName) <> 0) Or (Not srcDIB Is Nothing) Then
    
        If (LenB(resName) <> 0) Then IconsAndCursors.LoadResourceToDIB resName, srcDIB, imgWidth, imgHeight, 0&, useCustomColor, False, resampleAlgorithm, usePDResamplerInstead
        
        'Cache the width and height of the DIB; it serves as our reference measurements for subsequent blt operations.
        ' (We also check for these != 0 to verify that an image was successfully loaded.)
        m_ImageWidth = srcDIB.GetDIBWidth
        m_ImageHeight = srcDIB.GetDIBHeight
        
        'Create a vertical sprite-sheet DIB, and mark it as having premultiplied alpha
        If (m_Images Is Nothing) Then Set m_Images = New pdDIB
        m_Images.CreateBlank m_ImageWidth, m_ImageHeight * 3, srcDIB.GetDIBColorDepth, 0, 0
        m_Images.SetInitialAlphaPremultiplicationState True
        
        'Copy this normal-state DIB into place at the top of the sheet
        GDI.BitBltWrapper m_Images.GetDIBDC, 0, 0, m_ImageWidth, m_ImageHeight, srcDIB.GetDIBDC, 0, 0, vbSrcCopy
        
        'Next, make a copy of the source DIB.
        Dim tmpDIB As pdDIB
        Set tmpDIB = New pdDIB
        tmpDIB.CreateFromExistingDIB srcDIB
        
        'Convert this to a brighter, "glowing" version; we'll use this when rendering a hovered state.
        Filters_Layers.ScaleDIBRGBValues tmpDIB, UC_HOVER_BRIGHTNESS
        
        'Copy this DIB into position #2, beneath the base DIB
        GDI.BitBltWrapper m_Images.GetDIBDC, 0, m_ImageHeight, m_ImageWidth, m_ImageHeight, tmpDIB.GetDIBDC, 0, 0, vbSrcCopy
        
        'Finally, create a grayscale copy of the original image.  This will serve as the "disabled state" copy.
        tmpDIB.CreateFromExistingDIB srcDIB
        Filters_Layers.GrayscaleDIB tmpDIB
        
        'Place it into position #3, beneath the previous two DIBs
        GDI.BitBltWrapper m_Images.GetDIBDC, 0, m_ImageHeight * 2, m_ImageWidth, m_ImageHeight, tmpDIB.GetDIBDC, 0, 0, vbSrcCopy
        
        'Free whatever DIBs we can.  (If the caller passed us the source DIB, we trust them to release it.)
        Set tmpDIB = Nothing
        If (LenB(resName) <> 0) Then Set srcDIB = Nothing
        m_Images.SuspendDIB
        
    'If no DIB is provided, remove any existing images
    Else
        Set m_Images = Nothing
    End If
    
    'Request a control size update, which will also calculate a centered position for the new image
    UpdateControlLayout

End Sub

Private Sub ucSupport_RepaintRequired(ByVal updateLayoutToo As Boolean)
    If updateLayoutToo Then UpdateControlLayout Else RedrawBackBuffer
End Sub

Private Sub ucSupport_SetCustomTabTarget(ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    RaiseEvent SetCustomTabTarget(shiftTabWasPressed, newTargetHwnd)
End Sub

Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
    RaiseEvent Click
End Sub

Private Sub UserControl_Initialize()
    
    'Initialize a user control support class
    Set ucSupport = New pdUCSupport
    ucSupport.RegisterControl UserControl.hWnd, True
    ucSupport.RequestExtraFunctionality True, True
    ucSupport.SpecifyRequiredKeys VK_SPACE, VK_RETURN
    
    ucSupport.RequestCaptionSupport True
    ucSupport.SetCaptionAutomaticPainting False
    ucSupport.SetCaptionAlignment vbCenter
    
    m_MouseInsideUC = False
    m_FocusRectActive = False
    
    'Prep the color manager and load default colors
    Set m_Colors = New pdThemeColors
    Dim colorCount As PDBUTTON_COLOR_LIST: colorCount = [_Count]
    m_Colors.InitializeColorList "PDButton", colorCount
    If Not PDMain.IsProgramRunning() Then UpdateColorList
         
End Sub

Private Sub UserControl_InitProperties()
    BackColor = vbWhite
    BackgroundColor = vbWhite
    Caption = vbNullString
    CustomDragDropEnabled = False
    Enabled = True
    FontSize = 10
    RenderMode = BRM_Normal
    UseCustomBackColor = False
    UseCustomBackgroundColor = False
End Sub

Private Sub UserControl_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent CustomDragDrop(Data, Effect, Button, Shift, x, y)
End Sub

Private Sub UserControl_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
    RaiseEvent CustomDragOver(Data, Effect, Button, Shift, x, y, State)
End Sub

'At run-time, painting is handled by PD's pdWindowPainter class.  In the IDE, however, we must rely on VB's internal paint event.
Private Sub UserControl_Paint()
    If (Not PDMain.IsProgramRunning()) Then ucSupport.RequestIDERepaint UserControl.hDC
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        m_BackColor = .ReadProperty("BackColor", vbWhite)
        m_BackgroundColor = .ReadProperty("BackgroundColor", vbWhite)
        Caption = .ReadProperty("Caption", vbNullString)
        CustomDragDropEnabled = .ReadProperty("CustomDragDropEnabled", False)
        Enabled = .ReadProperty("Enabled", True)
        FontSize = .ReadProperty("FontSize", 10)
        RenderMode = .ReadProperty("RenderMode", 0)
        m_UseCustomBackColor = .ReadProperty("UseCustomBackColor", False)
        m_UseCustomBackgroundColor = .ReadProperty("UseCustomBackgroundColor", False)
    End With
End Sub

Private Sub UserControl_Resize()
    If (Not PDMain.IsProgramRunning()) Then ucSupport.NotifyIDEResize UserControl.Width, UserControl.Height
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "BackColor", m_BackColor, vbWhite
        .WriteProperty "BackgroundColor", m_BackgroundColor, vbWhite
        .WriteProperty "Caption", ucSupport.GetCaptionText, vbNullString
        .WriteProperty "CustomDragDropEnabled", Me.CustomDragDropEnabled, False
        .WriteProperty "Enabled", Me.Enabled, True
        .WriteProperty "FontSize", ucSupport.GetCaptionFontSize, 10
        .WriteProperty "RenderMode", Me.RenderMode, 0
        .WriteProperty "UseCustomBackColor", m_UseCustomBackColor, False
        .WriteProperty "UseCustomBackgroundColor", m_UseCustomBackgroundColor, False
    End With
End Sub

'Because this control automatically forces all internal buttons to identical sizes, we have to recalculate a number
' of internal sizing metrics whenever the control size changes.
Private Sub UpdateControlLayout()
    
    'Retrieve DPI-aware control dimensions from the support class
    Dim bWidth As Long, bHeight As Long
    bWidth = ucSupport.GetBackBufferWidth
    bHeight = ucSupport.GetBackBufferHeight
    
    'Next, we need to determine the positioning of the caption and/or image.  Both (or neither) of these may be missing, so handling
    ' can get a little complicated.
    
    'Start with the caption
    If ucSupport.IsCaptionActive Then
        
        'We need to find the available area for the caption.  The caption class will automatically fit any translated text inside
        ' this rect.
        Const hTextPadding As Long = 8&, vTextPadding As Long = 4&
        
        'The y-positioning of the caption is always constant
        m_CaptionRect.Top = vTextPadding
        m_CaptionRect.Bottom = bHeight - vTextPadding
        
        'Similarly, the right bound doesn't change either
        m_CaptionRect.Right = bWidth - hTextPadding
        
        'If a button image is active, forcibly calculate its position first.  Its position is hard-coded.
        If (Not m_Images Is Nothing) Then
        
            Const leftButtonPadding As Long = 12&
            btImageCoords.x = FixDPI(leftButtonPadding)
            btImageCoords.y = (bHeight - m_ImageHeight) \ 2
            
            'Use the button's position to calculate an x-coord for the caption, too
            m_CaptionRect.Left = btImageCoords.x + m_ImageWidth + hTextPadding
                    
        Else
            m_CaptionRect.Left = hTextPadding
        End If
        
        'Notify the support class of the caption's boundary rect.  It will use this to auto-fit the caption font.
        With m_CaptionRect
            ucSupport.SetCaptionCustomPosition .Left, .Top, .Right - .Left, .Bottom - .Top
        End With
        
    'If there's no caption, center the button image on the control
    Else
        
        'Determine positioning of the button image, if any
        If (Not m_Images Is Nothing) Then
            btImageCoords.x = (bWidth - m_ImageWidth) \ 2
            btImageCoords.y = (bHeight - m_ImageHeight) \ 2
        End If
        
    End If
    
    'No other special preparation is required for this control, so proceed with recreating the back buffer
    RedrawBackBuffer
            
End Sub

'Use this function to completely redraw the back buffer from scratch.  Note that this is computationally expensive compared to just flipping the
' existing buffer to the screen, so only redraw the backbuffer if the control state has somehow changed.
Private Sub RedrawBackBuffer()
    
    'Retrieve DPI-aware control dimensions from the support class
    Dim bWidth As Long, bHeight As Long
    bWidth = ucSupport.GetBackBufferWidth
    bHeight = ucSupport.GetBackBufferHeight
    
    'Figure out which background color to use.  This is normally determined by theme, but individual buttons also allow
    ' a custom .BackColor property (important if this instance lies atop a non-standard background, like a command bar).
    Dim targetColor As Long
    If m_UseCustomBackgroundColor Then
        targetColor = m_BackgroundColor
    Else
        targetColor = m_Colors.RetrieveColor(PDB_Background, Me.Enabled)
    End If
    
    'Request the back buffer DC, and ask the support module to erase any existing rendering for us.
    Dim bufferDC As Long
    bufferDC = ucSupport.GetBackBufferDC(True, targetColor)
    If (bufferDC = 0) Then Exit Sub
    
    'Colors used throughout this paint function are determined by several factors:
    ' 1) Control enablement (disabled buttons are grayed)
    ' 2) Hover state (hovered buttons glow)
    ' 3) Value (pressed buttons have a different appearance, obviously)
    ' 4) The central themer (which contains default color values for all these scenarios)
    Dim btnColorBorder As Long, btnColorFill As Long, btnColorText As Long
    btnColorBorder = m_Colors.RetrieveColor(PDB_Border, Me.Enabled, m_ButtonStateDown, m_MouseInsideUC Or m_FocusRectActive)
    btnColorText = m_Colors.RetrieveColor(PDB_Caption, Me.Enabled, m_ButtonStateDown, m_MouseInsideUC)
    
    If m_UseCustomBackColor Then
        btnColorFill = m_BackColor
    Else
        btnColorFill = m_Colors.RetrieveColor(PDB_ButtonFill, Me.Enabled, m_ButtonStateDown, m_MouseInsideUC)
    End If
    
    If PDMain.IsProgramRunning() Then
        
        'First, we fill the button interior with the established fill color
        Dim cSurface As pd2DSurface: Set cSurface = New pd2DSurface
        cSurface.SetSurfaceAntialiasing P2_AA_None
        cSurface.WrapSurfaceAroundDC bufferDC
        
        Dim cBrush As pd2DBrush: Set cBrush = New pd2DBrush
        cBrush.SetBrushColor btnColorFill
        
        PD2D.FillRectangleI cSurface, cBrush, 1, 1, bWidth - 2, bHeight - 2
        
        'A border is always drawn around the control; its size varies by hover state.  (This is standard Win 10 behavior.)
        Dim borderWidth As Single
        If m_MouseInsideUC Or m_FocusRectActive Then borderWidth = 3! Else borderWidth = 1!
        
        Dim cPen As pd2DPen: Set cPen = New pd2DPen
        cPen.SetPenWidth borderWidth
        cPen.SetPenColor btnColorBorder
        cPen.SetPenLineJoin P2_LJ_Miter
        
        PD2D.DrawRectangleI_AbsoluteCoords cSurface, cPen, 1, 1, bWidth - 2, bHeight - 2
        Set cSurface = Nothing
    
        'Paint the image, if any
        If (Not m_Images Is Nothing) Then
            
            'Determine which image from the spritesheet to use.  (This is just a pixel offset.)
            Dim pxOffset As Long
            If Me.Enabled Then
                If m_MouseInsideUC Then pxOffset = m_ImageHeight Else pxOffset = 0
            Else
                pxOffset = m_ImageHeight * 2
            End If
            
            m_Images.AlphaBlendToDCEx bufferDC, btImageCoords.x, btImageCoords.y, m_ImageWidth, m_ImageHeight, 0, pxOffset, m_ImageWidth, m_ImageHeight
            m_Images.SuspendDIB
            
        End If
        
    End If
    
    'Paint the caption, if any
    If ucSupport.IsCaptionActive Then
        With m_CaptionRect
            ucSupport.PaintCaptionManually_Clipped .Left, .Top, .Right - .Left, .Bottom - .Top, btnColorText, True, False, True
        End With
    End If
    
    'If owner-drawn rendering is active, raise the corresponding event now
    If (m_RenderMode = BRM_OwnerDrawn) Then
        RaiseEvent DrawButton(bufferDC, m_MouseInsideUC Or m_FocusRectActive, VarPtr(m_CaptionRect))
    End If
    
    'Paint the final result to the screen, as relevant
    ucSupport.RequestRepaint
    If (Not PDMain.IsProgramRunning()) Then UserControl.Refresh
    
End Sub

'Before this control does any painting, we need to retrieve relevant colors from PD's primary theming class.  Note that this
' step must also be called if/when PD's visual theme settings change.
Private Sub UpdateColorList()
    With m_Colors
        .LoadThemeColor PDB_Background, "Background", IDE_WHITE
        .LoadThemeColor PDB_ButtonFill, "ButtonFill", IDE_WHITE
        .LoadThemeColor PDB_Border, "Border", IDE_BLACK
        .LoadThemeColor PDB_Caption, "Caption", IDE_BLUE
    End With
End Sub

'External functions can call this to request a redraw.  This is helpful for live-updating theme settings, as in the Preferences dialog.
Public Sub UpdateAgainstCurrentTheme(Optional ByVal hostFormhWnd As Long = 0)
    If ucSupport.ThemeUpdateRequired Then
        UpdateColorList
        If PDMain.IsProgramRunning() Then NavKey.NotifyControlLoad Me, hostFormhWnd
        If PDMain.IsProgramRunning() Then ucSupport.UpdateAgainstThemeAndLanguage
    End If
End Sub

'By design, PD prefers to not use design-time tooltips.  Apply tooltips at run-time, using this function.
' (IMPORTANT NOTE: translations are handled automatically.  Always pass the original English text!)
Public Sub AssignTooltip(ByRef newTooltip As String, Optional ByRef newTooltipTitle As String = vbNullString, Optional ByVal raiseTipsImmediately As Boolean = False)
    ucSupport.AssignTooltip UserControl.ContainerHwnd, newTooltip, newTooltipTitle, raiseTipsImmediately
End Sub
