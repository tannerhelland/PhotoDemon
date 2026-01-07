VERSION 5.00
Begin VB.UserControl pdColorWheel 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   1950
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2070
   DrawStyle       =   5  'Transparent
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
   ScaleHeight     =   130
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   138
   ToolboxBitmap   =   "pdColorWheel.ctx":0000
End
Attribute VB_Name = "pdColorWheel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon "Color Wheel" color selector
'Copyright 2015-2026 by Tanner Helland
'Created: 19/October/15
'Last updated: 29/April/24
'Last update: better behavior on high DPI systems
'
'In 7.0, a "color selector" panel was added to the right-side toolbar.  Unlike PD's single-color color selector,
' this control is designed to provide a quick, on-canvas-friendly mechanism for rapidly switching colors.  The basic
' design owes much to other photo editors like MyPaint, who pioneered various "wheel" UIs for hue selection.
'
'I've designed the control as a UC in case I decide to reuse it elsewhere in PD, but for now, it only makes an
' appearance on the main canvas.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Just like PD's old color selector, this control will raise a ColorChanged event after user interactions.
Public Event ColorChanged(ByVal newColor As Long, ByVal srcIsInternal As Boolean)

'Because VB focus events are wonky, especially when we use CreateWindow within a UC, this control raises its own
' specialized focus events.  If you need to track focus, use these instead of the default VB functions.
Public Event GotFocusAPI()
Public Event LostFocusAPI()

'Individual UI components are rendered to their own DIBs, and composited only when necessary.  For some elements
' (particularly the hue wheel), creating them from scratch is costly, so reuse is advisable.
Private m_WheelBuffer As pdDIB, m_SquareBuffer As pdDIB

'These values help the central renderer know where the mouse is, so we can draw various indicators.
Private m_MouseInsideWheel As Boolean, m_MouseInsideBox As Boolean
Private m_MouseDownWheel As Boolean, m_MouseDownBox As Boolean

'Padding (in pixels) between the edges of the user control and the color wheel.  Automatically adjusted for DPI
' at run-time.  Note that this needs to be non-zero, because the padding area is used to render the "slice" overlay
' showing the user's current hue selection.
Private Const WHEEL_PADDING As Long = 2

'Width (in pixels) of the hue wheel.  This width is applied along the radial axis.
Private Const WHEEL_WIDTH_DEFAULT As Single = 15!
Private m_WheelWidth As Single

'Various hue wheel positioning values.  These are calculated by the CreateColorWheel function and cached here, as a convenience
' for subsequent hit-testing and rendering.
Private m_HueWheelCenterX As Single, m_HueWheelCenterY As Single
Private m_HueRadiusInner As Single, m_HueRadiusOuter As Single

'Various saturation + value box positioning values.  These are calculated by the CreateSVSquare function and cached here, as a
' convenience for subsequent hit-testing and rendering.
Private m_SVRectF As RectF

'Current control HSV values, on the range [0, 1].  Make sure to update these if a new color is supplied externally.
Private m_Hue As Double, m_Saturation As Double, m_Value As Double

'If the mouse is currently over the hue wheel, but the left mouse button is *not* down, this will be set to a value >= 0.
' We can use this to help orient the user.
Private m_HueHover As Double, m_SaturationHover As Double, m_ValueHover As Double

'User control support class.  Historically, many classes (and associated subclassers) were required by each user control,
' but I've since wrapped these into a single central support class.
Private WithEvents ucSupport As pdUCSupport
Attribute ucSupport.VB_VarHelpID = -1

'Local list of themable colors.  This list includes all potential colors used by this class, regardless of state change
' or internal control settings.  The list is updated by calling the UpdateColorList function.
' (Note also that this list does not include variants, e.g. "BorderColor" vs "BorderColor_Hovered".  Variant values are
'  automatically calculated by the color management class, and they are retrieved by passing boolean modifiers to that
'  class, rather than treating every imaginable variant as a separate constant.)
Private Enum PDCW_COLOR_LIST
    [_First] = 0
    PDCW_Background = 0
    PDCW_WheelBorder = 1
    PDCW_BoxBorder = 2
    [_Last] = 2
    [_Count] = 3
End Enum

'Color retrieval and storage is handled by a dedicated class; this allows us to optimize theme interactions,
' without worrying about the details locally.
Private m_Colors As pdThemeColors

Public Function GetControlType() As PD_ControlType
    GetControlType = pdct_ColorWheel
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
    PropertyChanged "Enabled"
End Property

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

Public Property Get ContainerHwnd() As Long
    ContainerHwnd = UserControl.ContainerHwnd
End Property

Public Property Get Color() As Long
    Color = GetCurrentRGB()
End Property

Public Property Let Color(ByVal newColor As Long)
    
    'Extract matching HSV values, then redraw the control to match
    Dim tmpHue As Double, tmpSaturation As Double, tmpValue As Double
    Colors.RGBtoHSV Colors.ExtractRed(newColor), Colors.ExtractGreen(newColor), Colors.ExtractBlue(newColor), tmpHue, tmpSaturation, tmpValue
    
    If ((tmpHue <> m_Hue) Or (tmpSaturation <> m_Saturation) Or (tmpValue <> m_Value)) Then
    
        If (tmpSaturation <> 0#) Then m_Hue = tmpHue
        If (tmpValue <> 0#) Then m_Saturation = tmpSaturation
        m_Value = tmpValue
        
        CreateSVSquare
        If ucSupport.AmIVisible() Then RedrawBackBuffer True
        
        'Raise a matching event, and note that the source was external
        RaiseEvent ColorChanged(newColor, False)
    
    End If
    
End Property

Public Property Get WheelWidth() As Single
    WheelWidth = m_WheelWidth
End Property

Public Property Let WheelWidth(ByVal newWidth As Single)
    m_WheelWidth = newWidth
    UpdateControlLayout
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

Public Sub SetPosition(ByVal newLeft As Long, ByVal newTop As Long)
    ucSupport.RequestNewPosition newLeft, newTop, True
End Sub

Public Sub SetPositionAndSize(ByVal newLeft As Long, ByVal newTop As Long, ByVal newWidth As Long, ByVal newHeight As Long)
    ucSupport.RequestFullMove newLeft, newTop, newWidth, newHeight, True
End Sub

Private Sub ucSupport_CustomMessage(ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturn As Long)
    If (wMsg = WM_PD_COLOR_MANAGEMENT_CHANGE) Then NotifyColorManagementChange
End Sub

'When the control receives focus, relay the event externally
Private Sub ucSupport_GotFocusAPI()
    RaiseEvent GotFocusAPI
End Sub

Private Sub ucSupport_KeyDownSystem(ByVal Shift As ShiftConstants, ByVal whichSysKey As PD_NavigationKey, markEventHandled As Boolean)
    
    'Enter/Esc get reported directly to the system key handler.  Note that we track the return, because TRUE
    ' means the key was successfully forwarded to the relevant handler.  (If FALSE is returned, no control
    ' accepted the keypress, meaning we should forward the event down the line.)
    markEventHandled = NavKey.NotifyNavKeypress(Me, whichSysKey, Shift)
    
End Sub

'When the control loses focus, relay the event externally
Private Sub ucSupport_LostFocusAPI()
    RaiseEvent LostFocusAPI
End Sub

Private Sub ucSupport_MouseDownCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal timeStamp As Long)
    
    'Right now, only left-clicks are addressed
    If (Button And pdLeftButton) <> 0 Then
    
        'See if the mouse cursor is inside the hue wheel
        Dim tmpHue As Double
        m_MouseDownWheel = IsMouseInsideHueWheel(x, y, True, tmpHue)
        
        'If the mouse is down inside the wheel area, assign a new hue value to the control
        If m_MouseDownWheel Then
            
            'Store the new hue value, and reset a number of other mouse values
            m_Hue = tmpHue
            m_HueHover = -1
            m_MouseDownBox = False
            
            'Set a persistent hand cursor
            ucSupport.RequestCursor IDC_HAND
            
            'Any time the hue changes, the SV square must be redrawn
            CreateSVSquare
            
            'Redraw the control to match
            RedrawBackBuffer True
            
            'Return the newly selected color
            RaiseEvent ColorChanged(Me.Color, True)
        
        Else
            
            'See if the mouse cursor is inside the saturation + value box
            Dim tmpSaturation As Double, tmpValue As Double
            m_MouseDownBox = IsMouseInsideSVBox(x, y, True, tmpSaturation, tmpValue)
            
            If m_MouseDownBox Then
                
                'Store the new saturation and value values, and reset a number of other mouse trackers
                m_Saturation = tmpSaturation
                m_Value = tmpValue
                m_MouseDownWheel = False
                
                'Set a persistent hand cursor
                ucSupport.RequestCursor IDC_HAND
                
                'Redraw the control to match
                RedrawBackBuffer True
                
                'Return the newly selected color
                RaiseEvent ColorChanged(Me.Color, True)
            
            End If
        
        End If
        
    End If
    
End Sub

Private Sub ucSupport_MouseLeave(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    m_MouseInsideWheel = False: m_MouseInsideBox = False
    Me.AssignTooltip vbNullString, , False
    ucSupport.RequestCursor IDC_DEFAULT
    RedrawBackBuffer
End Sub

Private Sub ucSupport_MouseMoveCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal timeStamp As Long)
    
    Dim tmpHue As Double, tmpSaturation As Double, tmpValue As Double
    
    'If the mouse button was originally clicked inside the hue wheel, continue re-calculating hue, regardless of mouse position.
    If m_MouseDownWheel Then
        
        'Calculate a corresponding hue for this mouse position
        IsMouseInsideHueWheel x, y, True, tmpHue
        
        'Store this as the active hue, and reset box parameters
        m_Hue = tmpHue
        m_HueHover = -1
        m_MouseDownBox = False
        m_MouseInsideBox = False
        
        'Any time the hue changes, the SV square must be redrawn
        CreateSVSquare
        
    'The mouse was not originally clicked inside the wheel.  See if the mouse was clicked inside the box
    ElseIf m_MouseDownBox Then
    
        'Calculate corresponding saturation and value values for this mouse position
        IsMouseInsideSVBox x, y, True, tmpSaturation, tmpValue
        
        'Store these as the active saturation+value, and reset wheel parameters
        m_Saturation = tmpSaturation
        m_Value = tmpValue
        m_MouseDownWheel = False
        m_MouseInsideWheel = False
        m_HueHover = -1
        
    'The mouse was not clicked inside the box or wheel.  Ignore other clicks, and update the cursor as necessary.
    Else
    
        'Wheel first
        m_MouseInsideWheel = IsMouseInsideHueWheel(x, y, True, tmpHue)
        
        If m_MouseInsideWheel Then
            ucSupport.RequestCursor IDC_HAND
            m_HueHover = tmpHue
            m_MouseInsideBox = False
        Else
            
            m_HueHover = -1
            
            'Box second
            m_MouseInsideBox = IsMouseInsideSVBox(x, y, True, tmpSaturation, tmpValue)
            
            If m_MouseInsideBox Then
                m_SaturationHover = tmpSaturation
                m_ValueHover = tmpValue
                ucSupport.RequestCursor IDC_HAND
            Else
                m_SaturationHover = -1
                m_ValueHover = -1
                ucSupport.RequestCursor IDC_DEFAULT
            End If
            
        End If
        
    End If
    
    'Redraw the UC to match
    RedrawBackBuffer True
    
    'If the mouse is inside the wheel or box, update our tooltip accordingly
    If (m_MouseInsideWheel Or m_MouseInsideBox) Then
        
        Dim hoverColor As Long
        hoverColor = Colors.ConvertSystemColor(GetProposedColor())
        
        'Construct hex and RGB string representations of the target color
        Dim hexString As String, rgbString As String
        hexString = "#" & UCase$(Colors.GetHexStringFromRGB(hoverColor))
        rgbString = g_Language.TranslateMessage("RGB(%1, %2, %3)", Colors.ExtractRed(hoverColor), Colors.ExtractGreen(hoverColor), Colors.ExtractBlue(hoverColor))
        Me.AssignTooltip hexString & vbCrLf & rgbString, , True
        
    Else
        Me.AssignTooltip vbNullString, , False
    End If
    
    'If the LMB is down, raise an event to match
    If m_MouseDownWheel Or m_MouseDownBox Then RaiseEvent ColorChanged(Me.Color, True)
    
End Sub

'Returns TRUE if the passed (x, y) coordinates lie inside the hue wheel.  An optional output parameter can be provided,
' and this function will automatically fill it with the hue value at that (x, y) position.
Private Function IsMouseInsideHueWheel(ByVal x As Single, ByVal y As Single, Optional ByVal calculateHue As Boolean = False, Optional ByRef dstHue As Double) As Boolean
    
    'Start by re-centering the (x, y) pair around the hue wheel's center point
    x = x - m_HueWheelCenterX
    y = y - m_HueWheelCenterY
    
    'Calculate a radius for the current position
    Dim pxRadius As Double
    pxRadius = Sqr(x * x + y * y)
    
    'If the radius lies between the outer and inner hue wheel radii, return true.
    IsMouseInsideHueWheel = ((pxRadius <= m_HueRadiusOuter) And (pxRadius >= m_HueRadiusInner))
    
    'If the caller wants us to calculate hue for them, do so now.  Note that we can successfully do this, even if the mouse is
    ' outside the hue wheel - this is important for enabling convenient click-drag behavior!
    If calculateHue Then
        
        'Calculate an angle for this pixel
        Dim pxAngle As Double
        pxAngle = PDMath.Atan2(y, x)
        
        'ATan2() returns an angle that is positive for counter-clockwise angles (y > 0), and negative for
        ' clockwise angles (y < 0), on the range [-Pi, +Pi].  Convert this angle to the absolute range [0, 1],
        ' which is the range used by PD's HSV conversion functions.
        dstHue = (pxAngle + PI) / PI_DOUBLE
        
    End If
    
End Function

'Returns TRUE if the passed (x, y) coordinates lie inside the saturation + value box.  Optional output parameters can be
' provided, and this function will automatically fill them with the SV values at that (x, y) position.
Private Function IsMouseInsideSVBox(ByVal x As Single, ByVal y As Single, Optional ByVal calculateSV As Boolean = False, Optional ByRef dstSaturation As Double, Optional ByRef dstValue As Double) As Boolean
    
    'Hit-detection is easy, since we cache the box coordinates when recreating the corresponding DIB
    IsMouseInsideSVBox = PDMath.IsPointInRectF(x, y, m_SVRectF)
    
    'If the caller wants us to calculate saturation and value outputs, do so now
    If calculateSV Then
        
        'In the current design, X controls saturation while Y controls value.  The values are also reversed in the
        ' on-screen display, so that the color itself sits closest to the canvas.
        dstSaturation = 1# - ((x - m_SVRectF.Left) / m_SVRectF.Width)
        dstValue = 1# - ((y - m_SVRectF.Top) / m_SVRectF.Height)
        
        'To prevent errors, clamp saturation and value now
        If (dstSaturation < 0#) Then dstSaturation = 0#
        If (dstSaturation > 1#) Then dstSaturation = 1#
        If (dstValue < 0#) Then dstValue = 0#
        If (dstValue > 1#) Then dstValue = 1#
        
        'The y-value is squared during rendering, to decrease the amount of space taken up by extremely dark color variants
        If (dstValue > 0#) Then dstValue = Sqr(dstValue)
        
    End If
    
End Function

Private Sub ucSupport_MouseUpCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal clickEventAlsoFiring As Boolean, ByVal timeStamp As Long)
    
    m_MouseDownWheel = False
    m_MouseDownBox = False
    
    'Reset the cursor and hover behavior accordingly
    Dim tmpHue As Double, tmpSaturation As Double, tmpValue As Double
    m_MouseInsideWheel = IsMouseInsideHueWheel(x, y, True, tmpHue)
    
    If m_MouseInsideWheel Then
        ucSupport.RequestCursor IDC_HAND
        m_HueHover = tmpHue
    Else
        
        m_HueHover = -1
        
        m_MouseInsideBox = IsMouseInsideSVBox(x, y, True, tmpSaturation, tmpValue)
        If m_MouseInsideBox Then
            ucSupport.RequestCursor IDC_HAND
        Else
            ucSupport.RequestCursor IDC_DEFAULT
        End If
        
    End If
    
    'Redraw the control to match
    RedrawBackBuffer
    
End Sub

Private Sub ucSupport_RepaintRequired(ByVal updateLayoutToo As Boolean)
    If updateLayoutToo Then UpdateControlLayout Else RedrawBackBuffer
End Sub

Private Sub ucSupport_VisibilityChange(ByVal newVisibility As Boolean)
    If newVisibility Then
        CreateSVSquare
        RedrawBackBuffer True
    End If
End Sub

Private Sub UserControl_Initialize()
    
    'Initialize a user control support class
    Set ucSupport = New pdUCSupport
    ucSupport.RegisterControl UserControl.hWnd, True
    ucSupport.RequestExtraFunctionality True
    ucSupport.SubclassCustomMessage WM_PD_COLOR_MANAGEMENT_CHANGE, True
    
    'Prep the color manager and load default colors
    Set m_Colors = New pdThemeColors
    Dim colorCount As PDCW_COLOR_LIST: colorCount = [_Count]
    m_Colors.InitializeColorList "PDColorWheel", colorCount
    If Not PDMain.IsProgramRunning() Then UpdateColorList
    
    'Draw the control at least once
    UpdateControlLayout
    
End Sub

Private Sub UserControl_InitProperties()
    Color = RGB(50, 200, 255)
    WheelWidth = WHEEL_WIDTH_DEFAULT
End Sub

'At run-time, painting is handled by the support class.  In the IDE, however, we must rely on VB's internal paint event.
Private Sub UserControl_Paint()
    ucSupport.RequestIDERepaint UserControl.hDC
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        Me.Color = .ReadProperty("Color", RGB(50, 200, 255))
        Me.WheelWidth = .ReadProperty("WheelWidth", WHEEL_WIDTH_DEFAULT)
    End With
End Sub

Private Sub UserControl_Resize()
    If Not PDMain.IsProgramRunning() Then ucSupport.RequestRepaint True
End Sub
    
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "Color", Me.Color, RGB(50, 200, 255)
        .WriteProperty "WheelWidth", Me.WheelWidth, WHEEL_WIDTH_DEFAULT
    End With
End Sub

Private Sub NotifyColorManagementChange()
    CreateColorWheel
    CreateSVSquare
    RedrawBackBuffer True
End Sub

'Call this to recreate all buffers against a changed control size.
Private Sub UpdateControlLayout()
    
    'Recreate all individual components, as their size is dependent on the container size
    If PDMain.IsProgramRunning() And ucSupport.AmIVisible Then
        CreateColorWheel
        CreateSVSquare
        RedrawBackBuffer
    End If
    
End Sub

'Create the color wheel portion of the selector.  Note that this function cannot fire until the backbuffer has been initialized,
' because it relies on that buffer for sizing.
Private Sub CreateColorWheel()
    
    'For now, the color wheel DIB is always square, sized to fit the smallest dimension of the back buffer
    Dim wheelDiameter As Long
    If (ucSupport.GetBackBufferWidth < ucSupport.GetBackBufferHeight) Then wheelDiameter = ucSupport.GetBackBufferWidth Else wheelDiameter = ucSupport.GetBackBufferHeight
    
    If (m_WheelBuffer Is Nothing) Then Set m_WheelBuffer = New pdDIB
    If (m_WheelBuffer.GetDIBWidth <> wheelDiameter) Or (m_WheelBuffer.GetDIBHeight <> wheelDiameter) Then
        m_WheelBuffer.CreateBlank wheelDiameter, wheelDiameter, 32, 0&, 255
    Else
        If PDMain.IsProgramRunning() Then GDI_Plus.GDIPlusFillDIBRect m_WheelBuffer, 0, 0, wheelDiameter, wheelDiameter, 0&, 255
    End If
    
    'We're now going to calculate the inner and outer radius of the wheel.  These are based off hard-coded padding constants,
    ' the max available diameter, and the current screen DPI.
    m_HueRadiusOuter = (CSng(wheelDiameter) / 2) - Interface.FixDPIFloat(WHEEL_PADDING)
    m_HueRadiusInner = m_HueRadiusOuter - Interface.FixDPIFloat(m_WheelWidth)
    If (m_HueRadiusInner < 5) Then m_HueRadiusInner = 5
    
    'At high DPIs, the outer wheel (on the primary toolbar) can dominate overall wheel size.
    ' To prevent this, limit outer diameter to a ratio of total object size.
    If ((m_HueRadiusOuter - m_HueRadiusInner) > wheelDiameter / 6) Then m_HueRadiusInner = wheelDiameter * (1 / 3)
    If (m_HueRadiusInner >= m_HueRadiusOuter - 5) Then m_HueRadiusInner = m_HueRadiusOuter - 5
    
    'We're now going to cheat a bit and use a 2D drawing hack to solve for the alpha bytes of our wheel.  The wheel image is
    ' already a black square, and atop that we're going to draw a white circle at the outer radius size, and a black circle
    ' at the inner radius size.  Both will be antialiased.  Black pixels will then be made transparent, while white pixels
    ' are fully opaque.  Gray pixels will be shaded on-the-fly.
    m_HueWheelCenterX = wheelDiameter / 2: m_HueWheelCenterY = m_HueWheelCenterX
    
    If PDMain.IsProgramRunning() Then
        Dim cSurface As pd2DSurface, cBrush As pd2DBrush
        Drawing2D.QuickCreateSurfaceFromDC cSurface, m_WheelBuffer.GetDIBDC, True
        Drawing2D.QuickCreateSolidBrush cBrush, RGB(255, 255, 255)
        PD2D.FillCircleF cSurface, cBrush, m_HueWheelCenterX, m_HueWheelCenterY, m_HueRadiusOuter
        cBrush.SetBrushColor RGB(0, 0, 0)
        PD2D.FillCircleF cSurface, cBrush, m_HueWheelCenterX, m_HueWheelCenterY, m_HueRadiusInner
        Set cBrush = Nothing: Set cSurface = Nothing
    End If
    
    'With our "alpha guidance" pixels drawn, we can now loop through the image, rendering actual hue colors as we go.
    ' For convenience, we will place hue 0 at angle 0.
    Dim hPixels() As Byte, hueSA As SafeArray1D
    Dim x As Long, y As Long
    Dim r As Double, g As Double, b As Double, a As Long, aFloat As Single
    
    Dim nX As Double, nY As Double, pxAngle As Double
    
    Dim loopWidth As Long, loopHeight As Long
    loopWidth = (m_WheelBuffer.GetDIBWidth - 1) * 4
    loopHeight = (m_WheelBuffer.GetDIBHeight - 1)
    
    Dim fastDivisor As Single
    fastDivisor = 1! / 255!
    
    Dim fLookup() As Single
    ReDim fLookup(0 To 255) As Single
    For x = 0 To 255
        fLookup(x) = CSng(x) * fastDivisor
    Next x
    
    fastDivisor = 1# / PI_DOUBLE
    
    Dim hPtr As Long, hStride As Long
    m_WheelBuffer.WrapArrayAroundScanline hPixels, hueSA, 0
    hPtr = hueSA.pvData
    hStride = hueSA.cElements
    
    For y = 0 To loopHeight
        hueSA.pvData = hPtr + hStride * y
    For x = 0 To loopWidth Step 4
        
        'Before calculating anything, check the color at this position.  (Because the image is grayscale, we only need to
        ' pull a single color value.)
        b = hPixels(x)
        
        'If this pixel is black, it will be forced to full transparency.  Apply that now.
        If (b = 0) Then
            hPixels(x) = 0
            hPixels(x + 1) = 0
            hPixels(x + 2) = 0
            hPixels(x + 3) = 0
        
        'If this pixel is non-black, it must be colored.  Proceed with hue calculation.
        Else
        
            'Remap the coordinates so that (0, 0) represents the center of the image
            nX = (x * 0.25) - m_HueWheelCenterX
            nY = y - m_HueWheelCenterY
            
            'Calculate an angle for this pixel
            pxAngle = PDMath.Atan2_Faster(nY, nX)
            
            'ATan2() returns an angle that is positive for counter-clockwise angles (y > 0), and negative for
            ' clockwise angles (y < 0), on the range [-Pi, +Pi].  Convert this angle to the absolute range [0, 1],
            ' which is the range used by our HSV conversion function.
            pxAngle = (pxAngle + PI) * fastDivisor
            
            'Calculate an RGB triplet that corresponds to this hue (with max value and saturation)
            Colors.fHSVtoRGB pxAngle, 1#, 1#, r, g, b
            
            'Retrieve the "alpha" clue for this pixel
            a = hPixels(x)
            aFloat = fLookup(a)
            
            'Premultiply alpha
            r = r * aFloat
            g = g * aFloat
            b = b * aFloat
            
            'Store the new color values
            hPixels(x) = b * 255#
            hPixels(x + 1) = g * 255#
            hPixels(x + 2) = r * 255#
            hPixels(x + 3) = a
            
        End If
    
    Next x
    Next y
    
    'With our work complete, point the array away from the DIB before VB attempts to deallocate it
    m_WheelBuffer.UnwrapArrayFromDIB hPixels
    
    'If color management is active, apply it now
    ColorManagement.ApplyDisplayColorManagement m_WheelBuffer
    
    'Mark the wheel DIB's premultiplied alpha state
    m_WheelBuffer.SetInitialAlphaPremultiplicationState True
    
End Sub

'Create a new Saturation + Value square (the square in the middle of the UC).  The square must be redrawn whenever
' hue changes, because the hue value determines the square's appearance.
Private Sub CreateSVSquare()
    
    'The SV square is a square that fits (inclusively) within the color wheel.
    ' Basic geometry tells us that one side of the square is equal to hypotenuse * sin(45),
    ' and we know the hypotenuse already because it's the inner radius of the hue wheel.
    m_SVRectF.Width = (m_HueRadiusInner * 2) * Sin(PI * 0.25): m_SVRectF.Height = m_SVRectF.Width
    
    If (m_SquareBuffer Is Nothing) Then Set m_SquareBuffer = New pdDIB
    If (m_SquareBuffer.GetDIBWidth <> CLng(m_SVRectF.Width)) Or (m_SquareBuffer.GetDIBHeight <> CLng(m_SVRectF.Height)) Or (m_SVRectF.Width = 0) Or (m_SVRectF.Height = 0) Then
        m_SquareBuffer.CreateBlank CLng(m_SVRectF.Width), CLng(m_SVRectF.Height), 32, , 255
    Else
        m_SquareBuffer.ResetDIB 255
    End If
    
    m_SquareBuffer.SetInitialAlphaPremultiplicationState True
    
    'To prevent IDE crashes, bail now during compilation
    If (Not PDMain.IsProgramRunning()) Or (Not ucSupport.AmIVisible) Then Exit Sub
    
    'We now need to fill the square with all possible saturation and value variants, in a pattern where...
    ' - The y-axis position determines value (1 -> 0)
    ' - The x-axis position determines saturation (1 -> 0)
    Dim svPixels() As Byte, svSA As SafeArray1D
    
    Dim xPxWidth As Long
    xPxWidth = m_SquareBuffer.GetDIBColorDepth \ 8
    
    Dim x As Long, y As Long
    Dim r As Long, g As Long, b As Long
    
    Dim loopWidth As Long, loopHeight As Long
    loopWidth = (m_SquareBuffer.GetDIBWidth - 1) * xPxWidth
    loopHeight = (m_SquareBuffer.GetDIBHeight - 1)
    
    Dim lineValue As Double
    
    'To improve performance, pre-calculate all value variants, so we don't need to re-calculate them in the inner loop.
    ' (They are constant for each line.)
    Dim xPresets() As Double
    ReDim xPresets(0 To loopWidth) As Double
    
    Dim xMultiplier As Double, yMultiplier As Double
    If (loopWidth <> 0#) Then xMultiplier = 1# / loopWidth Else xMultiplier = 1#
    If (loopHeight <> 0#) Then yMultiplier = 1# / loopHeight Else yMultiplier = 1#
    
    For x = 0 To loopWidth Step xPxWidth
        xPresets(x) = (loopWidth - x) * xMultiplier
    Next x
    
    For y = 0 To loopHeight
        
        'Y-values are (obviously) consistent for each y-position
        lineValue = (loopHeight - y) * yMultiplier
        lineValue = Sqr(lineValue)
        m_SquareBuffer.WrapArrayAroundScanline svPixels, svSA, y
        
    For x = 0 To loopWidth Step xPxWidth
        
        'The x-axis position determines saturation (1 -> 0)
        'The y-axis position determines value (1 -> 0)
        Colors.HSVtoRGB m_Hue, xPresets(x), lineValue, r, g, b
        
        svPixels(x) = b
        svPixels(x + 1) = g
        svPixels(x + 2) = r
        
    Next x
    Next y
    
    'With our work complete, point the ImageData() array away from the DIBs and deallocate it
    m_SquareBuffer.UnwrapArrayFromDIB svPixels
    
    'While we're here, let's also calculate the top-left rendering origin for the square, so we don't have to do it in the core
    ' rendering function.
    Dim tmpX As Double, tmpY As Double
    PDMath.ConvertPolarToCartesian -0.75 * PI, m_HueRadiusInner, tmpX, tmpY, m_HueWheelCenterX, m_HueWheelCenterY
    m_SVRectF.Left = tmpX
    m_SVRectF.Top = tmpY
    
    'If color management is active, apply it now
    ColorManagement.ApplyDisplayColorManagement m_SquareBuffer, , False
    
End Sub

'Get the hypothetical color if the use clicks the mouse in its current position.  If the mouse is not in a valid location,
' the control's current color will be returned.
Private Function GetProposedColor() As Long
    
    If m_MouseDownBox Or m_MouseDownWheel Then
        GetProposedColor = GetCurrentRGB
    Else
        If m_MouseInsideBox Then
            If (m_SaturationHover <> -1) And (m_ValueHover <> -1) Then
                GetProposedColor = GetHypotheticalRGB(m_Hue, m_SaturationHover, m_ValueHover)
            Else
                GetProposedColor = GetCurrentRGB
            End If
        Else
            If (m_HueHover <> -1) Then
                GetProposedColor = GetHypotheticalRGB(m_HueHover, m_Saturation, m_Value)
            Else
                GetProposedColor = GetCurrentRGB
            End If
        End If
    End If
            
End Function

'Redraw the UC.  Note that some UI elements must be created prior to calling this function (e.g. the color wheel).
Private Sub RedrawBackBuffer(Optional ByVal paintImmediately As Boolean = False)
    
    'Request the back buffer DC, and ask the support module to erase any existing rendering for us.
    Dim bufferDC As Long
    bufferDC = ucSupport.GetBackBufferDC(True, m_Colors.RetrieveColor(PDCW_Background, Me.Enabled))
    If (bufferDC = 0) Then Exit Sub
    
    Dim bWidth As Long, bHeight As Long
    bWidth = ucSupport.GetBackBufferWidth
    bHeight = ucSupport.GetBackBufferHeight
    
    Dim wheelBorderColor As Long, boxBorderColor As Long, colorPreviewBorder As Long
    wheelBorderColor = m_Colors.RetrieveColor(PDCW_WheelBorder, Me.Enabled, False, m_MouseInsideWheel)
    boxBorderColor = m_Colors.RetrieveColor(PDCW_BoxBorder, Me.Enabled, False, m_MouseInsideBox)
    colorPreviewBorder = m_Colors.RetrieveColor(PDCW_BoxBorder, Me.Enabled, False, False)
    
    If PDMain.IsProgramRunning() And (m_HueRadiusInner > 0) Then
        
        'Paint the hue wheel (currently left-aligned)
        If (Not m_WheelBuffer Is Nothing) Then
            m_WheelBuffer.AlphaBlendToDC bufferDC
            m_WheelBuffer.FreeFromDC
        End If
        
        'Prep various painting objects
        Dim cSurface As pd2DSurface, cBrush As pd2DBrush, cPen As pd2DPen
        Dim cPenUIBase As pd2DPen, cPenUITop As pd2DPen
        Drawing2D.QuickCreateSurfaceFromDC cSurface, bufferDC, True
        
        'Trace the edges of the hue wheel, to help separate the bright portions from the background.
        Dim borderWidth As Single, borderTransparency As Single
        If m_MouseInsideWheel Then
            borderWidth = 2!
            borderTransparency = 100!
        Else
            borderWidth = 1!
            borderTransparency = 75!
        End If
        Drawing2D.QuickCreateSolidPen cPen, borderWidth, wheelBorderColor, borderTransparency
        PD2D.DrawCircleF cSurface, cPen, m_HueWheelCenterX, m_HueWheelCenterY, m_HueRadiusInner
        PD2D.DrawCircleF cSurface, cPen, m_HueWheelCenterX, m_HueWheelCenterY, m_HueRadiusOuter
        
        'Paint the saturation+value square
        If (Not m_SquareBuffer Is Nothing) Then
            
            'Copy the square into place.  Note that we must use GDI+ to support subpixel positioning.
            With m_SVRectF
                GDI_Plus.GDIPlus_StretchBlt Nothing, .Left, .Top, .Width, .Height, m_SquareBuffer, 0, 0, m_SquareBuffer.GetDIBWidth, m_SquareBuffer.GetDIBHeight, , GP_IM_NearestNeighbor, bufferDC
                m_SquareBuffer.FreeFromDC
            End With
            
            'Trace the edges of the square, to help separate the bright portions from the background
            If m_MouseInsideBox Then
                borderWidth = 2!
                borderTransparency = 100!
            Else
                borderWidth = 1!
                borderTransparency = 50!
            End If
            Drawing2D.QuickCreateSolidPen cPen, borderWidth, boxBorderColor, borderTransparency, P2_LJ_Miter, P2_LC_Flat
            If (Not m_MouseInsideBox) Then cSurface.SetSurfacePixelOffset P2_PO_Half
            PD2D.DrawRectangleF_FromRectF cSurface, cPen, m_SVRectF
            cSurface.SetSurfacePixelOffset P2_PO_Normal
            
        End If
        
        'Draw a "pie-slice" outline around the current hue value.  Start by retrieving the UI angle of the current hue value
        Dim hueAngle As Single
        hueAngle = GetUIAngleOfHue(m_Hue)
        
        'We are now going to construct a "slice-like" overlay for the current hue position.
        Dim slicePath As pd2DPath
        Set slicePath = New pd2DPath
        
        'The sweep of the slice should really be contingent on the radius, but for this first draft, we'll simply hard-code it.
        Dim sliceSweep As Single
        sliceSweep = 0.18
        
        'Also, the slice will extend beyond the interior and exterior edges of the hue wheel by some fixed amount (currently 0.5 pixels)
        Dim sliceExtend As Single
        sliceExtend = 0.5
        
        'Next, calculate (x, y) coordinates for the four corners of the slice.  We use these as the endpoints for the radial lines
        ' marking either side of the "slice".
        Dim x1 As Double, x2 As Double, x3 As Double, x4 As Double, y1 As Double, y2 As Double, y3 As Double, y4 As Double
        PDMath.ConvertPolarToCartesian hueAngle - (sliceSweep / 2), m_HueRadiusInner - sliceExtend, x1, y1, m_HueWheelCenterX, m_HueWheelCenterY
        PDMath.ConvertPolarToCartesian hueAngle - (sliceSweep / 2), m_HueRadiusOuter + sliceExtend, x2, y2, m_HueWheelCenterX, m_HueWheelCenterY
        PDMath.ConvertPolarToCartesian hueAngle + (sliceSweep / 2), m_HueRadiusInner - sliceExtend, x3, y3, m_HueWheelCenterX, m_HueWheelCenterY
        PDMath.ConvertPolarToCartesian hueAngle + (sliceSweep / 2), m_HueRadiusOuter + sliceExtend, x4, y4, m_HueWheelCenterX, m_HueWheelCenterY
        
        'Add those two lines to the path object, and place connecting arcs between them
        slicePath.AddLine x1, y1, x2, y2
        slicePath.AddArcCircular m_HueWheelCenterX, m_HueWheelCenterY, m_HueRadiusOuter + sliceExtend, RadiansToDegrees(hueAngle - (sliceSweep / 2)), RadiansToDegrees(sliceSweep)
        slicePath.AddLine x4, y4, x3, y3
        slicePath.AddArcCircular m_HueWheelCenterX, m_HueWheelCenterY, m_HueRadiusInner - sliceExtend, RadiansToDegrees(hueAngle + (sliceSweep / 2)), RadiansToDegrees(-sliceSweep)
        slicePath.CloseCurrentFigure
        
        'Render the completed slice onto the overlay
        Drawing2D.QuickCreatePairOfUIPens cPenUIBase, cPenUITop, m_MouseDownWheel
        PD2D.DrawPath cSurface, cPenUIBase, slicePath
        PD2D.DrawPath cSurface, cPenUITop, slicePath
        cSurface.SetSurfacePixelOffset P2_PO_Normal
        
        'Lastly, let's draw a circle around the current saturation + value point.
        
        'Convert the saturation + value point to an (x, y) pair
        Dim svX As Double, svY As Double
        svX = m_SVRectF.Width * (1 - m_Saturation)
        svY = m_SVRectF.Height * (1 - m_Value * m_Value)
        
        'The (x, y) pair we've calculated will lie well outside the SV square's borders as-is.  Trim the edges a bit,
        ' to make it look better.
        Dim COLOR_CIRCLE_RADIUS As Single, COLOR_CIRCLE_CHECK As Single
        COLOR_CIRCLE_RADIUS = FixDPIFloat(5#)
        COLOR_CIRCLE_CHECK = COLOR_CIRCLE_RADIUS - FixDPIFloat(3#)
        If (svX < COLOR_CIRCLE_CHECK) Then svX = COLOR_CIRCLE_CHECK
        If (svY < COLOR_CIRCLE_CHECK) Then svY = COLOR_CIRCLE_CHECK
        If (svX > (m_SVRectF.Width - (COLOR_CIRCLE_CHECK + 1))) Then svX = (m_SVRectF.Width - (COLOR_CIRCLE_CHECK + 1))
        If (svY > (m_SVRectF.Height - (COLOR_CIRCLE_CHECK + 1))) Then svY = (m_SVRectF.Height - (COLOR_CIRCLE_CHECK + 1))
        
        'Pad the circle by the current SV square's offset
        svX = svX + m_SVRectF.Left
        svY = svY + m_SVRectF.Top
        
        'Draw a canvas-style circle around that point
        Drawing.BorrowCachedUIPens cPenUIBase, cPenUITop, m_MouseDownBox
        PD2D.DrawCircleF cSurface, cPenUIBase, svX, svY, COLOR_CIRCLE_RADIUS
        PD2D.DrawCircleF cSurface, cPenUITop, svX, svY, COLOR_CIRCLE_RADIUS
        
        'Finally, if the mouse is over the hue wheel or SV box, but the mouse is *NOT* down, we want to paint a little
        ' color triangle to let the user know what color they'd get if they DID click the mouse button in this location.
        If (m_MouseInsideBox Or m_MouseInsideWheel) Then
            
            'Paint the color in a small triangle in the corner
            Dim pcPath As pd2DPath
            Set pcPath = New pd2DPath
            
            Dim pcLength As Double, wWidth As Double, wHeight As Double
            wWidth = m_WheelBuffer.GetDIBWidth - 1: wHeight = m_WheelBuffer.GetDIBHeight - 1
            pcLength = Sqr(wWidth * wWidth + wHeight * wHeight) * 0.5
            pcLength = (pcLength - (wWidth * 0.5)) * (PI_HALF * 0.8)
            pcPath.AddTriangle wWidth, wHeight, wWidth - pcLength, wHeight, wWidth, wHeight - pcLength
            
            Drawing2D.QuickCreateSolidBrush cBrush, GetProposedColor()
            PD2D.FillPath cSurface, cBrush, pcPath
            
            Drawing2D.QuickCreateSolidPen cPen, 1, colorPreviewBorder, 75, P2_LJ_Round, P2_LC_Round
            PD2D.DrawPath cSurface, cPen, pcPath
            
            Set pcPath = Nothing
            
        End If
        
        Set cSurface = Nothing: Set cBrush = Nothing: Set cPen = Nothing
        
    End If
    
    'Paint the final result to the screen, as relevant
    ucSupport.RequestRepaint paintImmediately

End Sub

'Given a hue on the range [0, 1], return a GDIPlus-friendly UI angle for the hue wheel.
' (A hue of "0" corresponds to an angle of Pi.  A hue of "1" corresponds to an angle of -Pi (hue is circular).)
Private Function GetUIAngleOfHue(ByVal srcHue As Single) As Single
    GetUIAngleOfHue = (srcHue * PI_DOUBLE) - PI
End Function

'Convert the control's current HSV triplet into a corresponding RGB long
Private Function GetCurrentRGB() As Long
    Dim r As Long, g As Long, b As Long
    Colors.HSVtoRGB m_Hue, m_Saturation, m_Value, r, g, b
    GetCurrentRGB = RGB(r, g, b)
End Function

'Given an arbitrary HSV triplet, return the corresponding RGB long
Private Function GetHypotheticalRGB(ByVal h As Double, ByVal s As Double, ByVal v As Double) As Long
    If (s < 0#) Then s = 0#:      If (s > 1#) Then s = 1#
    If (v < 0#) Then v = 0#:      If (v > 1#) Then v = 1#
    Dim r As Long, g As Long, b As Long
    Colors.HSVtoRGB h, s, v, r, g, b
    GetHypotheticalRGB = RGB(r, g, b)
End Function

'Before this control does any painting, we need to retrieve relevant colors from PD's primary theming class.  Note that this
' step must also be called if/when PD's visual theme settings change.
Private Sub UpdateColorList()
    With m_Colors
        .LoadThemeColor PDCW_Background, "Background", IDE_WHITE
        .LoadThemeColor PDCW_WheelBorder, "WheelBorder", IDE_GRAY
        .LoadThemeColor PDCW_BoxBorder, "BoxBorder", IDE_GRAY
    End With
End Sub

'External functions can call this to request a redraw.  This is helpful for live-updating theme settings, as in the Preferences dialog,
' and/or retranslating any text against the current language.
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
