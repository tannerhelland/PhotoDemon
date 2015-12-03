VERSION 5.00
Begin VB.UserControl pdColorWheel 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   1950
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2070
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
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
'Copyright 2015-2015 by Tanner Helland
'Created: 19/October/15
'Last updated: 22/October/15
'Last update: wrap up initial build
'
'In 7.0, a "color selector" panel was added to the right-side toolbar.  Unlike PD's single-color color selector,
' this control is designed to provide a quick, on-canvas-friendly mechanism for rapidly switching colors.  The basic
' design owes much to other photo editors like MyPaint, who pioneered the "wheel" UI for hue selection.
'
'I've designed the control as a UC in case I decide to reuse it elsewhere in PD, but for now, it only makes an
' appearance on the main canvas.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Just like PD's old color selector, this control will raise a ColorChanged event after user interactions.
Public Event ColorChanged(ByVal newColor As Long, ByVal srcIsInternal As Boolean)

'A specialized class handles mouse input for this control
Private WithEvents cMouseEvents As pdInputMouse
Attribute cMouseEvents.VB_VarHelpID = -1

'Reliable focus detection requires a specialized subclasser
Private WithEvents cFocusDetector As pdFocusDetector
Attribute cFocusDetector.VB_VarHelpID = -1
Public Event GotFocusAPI()
Public Event LostFocusAPI()

'Flicker-free window painter
Private WithEvents cPainter As pdWindowPainter
Attribute cPainter.VB_VarHelpID = -1

'Additional helper for rendering themed and multiline tooltips
Private toolTipManager As pdToolTip

'This back buffer is for the composited wheel and center HSV box; it is what gets copied to the screen on Paint events.
Private m_BackBuffer As pdDIB

'Individual UI components are rendered to their own DIBs, and composited only when necessary.  For some elements
' (particularly the hue wheel), creating them from scratch is costly, so reuse is advisable.
Private m_WheelBuffer As pdDIB, m_SquareBuffer As pdDIB

'These values help the central renderer know where the mouse is, so we can draw various indicators.
Private m_MouseInsideWheel As Boolean, m_MouseInsideBox As Boolean
Private m_MouseDownWheel As Boolean, m_MouseDownBox As Boolean

'API technique for drawing a focus rectangle; used only for designer mode (see the Paint method for details)
Private Declare Function DrawFocusRect Lib "user32" (ByVal hDC As Long, lpRect As RECT) As Long

'Padding (in pixels) between the edges of the user control and the color wheel.  Automatically adjusted for DPI
' at run-time.  Note that this needs to be non-zero, because the padding area is used to render the "slice" overlay
' showing the user's current hue selection.
Private Const WHEEL_PADDING As Long = 3

'Width (in pixels) of the hue wheel.  This width is applied along the radial axis.
Private Const WHEEL_WIDTH As Single = 15#

'Various hue wheel positioning values.  These are calculated by the CreateColorWheel function and cached here, as a convenience
' for subsequent hit-testing and rendering.
Private m_HueWheelCenterX As Single, m_HueWheelCenterY As Single
Private m_HueRadiusInner As Single, m_HueRadiusOuter As Single

'Various saturation + value box positioning values.  These are calculated by the CreateSVSquare function and cached here, as a
' convenience for subsequent hit-testing and rendering.
Private m_SVRectF As RECTF

'Current control HSV values, on the range [0, 1].  Make sure to update these if a new color is supplied externally.
Private m_Hue As Double, m_Saturation As Double, m_Value As Double

'If the mouse is currently over the hue wheel, but the left mouse button is *not* down, this will be set to a value >= 0.
' We can use this to help orient the user.
Private m_HueHover As Double

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
    Color_Functions.RGBtoHSV Color_Functions.ExtractR(newColor), Color_Functions.ExtractG(newColor), Color_Functions.ExtractB(newColor), m_Hue, m_Saturation, m_Value
    CreateSVSquare
    DrawUC
    
    'Raise a matching event, and note that the source was external
    RaiseEvent ColorChanged(newColor, False)
    
End Property

'When the control receives focus, relay the event externally
Private Sub cFocusDetector_GotFocusReliable()
    RaiseEvent GotFocusAPI
End Sub

'When the control loses focus, relay the event externally
Private Sub cFocusDetector_LostFocusReliable()
    RaiseEvent LostFocusAPI
End Sub

Private Sub cMouseEvents_MouseDownCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    
    'Right now, only left-clicks are addressed
    If (Button And pdLeftButton) <> 0 Then
    
        'See if the mouse cursor is inside the hue wheel
        Dim tmpHue As Double
        m_MouseDownWheel = isMouseInsideHueWheel(x, y, True, tmpHue)
        
        'If the mouse is down inside the wheel area, assign a new hue value to the control
        If m_MouseDownWheel Then
            
            'Store the new hue value, and reset a number of other mouse values
            m_Hue = tmpHue
            m_HueHover = -1
            m_MouseDownBox = False
            
            'Set a persistent hand cursor
            cMouseEvents.setSystemCursor IDC_HAND
            
            'Any time the hue changes, the SV square must be redrawn
            CreateSVSquare
            
            'Redraw the control to match
            DrawUC
            
            'Return the newly selected color
            RaiseEvent ColorChanged(Me.Color, True)
        
        Else
            
            'See if the mouse cursor is inside the saturation + value box
            Dim tmpSaturation As Double, tmpValue As Double
            m_MouseDownBox = isMouseInsideSVBox(x, y, True, tmpSaturation, tmpValue)
            
            If m_MouseDownBox Then
                
                'Store the new saturation and value values, and reset a number of other mouse trackers
                m_Saturation = tmpSaturation
                m_Value = tmpValue
                m_MouseDownWheel = False
                
                'Set a persistent hand cursor
                cMouseEvents.setSystemCursor IDC_HAND
                
                'Redraw the control to match
                DrawUC
                
                'Return the newly selected color
                RaiseEvent ColorChanged(Me.Color, True)
            
            End If
        
        End If
        
    End If
    
End Sub

Private Sub cMouseEvents_MouseLeave(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    cMouseEvents.setSystemCursor IDC_DEFAULT
End Sub

Private Sub cMouseEvents_MouseMoveCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    
    Dim tmpHue As Double, tmpSaturation As Double, tmpValue As Double
    
    'If the mouse button was originally clicked inside the hue wheel, continue re-calculating hue, regardless of mouse position.
    If m_MouseDownWheel Then
        
        'Calculate a corresponding hue for this mouse position
        isMouseInsideHueWheel x, y, True, tmpHue
        
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
        isMouseInsideSVBox x, y, True, tmpSaturation, tmpValue
        
        'Store these as the active saturation+value, and reset wheel parameters
        m_Saturation = tmpSaturation
        m_Value = tmpValue
        m_MouseDownWheel = False
        m_MouseInsideWheel = False
        m_HueHover = -1
        
    'The mouse was not clicked inside the box or wheel.  Ignore other clicks, and update the cursor as necessary.
    Else
    
        'Wheel first
        m_MouseInsideWheel = isMouseInsideHueWheel(x, y, True, tmpHue)
        
        If m_MouseInsideWheel Then
            cMouseEvents.setSystemCursor IDC_HAND
            m_HueHover = tmpHue
            m_MouseInsideBox = False
        Else
            
            m_HueHover = -1
            
            'Box second
            m_MouseInsideBox = isMouseInsideSVBox(x, y, True, tmpSaturation, tmpValue)
            
            If m_MouseInsideBox Then
                cMouseEvents.setSystemCursor IDC_HAND
            Else
                cMouseEvents.setSystemCursor IDC_DEFAULT
            End If
            
        End If
        
    End If
    
    'Redraw the UC to match
    DrawUC
    
    'If the LMB is down, raise an event to match
    If m_MouseDownWheel Or m_MouseDownBox Then RaiseEvent ColorChanged(Me.Color, True)
    
End Sub

'Returns TRUE if the passed (x, y) coordinates lie inside the hue wheel.  An optional output parameter can be provided,
' and this function will automatically fill it with the hue value at that (x, y) position.
Private Function isMouseInsideHueWheel(ByVal x As Single, ByVal y As Single, Optional ByVal calculateHue As Boolean = False, Optional ByRef dstHue As Double) As Boolean
    
    'Start by re-centering the (x, y) pair around the hue wheel's center point
    x = x - m_HueWheelCenterX
    y = y - m_HueWheelCenterY
    
    'Calculate a radius for the current position
    Dim pxRadius As Double
    pxRadius = Sqr(x * x + y * y)
    
    'If the radius lies between the outer and inner hue wheel radii, return true.
    isMouseInsideHueWheel = CBool((pxRadius <= m_HueRadiusOuter) And (pxRadius >= m_HueRadiusInner))
    
    'If the caller wants us to calculate hue for them, do so now.  Note that we can successfully do this, even if the mouse is
    ' outside the hue wheel - this is important for enabling convenient click-drag behavior!
    If calculateHue Then
        
        'Calculate an angle for this pixel
        Dim pxAngle As Double
        pxAngle = Math_Functions.Atan2(y, x)
        
        'ATan2() returns an angle that is positive for counter-clockwise angles (y > 0), and negative for
        ' clockwise angles (y < 0), on the range [-Pi, +Pi].  Convert this angle to the absolute range [0, 1],
        ' which is the range used by PD's HSV conversion functions.
        dstHue = (pxAngle + PI) / PI_DOUBLE
        
    End If
    
End Function

'Returns TRUE if the passed (x, y) coordinates lie inside the saturation + value box.  Optional output parameters can be
' provided, and this function will automatically fill them with the SV values at that (x, y) position.
Private Function isMouseInsideSVBox(ByVal x As Single, ByVal y As Single, Optional ByVal calculateSV As Boolean = False, Optional ByRef dstSaturation As Double, Optional ByRef dstValue As Double) As Boolean
    
    'Hit-detection is easy, since we cache the box coordinates when recreating the corresponding DIB
    isMouseInsideSVBox = Math_Functions.isPointInRectF(x, y, m_SVRectF)
    
    'If the caller wants us to calculate saturation and value outputs, do so now
    If calculateSV Then
        
        'In the current design, X controls saturation while Y controls value.  The values are also reversed in the
        ' on-screen display, so that the color itself sits closest to the canvas.
        dstSaturation = 1 - ((x - m_SVRectF.Left) / m_SVRectF.Width)
        dstValue = 1 - ((y - m_SVRectF.Top) / m_SVRectF.Height)
        
        'To prevent errors, clamp saturation and value now
        If dstSaturation < 0 Then dstSaturation = 0
        If dstSaturation > 1 Then dstSaturation = 1
        If dstValue < 0 Then dstValue = 0
        If dstValue > 1 Then dstValue = 1
        
        'The y-value is squared during rendering, to decrease the amount of space taken up by extremely dark color variants
        If dstValue > 0 Then dstValue = Sqr(dstValue)
        
    End If
    
End Function

Private Sub cMouseEvents_MouseUpCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal ClickEventAlsoFiring As Boolean)
    
    m_MouseDownWheel = False
    m_MouseDownBox = False
    
    'Reset the cursor and hover behavior accordingly
    Dim tmpHue As Double, tmpSaturation As Double, tmpValue As Double
    m_MouseInsideWheel = isMouseInsideHueWheel(x, y, True, tmpHue)
    
    If m_MouseInsideWheel Then
        cMouseEvents.setSystemCursor IDC_HAND
        m_HueHover = tmpHue
    Else
        
        m_HueHover = -1
        
        m_MouseInsideBox = isMouseInsideSVBox(x, y, True, tmpSaturation, tmpValue)
        If m_MouseInsideBox Then
            cMouseEvents.setSystemCursor IDC_HAND
        Else
            cMouseEvents.setSystemCursor IDC_DEFAULT
        End If
        
    End If
    
    'Redraw the control to match
    DrawUC
    
End Sub

'The pdWindowPaint class raises this event when the navigator box needs to be redrawn.  The passed coordinates contain
' the rect returned by GetUpdateRect (but with right/bottom measurements pre-converted to width/height).
Private Sub cPainter_PaintWindow(ByVal winLeft As Long, ByVal winTop As Long, ByVal winWidth As Long, ByVal winHeight As Long)
    
    'Flip the relevant chunk of the buffer to the screen
    BitBlt UserControl.hDC, winLeft, winTop, winWidth, winHeight, m_BackBuffer.getDIBDC, winLeft, winTop, vbSrcCopy
    
End Sub

Private Sub UserControl_Initialize()
    
    If g_IsProgramRunning Then
        
        'Initialize mouse handling
        Set cMouseEvents = New pdInputMouse
        cMouseEvents.addInputTracker UserControl.hWnd, True, True, , True, True
        cMouseEvents.setSystemCursor IDC_HAND
        
        'Also start a focus detector
        Set cFocusDetector = New pdFocusDetector
        cFocusDetector.startFocusTracking Me.hWnd
        
        'Also start a flicker-free window painter
        Set cPainter = New pdWindowPainter
        cPainter.StartPainter UserControl.hWnd
        
        'Create a tooltip engine
        Set toolTipManager = New pdToolTip
    
    'In design mode, initialize a base theming class, so our paint function doesn't fail
    Else
        If g_Themer Is Nothing Then Set g_Themer = New pdVisualThemes
    End If
    
    'Draw the control at least once
    UpdateControlSize
    
End Sub

Private Sub UserControl_InitProperties()
    Color = RGB(50, 200, 255)
End Sub

'At run-time, painting is handled by PD's pdWindowPainter class.  In the IDE, however, we must rely on VB's internal paint event.
Private Sub UserControl_Paint()
    
    'Provide minimal painting within the designer
    If Not g_IsProgramRunning Then DrawUC
    
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Me.Color = PropBag.ReadProperty("Color", RGB(50, 200, 255))
End Sub

Private Sub UserControl_Resize()
    UpdateControlSize
End Sub
    
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "Color", Me.Color, RGB(50, 200, 255)
    End With
End Sub

'Call this to recreate all buffers against a changed control size.
Private Sub UpdateControlSize()
    
    'Resize the back buffer to match the container dimensions.
    If m_BackBuffer Is Nothing Then Set m_BackBuffer = New pdDIB
    If (m_BackBuffer.getDIBWidth <> UserControl.ScaleWidth) Or (m_BackBuffer.getDIBHeight <> UserControl.ScaleHeight) Then
        m_BackBuffer.createBlank UserControl.ScaleWidth, UserControl.ScaleHeight, 24
    Else
        m_BackBuffer.resetDIB 0
    End If
    
    'Recreate the color wheel, as its size is dependent on the container size
    If g_IsProgramRunning Then
    
        CreateColorWheel
    
        'Any time the hue wheel changes, the SV square must be redrawn to match
        CreateSVSquare
        
    End If
    
    'With the backbuffer and color wheel successfully created, we can finally redraw the rest of the control
    DrawUC
    
End Sub

'Create the color wheel portion of the selector.  Note that this function cannot fire until the backbuffer has been initialized,
' because it relies on that buffer for sizing.
Private Sub CreateColorWheel()
    
    'Make sure the backbuffer exists
    If (m_BackBuffer.getDIBWidth <> 0) And (m_BackBuffer.getDIBHeight <> 0) Then
    
        'For now, the color wheel DIB is always square, sized to fit the smallest dimension of the back buffer
        Dim wheelDiameter As Long
        If m_BackBuffer.getDIBWidth < m_BackBuffer.getDIBHeight Then wheelDiameter = m_BackBuffer.getDIBWidth Else wheelDiameter = m_BackBuffer.getDIBHeight
        
        If (m_WheelBuffer Is Nothing) Then Set m_WheelBuffer = New pdDIB
        If (m_WheelBuffer.getDIBWidth <> wheelDiameter) Or (m_WheelBuffer.getDIBHeight <> wheelDiameter) Then
            m_WheelBuffer.createBlank wheelDiameter, wheelDiameter, 32, 0&, 255
        Else
            GDI_Plus.GDIPlusFillDIBRect m_WheelBuffer, 0, 0, wheelDiameter, wheelDiameter, 0&, 255
        End If
        
        'We're now going to calculate the inner and outer radius of the wheel.  These are based off hard-coded padding constants,
        ' the max available diameter, and the current screen DPI.
        m_HueRadiusOuter = (CSng(wheelDiameter) / 2) - FixDPIFloat(WHEEL_PADDING)
        m_HueRadiusInner = m_HueRadiusOuter - FixDPIFloat(WHEEL_WIDTH)
        If m_HueRadiusInner < 5 Then m_HueRadiusInner = 5
        
        'We're now going to cheat a bit and use a 2D drawing hack to solve for the alpha bytes of our wheel.  The wheel image is
        ' already a black square, and atop that we're going to draw a white circle at the outer radius size, and a black circle
        ' at the inner radius size.  Both will be antialiased.  Black pixels will then be made transparent, while white pixels
        ' are fully opaque.  Gray pixels will be shaded on-the-fly.
        m_HueWheelCenterX = wheelDiameter / 2: m_HueWheelCenterY = m_HueWheelCenterX
        
        GDI_Plus.GDIPlusFillCircleToDC m_WheelBuffer.getDIBDC, m_HueWheelCenterX, m_HueWheelCenterY, m_HueRadiusOuter, RGB(255, 255, 255), 255
        GDI_Plus.GDIPlusFillCircleToDC m_WheelBuffer.getDIBDC, m_HueWheelCenterX, m_HueWheelCenterY, m_HueRadiusInner, RGB(0, 0, 0), 255
        
        'With our "alpha guidance" pixels drawn, we can now loop through the image, rendering actual hue colors as we go.
        ' For convenience, we will place hue 0 at angle 0.
        Dim hPixels() As Byte
        Dim hueSA As SAFEARRAY2D
        prepSafeArray hueSA, m_WheelBuffer
        CopyMemory ByVal VarPtrArray(hPixels()), VarPtr(hueSA), 4
        
        Dim x As Long, y As Long
        Dim r As Long, g As Long, b As Long, a As Long, aFloat As Single
        
        Dim nX As Double, nY As Double, pxAngle As Double
        
        Dim loopWidth As Long, loopHeight As Long
        loopWidth = (m_WheelBuffer.getDIBWidth - 1) * 4
        loopHeight = (m_WheelBuffer.getDIBHeight - 1)
        
        For y = 0 To loopHeight
        For x = 0 To loopWidth Step 4
            
            'Before calculating anything, check the color at this position.  (Because the image is grayscale, we only need to
            ' pull a single color value.)
            b = hPixels(x, y)
            
            'If this pixel is black, it will be forced to full transparency.  Apply that now.
            If b = 0 Then
                hPixels(x, y) = 0
                hPixels(x + 1, y) = 0
                hPixels(x + 2, y) = 0
                hPixels(x + 3, y) = 0
            
            'If this pixel is non-black, it must be colored.  Proceed with hue calculation.
            Else
            
                'Remap the coordinates so that (0, 0) represents the center of the image
                nX = (x \ 4) - m_HueWheelCenterX
                nY = y - m_HueWheelCenterY
                
                'Calculate an angle for this pixel
                pxAngle = Math_Functions.Atan2(nY, nX)
                
                'ATan2() returns an angle that is positive for counter-clockwise angles (y > 0), and negative for
                ' clockwise angles (y < 0), on the range [-Pi, +Pi].  Convert this angle to the absolute range [0, 1],
                ' which is the range used by our HSV conversion function.
                pxAngle = (pxAngle + PI) / PI_DOUBLE
                
                'Calculate an RGB triplet that corresponds to this hue (with max value and saturation)
                Color_Functions.HSVtoRGB pxAngle, 1#, 1#, r, g, b
                
                'Retrieve the "alpha" clue for this pixel
                a = hPixels(x, y)
                aFloat = CDbl(a) / 255
                
                'Premultiply alpha
                r = r * aFloat
                g = g * aFloat
                b = b * aFloat
                
                'Store the new color values
                hPixels(x, y) = b
                hPixels(x + 1, y) = g
                hPixels(x + 2, y) = r
                hPixels(x + 3, y) = a
                
            End If
        
        Next x
        Next y
        
        'With our work complete, point the array away from the DIB before VB attempts to deallocate it
        CopyMemory ByVal VarPtrArray(hPixels), 0&, 4
        
        'Mark the wheel DIB's premultiplied alpha state
        m_WheelBuffer.setInitialAlphaPremultiplicationState True
        
    End If
    
End Sub

'Create a new Saturation + Value square (the square in the middle of the UC).  The square must be redrawn whenever
' hue changes, because the hue value determines the square's appearance.
Private Sub CreateSVSquare()
    
    'Make sure the backbuffer exists
    If (m_BackBuffer.getDIBWidth <> 0) And (m_BackBuffer.getDIBHeight <> 0) Then
    
        'The SV square is a square that fits (inclusively) within the color wheel.  Basic geometry tells us that one side of the square
        ' is equal to hypotenuse * sin(45), and we know the hypotenuse already because it's the inner radius of the hue wheel.
        m_SVRectF.Width = (m_HueRadiusInner * 2) * Sin(PI / 4): m_SVRectF.Height = m_SVRectF.Width
        
        If (m_SquareBuffer Is Nothing) Then Set m_SquareBuffer = New pdDIB
        If (m_SquareBuffer.getDIBWidth <> CLng(m_SVRectF.Width)) Or (m_SquareBuffer.getDIBHeight <> CLng(m_SVRectF.Height)) Then
            m_SquareBuffer.createBlank CLng(m_SVRectF.Width), CLng(m_SVRectF.Height), 24
        Else
            m_SquareBuffer.resetDIB 0
        End If
        
        'To prevent IDE crashes, bail now during compilation
        If Not g_IsProgramRunning Then Exit Sub
        
        'We now need to fill the square with all possible saturation and value variants, in a pattern where...
        ' - The y-axis position determines value (1 -> 0)
        ' - The x-axis position determines saturation (1 -> 0)
        Dim svPixels() As Byte
        Dim svSA As SAFEARRAY2D
        prepSafeArray svSA, m_SquareBuffer
        CopyMemory ByVal VarPtrArray(svPixels()), VarPtr(svSA), 4
        
        Dim x As Long, y As Long
        Dim r As Long, g As Long, b As Long
        
        Dim loopWidth As Long, loopHeight As Long
        loopWidth = (m_SquareBuffer.getDIBWidth - 1) * 3
        loopHeight = (m_SquareBuffer.getDIBHeight - 1)
        
        Dim lineValue As Double
        
        'To improve performance, pre-calculate all value variants, so we don't need to re-calculate them in the inner loop.
        ' (They are constant for each line.)
        Dim xPresets() As Double
        ReDim xPresets(0 To loopWidth) As Double
        For x = 0 To loopWidth Step 3
            xPresets(x) = (loopWidth - x) / loopWidth
        Next x
        
        For y = 0 To loopHeight
            
            'Y-values are (obviously) consistent for each y-position
            lineValue = (loopHeight - y) / loopHeight
            lineValue = Sqr(lineValue)
            
        For x = 0 To loopWidth Step 3
            
            'The x-axis position determines saturation (1 -> 0)
            'The y-axis position determines value (1 -> 0)
            HSVtoRGB m_Hue, xPresets(x), lineValue, r, g, b
            
            svPixels(x, y) = b
            svPixels(x + 1, y) = g
            svPixels(x + 2, y) = r
            
        Next x
        Next y
        
        'With our work complete, point the ImageData() array away from the DIBs and deallocate it
        CopyMemory ByVal VarPtrArray(svPixels), 0&, 4
        
        'While we're here, let's also calculate the top-left rendering origin for the square, so we don't have to do it in the core
        ' rendering function.
        Dim tmpX As Double, tmpY As Double
        Math_Functions.convertPolarToCartesian -(3 * PI) / 4, m_HueRadiusInner, tmpX, tmpY, m_HueWheelCenterX, m_HueWheelCenterY
        m_SVRectF.Left = tmpX
        m_SVRectF.Top = tmpY
        
    End If

End Sub

'Redraw the UC.  Note that some UI elements must be created prior to calling this function (e.g. the color wheel).
Private Sub DrawUC()

    'Create the back buffer as necessary.  (This is primarily for solving IDE issues.)
    If m_BackBuffer Is Nothing Then m_BackBuffer.createBlank UserControl.ScaleWidth, UserControl.ScaleHeight, 24, RGB(255, 255, 255)
    
    If g_IsProgramRunning Then
    
        'Paint the background.
        GDI_Plus.GDIPlusFillDIBRect m_BackBuffer, 0, 0, m_BackBuffer.getDIBWidth, m_BackBuffer.getDIBHeight, g_Themer.GetThemeColor(PDTC_BACKGROUND_DEFAULT), 255
        
        'Paint the hue wheel (currently left-aligned)
        If Not (m_WheelBuffer Is Nothing) Then m_WheelBuffer.alphaBlendToDC m_BackBuffer.getDIBDC
        
        'Trace the edges of the hue wheel, to help separate the bright portions from the background.
        GDI_Plus.GDIPlusDrawCircleToDC m_BackBuffer.getDIBDC, m_HueWheelCenterX, m_HueWheelCenterY, m_HueRadiusOuter, RGB(128, 128, 128), 128
        GDI_Plus.GDIPlusDrawCircleToDC m_BackBuffer.getDIBDC, m_HueWheelCenterX, m_HueWheelCenterY, m_HueRadiusInner, RGB(128, 128, 128), 128
        
        'Paint the saturation+value square
        If Not (m_SquareBuffer Is Nothing) Then
            
            'Copy the square into place.  Note that we must use GDI+ to support subpixel positioning.
            With m_SVRectF
                GDI_Plus.GDIPlus_StretchBlt m_BackBuffer, .Left, .Top, .Width, .Height, m_SquareBuffer, 0, 0, m_SquareBuffer.getDIBWidth, m_SquareBuffer.getDIBHeight, , InterpolationModeBilinear
            End With
            
            'Trace the edges of the square, to help separate the bright portions from the background
            GDI_Plus.GDIPlusDrawRectFOutlineToDC m_BackBuffer.getDIBDC, m_SVRectF, RGB(128, 128, 128), 128, 1, True, LineJoinRound, True
            
        End If
        
        'Draw a "pie-slice" outline around the current hue value.  Start by retrieving the UI angle of the current hue value
        Dim hueAngle As Single
        hueAngle = GetUIAngleOfHue(m_Hue)
        
        'We are now going to construct a "slice-like" overlay for the current hue position.
        Dim slicePath As pdGraphicsPath
        Set slicePath = New pdGraphicsPath
        
        'The sweep of the slice should really be contingent on the radius, but for this first draft, we'll simply hard-code it.
        Dim sliceSweep As Single
        sliceSweep = 0.18
        
        'Also, the slice will extend beyond the interior and exterior edges of the hue wheel by some fixed amount (currently 0.5 pixels)
        Dim sliceExtend As Single
        sliceExtend = 0.5
        
        'Next, calculate (x, y) coordinates for the four corners of the slice.  We use these as the endpoints for the radial lines
        ' marking either side of the "slice".
        Dim x1 As Double, x2 As Double, x3 As Double, x4 As Double, y1 As Double, y2 As Double, y3 As Double, y4 As Double
        Math_Functions.convertPolarToCartesian hueAngle - (sliceSweep / 2), m_HueRadiusInner - sliceExtend, x1, y1, m_HueWheelCenterX, m_HueWheelCenterY
        Math_Functions.convertPolarToCartesian hueAngle - (sliceSweep / 2), m_HueRadiusOuter + sliceExtend, x2, y2, m_HueWheelCenterX, m_HueWheelCenterY
        Math_Functions.convertPolarToCartesian hueAngle + (sliceSweep / 2), m_HueRadiusInner - sliceExtend, x3, y3, m_HueWheelCenterX, m_HueWheelCenterY
        Math_Functions.convertPolarToCartesian hueAngle + (sliceSweep / 2), m_HueRadiusOuter + sliceExtend, x4, y4, m_HueWheelCenterX, m_HueWheelCenterY
        
        'Add those two lines to the path object, and place connecting arcs between them
        slicePath.addLine x1, y1, x2, y2
        slicePath.addArcCircular m_HueWheelCenterX, m_HueWheelCenterY, m_HueRadiusOuter + sliceExtend, RadiansToDegrees(hueAngle - (sliceSweep / 2)), RadiansToDegrees(sliceSweep)
        slicePath.addLine x4, y4, x3, y3
        slicePath.addArcCircular m_HueWheelCenterX, m_HueWheelCenterY, m_HueRadiusInner - sliceExtend, RadiansToDegrees(hueAngle + (sliceSweep / 2)), RadiansToDegrees(-sliceSweep)
        slicePath.closeCurrentFigure
        
        'Render the completed slice onto the overlay
        slicePath.strokePathToDIB_UIStyle m_BackBuffer, , m_MouseDownWheel
        
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
        If svX < COLOR_CIRCLE_CHECK Then svX = COLOR_CIRCLE_CHECK
        If svY < COLOR_CIRCLE_CHECK Then svY = COLOR_CIRCLE_CHECK
        If svX > (m_SVRectF.Width - (COLOR_CIRCLE_CHECK + 1)) Then svX = (m_SVRectF.Width - (COLOR_CIRCLE_CHECK + 1))
        If svY > (m_SVRectF.Height - (COLOR_CIRCLE_CHECK + 1)) Then svY = (m_SVRectF.Height - (COLOR_CIRCLE_CHECK + 1))
        
        'Pad the circle by the current SV square's offset
        svX = svX + m_SVRectF.Left
        svY = svY + m_SVRectF.Top
        
        'Draw a canvas-style circle around that point
        GDI_Plus.GDIPlusDrawCanvasCircle m_BackBuffer.getDIBDC, svX, svY, COLOR_CIRCLE_RADIUS, , m_MouseDownBox
        
    'In the designer, draw a focus rect around the control; this is minimal feedback required for positioning
    Else
        
        Dim tmpRect As RECT
        With tmpRect
            .Left = 0
            .Top = 0
            .Right = m_BackBuffer.getDIBWidth
            .Bottom = m_BackBuffer.getDIBHeight
        End With
        
        DrawFocusRect m_BackBuffer.getDIBDC, tmpRect
    
    End If
    
    'Paint the final result to the screen, as relevant
    If g_IsProgramRunning Then
        cPainter.RequestRepaint
    Else
        BitBlt UserControl.hDC, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, m_BackBuffer.getDIBDC, 0, 0, vbSrcCopy
    End If

End Sub

'Given a hue on the range [0, 1], return a GDIPlus-friendly UI angle for the hue wheel
Private Function GetUIAngleOfHue(ByVal srcHue As Single) As Single
    
    'A hue of "0" corresponds to an angle of Pi.  A hue of "1" corresponds to an angle of -Pi (hue is circular).
    GetUIAngleOfHue = (srcHue * PI_DOUBLE) - PI
    
End Function

'Convert the control's current HSV triplet into a corresponding RGB long
Private Function GetCurrentRGB() As Long
    Dim r As Long, g As Long, b As Long
    Color_Functions.HSVtoRGB m_Hue, m_Saturation, m_Value, r, g, b
    GetCurrentRGB = RGB(r, g, b)
End Function

'Due to complex interactions between user controls and PD's translation engine, tooltips require this dedicated function.
' (IMPORTANT NOTE: the tooltip class will handle translations automatically.  Always pass the original English text!)
Public Sub AssignTooltip(ByVal newTooltip As String, Optional ByVal newTooltipTitle As String, Optional ByVal newTooltipIcon As TT_ICON_TYPE = TTI_NONE)
    toolTipManager.SetTooltip Me.hWnd, Me.ContainerHwnd, newTooltip, newTooltipTitle, newTooltipIcon
End Sub

'External functions can call this to request a redraw.  This is helpful for live-updating theme settings, as in the Preferences dialog,
' and/or retranslating any text against the current language.
Public Sub UpdateAgainstCurrentTheme()
    
    'Update the tooltip, if any
    If g_IsProgramRunning Then toolTipManager.UpdateAgainstCurrentTheme
        
    'Redraw the control (in case anything has changed)
    UpdateControlSize
    
End Sub
