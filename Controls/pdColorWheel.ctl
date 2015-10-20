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
'Last updated: 19/October/15
'Last update: start initial build
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
Public Event ColorChanged()

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
Private m_WheelBuffer As pdDIB

'These values will be TRUE while the mouse is inside various parts of the UC
Private m_MouseInsideWheel As Boolean, m_MouseInsideBox As Boolean

'API technique for drawing a focus rectangle; used only for designer mode (see the Paint method for details)
Private Declare Function DrawFocusRect Lib "user32" (ByVal hDC As Long, lpRect As RECT) As Long

'Padding (in pixels) between the edges of the user control and the color wheel.  Automatically adjusted for DPI
' at run-time.  Note that this needs to be non-zero, because the padding area is used to render the "slice" overlay
' showing the user's current hue selection.
Private Const WHEEL_PADDING As Long = 3

'Width (in pixels) of the hue wheel.  This width is applied along the radial axis.
Private Const WHEEL_WIDTH As Single = 15#

'The inner and outer radius of the hue wheel.  These are calculated by the CreateColorWheel function and cached here,
' as a convenience for subsequent mouse hit-testing.
Private m_HueRadiusInner As Single, m_HueRadiusOuter As Single

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

Public Property Get ContainerHwnd() As Long
    ContainerHwnd = UserControl.ContainerHwnd
End Property

'When the control receives focus, relay the event externally
Private Sub cFocusDetector_GotFocusReliable()
    RaiseEvent GotFocusAPI
End Sub

'When the control loses focus, relay the event externally
Private Sub cFocusDetector_LostFocusReliable()
    RaiseEvent LostFocusAPI
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
        cPainter.startPainter UserControl.hWnd
        
        'Create a tooltip engine
        Set toolTipManager = New pdToolTip
    
    'In design mode, initialize a base theming class, so our paint function doesn't fail
    Else
        If g_Themer Is Nothing Then Set g_Themer = New pdVisualThemes
    End If
    
    'Draw the control at least once
    UpdateControlSize
    
End Sub

'At run-time, painting is handled by PD's pdWindowPainter class.  In the IDE, however, we must rely on VB's internal paint event.
Private Sub UserControl_Paint()
    
    'Provide minimal painting within the designer
    If Not g_IsProgramRunning Then DrawUC
    
End Sub

Private Sub UserControl_Resize()
    UpdateControlSize
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
    If g_IsProgramRunning Then CreateColorWheel
    
    'With the backbuffer and color wheel successfully created, we can finally redraw the rest of the control
    DrawUC
    
End Sub

'Create the color wheel portion of the selector.  Note that this function cannot fire until the backbuffer has been initialized,
' because it relies on that buffer for sizing.
Private Sub CreateColorWheel()
    
    'Make sure the backbuffer exists
    If (m_BackBuffer.getDIBWidth <> 0) And (m_BackBuffer.getDIBHeight <> 0) Then
    
        'For now, the color wheel is always a square, sized to fit the smallest dimension of the back buffer
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
        Dim cX As Single, cY As Single
        cX = wheelDiameter / 2: cY = cX
        
        GDI_Plus.GDIPlusFillCircleToDC m_WheelBuffer.getDIBDC, cX, cY, m_HueRadiusOuter, RGB(255, 255, 255), 255
        GDI_Plus.GDIPlusFillCircleToDC m_WheelBuffer.getDIBDC, cX, cY, m_HueRadiusInner, RGB(0, 0, 0), 255
        
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
                nX = (x \ 4) - cX
                nY = y - cY
                
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
        
    End If
    
End Sub

'Redraw the UC.  Note that some UI elements must be created prior to calling this function (e.g. the color wheel).
Private Sub DrawUC()

    'Create the back buffer as necessary.  (This is primarily for solving IDE issues.)
    If m_BackBuffer Is Nothing Then m_BackBuffer.createBlank UserControl.ScaleWidth, UserControl.ScaleHeight, 24, RGB(255, 255, 255)
    
    If g_IsProgramRunning Then
    
        'Paint the background.
        GDI_Plus.GDIPlusFillDIBRect m_BackBuffer, 0, 0, m_BackBuffer.getDIBWidth, m_BackBuffer.getDIBHeight, g_Themer.getThemeColor(PDTC_BACKGROUND_DEFAULT), 255
        
        'Paint the hue wheel (currently left-aligned)
        If Not (m_WheelBuffer Is Nothing) Then m_WheelBuffer.alphaBlendToDC m_BackBuffer.getDIBDC
        
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
        cPainter.requestRepaint
    Else
        BitBlt UserControl.hDC, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, m_BackBuffer.getDIBDC, 0, 0, vbSrcCopy
    End If

End Sub

'Due to complex interactions between user controls and PD's translation engine, tooltips require this dedicated function.
' (IMPORTANT NOTE: the tooltip class will handle translations automatically.  Always pass the original English text!)
Public Sub AssignTooltip(ByVal newTooltip As String, Optional ByVal newTooltipTitle As String, Optional ByVal newTooltipIcon As TT_ICON_TYPE = TTI_NONE)
    toolTipManager.setTooltip Me.hWnd, Me.ContainerHwnd, newTooltip, newTooltipTitle, newTooltipIcon
End Sub

'External functions can call this to request a redraw.  This is helpful for live-updating theme settings, as in the Preferences dialog,
' and/or retranslating any text against the current language.
Public Sub UpdateAgainstCurrentTheme()
    
    'Update the tooltip, if any
    If g_IsProgramRunning Then toolTipManager.UpdateAgainstCurrentTheme
        
    'Redraw the control (in case anything has changed)
    UpdateControlSize
    
End Sub
    
