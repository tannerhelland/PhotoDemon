VERSION 5.00
Begin VB.UserControl pdColorVariants 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   1980
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2385
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   132
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   159
   ToolboxBitmap   =   "pdColorVariants.ctx":0000
End
Attribute VB_Name = "pdColorVariants"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon "Color Variants" color selector
'Copyright 2015-2015 by Tanner Helland
'Created: 22/October/15
'Last updated: 23/October/15
'Last update: switch to a pure path-based system for rendering and hit-detection
'
'In 7.0, a "color selector" panel was added to the right-side toolbar.  Unlike PD's single-color color selector,
' this panel is designed to provide a quick, on-canvas-friendly mechanism for rapidly switching colors.
'
'In particular, this "color variant" color selector provides a way to quickly "nudge" a color toward a nearby
' variants.  It uses an original design, which is always sketchy, but the goal here is to save the poor artist
' from needing to drop into a separate color dialog (at just about any cost!).
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

'This back buffer is for the fully composited control; it is what gets copied to the screen on Paint events.
Private m_BackBuffer As pdDIB

'These values help the central renderer know where the mouse is, so we can draw various on-screen indicators.
' If set to -1, the mouse is not inside any box.
Private m_MouseInsideRegion As Long

'API technique for drawing a focus rectangle; used only for designer mode (see the Paint method for details)
Private Declare Function DrawFocusRect Lib "user32" (ByVal hDC As Long, lpRect As RECT) As Long

'Size (in pixels) of the variant selectors surrounding the primary color box.  This must be manually adjusted for
' DPI settings at run-time.  Note that at least 1px is lost to borders on either side, as well.
Private Const VARIANT_BOX_SIZE As Long = 16

'The list of variant selectors.  With the exception of the primary selector (which gets preference at position 0),
' these start in the top-left and move clockwise around the control border.
Private Const NUM_OF_VARIANTS = 13

Private Enum COLOR_VARIANTS
    CV_Primary = 0
    CV_HueUp = 1
    CV_SaturationUp = 2
    CV_ValueUp = 3
    CV_RedUp = 4
    CV_GreenUp = 5
    CV_BlueUp = 6
    CV_ValueDown = 7
    CV_SaturationDown = 8
    CV_HueDown = 9
    CV_BlueDown = 10
    CV_GreenDown = 11
    CV_RedDown = 12
End Enum

#If False Then
    Private Const CV_Primary = 0, CV_HueUp = 1, CV_ValueUp = 2, CV_SaturationUp = 3, CV_RedUp = 4, CV_GreenUp = 5, CV_BlueUp = 6
    Private Const CV_SaturationDown = 7, CV_ValueDown = 8, CV_HueDown = 9, CV_BlueDown = 10, CV_GreenDown = 11, CV_RedDown = 12
#End If

'Current color values of each variant.  These are pre-calculated when the primary color changes, to spare us having
' to calculate them in the rendering loop.
Private m_ColorList() As Long

'Initially, we used a collection of RectF objects to house the coordinates for each subregion, but to increase flexibility,
' these were later moved to generic path objects.  This is how we are able to provide both rectangular and circular appearances,
' with almost no changes to the underlying code.
Private m_ColorRegions() As pdGraphicsPath

'This control supports both rectangular and circular shapes
Public Enum COLOR_WHEEL_SHAPE
    CWS_Circular = 0
    CWS_Rectangular = 1
End Enum

#If False Then
    Private Const CWS_Circular = 0, CWS_Rectangular = 1
#End If

Private m_ControlShape As COLOR_WHEEL_SHAPE

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

Public Property Get containerHwnd() As Long
    containerHwnd = UserControl.containerHwnd
End Property

Public Property Get Color() As Long
    Color = m_ColorList(0)
End Property

Public Property Let Color(ByVal newColor As Long)
    
    m_ColorList(0) = newColor
    MakeNewTooltip CV_Primary
    
    'Recalculate all color variants, then redraw the control
    CalculateVariantColors
    DrawUC
    
    RaiseEvent ColorChanged(m_ColorList(0), False)
    PropertyChanged "Color"
    
End Property

Public Property Get WheelShape() As COLOR_WHEEL_SHAPE
    WheelShape = m_ControlShape
End Property

Public Property Let WheelShape(ByVal newShape As COLOR_WHEEL_SHAPE)
    If m_ControlShape <> newShape Then
        m_ControlShape = newShape
        UpdateControlSize
        PropertyChanged "WheelShape"
    End If
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
    
        'See if the mouse cursor is inside a box
        m_MouseInsideRegion = GetRegionFromPoint(x, y)
        
        If m_MouseInsideRegion >= 0 Then
        
            'If the primary color box is clicked, raise a full color selection dialog
            If m_MouseInsideRegion = 0 Then
                DisplayColorSelection
            Else
                m_ColorList(0) = m_ColorList(m_MouseInsideRegion)
            End If
            
            'Recalculate all color variants to match the new color (if any)
            MakeNewTooltip m_MouseInsideRegion
            CalculateVariantColors
            
            'Redraw the control to reflect this new color
            DrawUC
            
            'Raise an event to match
            RaiseEvent ColorChanged(m_ColorList(0), True)
        
        End If
        
    End If
    
End Sub

Private Sub cMouseEvents_MouseLeave(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    cMouseEvents.setSystemCursor IDC_DEFAULT
    m_MouseInsideRegion = -1
    DrawUC
End Sub

Private Sub cMouseEvents_MouseMoveCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    
    'Calculate a new hovered box ID, if any
    Dim oldMouseIndex As Long
    oldMouseIndex = m_MouseInsideRegion
    m_MouseInsideRegion = GetRegionFromPoint(x, y)
    
    'Modify the cursor to match
    If (m_MouseInsideRegion >= 0) Then cMouseEvents.setSystemCursor IDC_HAND Else cMouseEvents.setSystemCursor IDC_DEFAULT
    
    'If the box ID has changed, update the tooltip and redraw the control to match
    If (m_MouseInsideRegion <> oldMouseIndex) Then
        MakeNewTooltip m_MouseInsideRegion
        DrawUC
    End If
    
End Sub

'Given an (x, y) coordinate pair from the mouse, return the index of the containing rect (if any).
' Returns -1 if the point lies outside all rects.
Private Function GetRegionFromPoint(ByVal x As Single, ByVal y As Single) As Long

    GetRegionFromPoint = -1
    
    Dim i As Long
    For i = 0 To NUM_OF_VARIANTS - 1
        
        If m_ColorRegions(i).isPointInsidePath(x, y) Then
            GetRegionFromPoint = i
            Exit Function
        End If
        
    Next i

End Function

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
    
    m_MouseInsideRegion = -1
    
    'Prep the various color variant lists
    ReDim m_ColorList(0 To NUM_OF_VARIANTS - 1) As Long
    ReDim m_ColorRegions(0 To NUM_OF_VARIANTS - 1) As pdGraphicsPath
    
    Dim i As Long
    For i = 0 To NUM_OF_VARIANTS - 1
        Set m_ColorRegions(i) = New pdGraphicsPath
    Next i
    
    CalculateVariantColors
    
    'Draw the control at least once
    UpdateControlSize
    
End Sub

Private Sub UserControl_InitProperties()
    Color = vbRed
    WheelShape = CWS_Circular
End Sub

'At run-time, painting is handled by PD's pdWindowPainter class.  In the IDE, however, we must rely on VB's internal paint event.
Private Sub UserControl_Paint()
    
    'Provide minimal painting within the designer
    If Not g_IsProgramRunning Then DrawUC
    
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Me.Color = PropBag.ReadProperty("Color", vbRed)
    Me.WheelShape = PropBag.ReadProperty("WheelShape", CWS_Circular)
End Sub

Private Sub UserControl_Resize()
    UpdateControlSize
End Sub
    
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "Color", Me.Color, vbRed
        .WriteProperty "WheelShape", Me.WheelShape, CWS_Circular
    End With
End Sub

'Call this to force a display of the color window.  Note that it's *public*, so outside callers can raise dialogs, too.
Public Sub DisplayColorSelection()
    
    'Store the current color
    Dim newColor As Long, oldColor As Long
    oldColor = m_ColorList(0)
    m_MouseInsideRegion = -1
    
    'Use the default color dialog to select a new color
    If showColorDialog(newColor, oldColor, Nothing) Then
        m_ColorList(0) = newColor
    Else
        m_ColorList(0) = oldColor
    End If
    
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
    
    If g_IsProgramRunning Then
        
        'Re-calculate all control subregions.  This is a little confusing (okay, a LOT confusing), but basically we want to
        ' create an evenly spaced border around the central color rect, with subdivided regions that provide some dynamic
        ' color variants for the user to choose from.
        Dim i As Long
        For i = 0 To NUM_OF_VARIANTS - 1
            m_ColorRegions(i).resetPath
        Next i
        
        'Leave a little room around the control, so we can draw chunky borders around hovered sub-regions.
        Dim ucLeft As Long, ucTop As Long, ucBottom As Long, ucRight As Long
        ucLeft = 1
        ucTop = 1
        ucBottom = m_BackBuffer.getDIBHeight - 2
        ucRight = m_BackBuffer.getDIBWidth - 2
        
        'How we actually create the regions varies depending on the current control orientation.
        If m_ControlShape = CWS_Circular Then
            CreateSubregions_Circular ucLeft, ucTop, ucBottom, ucRight
        Else
            CreateSubregions_Rectangular ucLeft, ucTop, ucBottom, ucRight
        End If
        
    End If
    
    'With the backbuffer and rects successfully constructed, we can finally redraw the control
    DrawUC
    
End Sub

'Create a rectangular grid-based variant control
Private Sub CreateSubregions_Rectangular(ByVal ucLeft As Long, ByVal ucTop As Long, ByVal ucBottom As Long, ByVal ucRight As Long)

    'First, make sure our border size is DPI-aware
    Dim dpiAwareBorderSize As Long
    dpiAwareBorderSize = FixDPI(VARIANT_BOX_SIZE)
    
    'For this control layout, we are going to use a temporary rect collection to define the position of all color variants.
    ' This simplifies things a bit, and when we're done, we'll simply copy all rects into the master pdGraphicsPath array.
    Dim colorRects() As RECTF
    ReDim colorRects(0 To NUM_OF_VARIANTS - 1) As RECTF
    
    'Start by calculating the primary color rect.  It is the focus of the control, and its position affects all
    ' subsequent controls.
    With colorRects(CV_Primary)
        .Left = ucLeft + dpiAwareBorderSize
        .Top = ucTop + dpiAwareBorderSize
        .Width = (ucRight - dpiAwareBorderSize) - .Left
        .Height = (ucBottom - dpiAwareBorderSize) - .Top
    End With
    
    'Next, loop through rects that share one or more position values.
    Dim i As Long
    
    For i = CV_HueUp To CV_ValueUp
        With colorRects(i)
            .Top = ucTop
            .Height = dpiAwareBorderSize
        End With
    Next i
    
    For i = CV_ValueUp To CV_ValueDown
        With colorRects(i)
            .Left = colorRects(CV_Primary).Left + colorRects(CV_Primary).Width
            .Width = dpiAwareBorderSize
        End With
    Next i
    
    For i = CV_ValueDown To CV_HueDown
        With colorRects(i)
            .Top = colorRects(CV_Primary).Top + colorRects(CV_Primary).Height
            .Height = dpiAwareBorderSize
        End With
    Next i
    
    For i = CV_HueDown To CV_RedDown
        With colorRects(i)
            .Left = ucLeft
            .Width = dpiAwareBorderSize
        End With
    Next i
    colorRects(CV_HueUp).Left = ucLeft
    colorRects(CV_HueUp).Width = dpiAwareBorderSize
    
    'Next, we must manually calculate all remaining rect positions.
    
    'The HSV boxes split their width evenly across the control's available space
    Dim hsvWidth As Single
    hsvWidth = (ucRight - ucLeft) / 3
    
    For i = CV_HueUp To CV_ValueUp
        colorRects(i).Width = hsvWidth
    Next i
    For i = CV_ValueDown To CV_HueDown
        colorRects(i).Width = hsvWidth
    Next i
    
    colorRects(CV_HueUp).Left = ucLeft
    colorRects(CV_SaturationUp).Left = ucLeft + hsvWidth
    colorRects(CV_ValueUp).Left = ucLeft + hsvWidth * 2
    colorRects(CV_HueDown).Left = colorRects(CV_HueUp).Left
    colorRects(CV_SaturationDown).Left = colorRects(CV_SaturationUp).Left
    colorRects(CV_ValueDown).Left = colorRects(CV_ValueUp).Left
    
    'The only remaining rects to calculate are the RGB boxes that sit on either side of the main color box.
    ' Their vertical positioning is equally split between the 3 boxes, so it is contingent on the control's size
    ' as a whole.
    Dim rgbHeight As Single
    rgbHeight = colorRects(CV_Primary).Height / 3
    
    'Start by assigning all boxes a uniform height
    For i = CV_RedUp To CV_BlueUp
        colorRects(i).Height = rgbHeight
    Next i
    For i = CV_BlueDown To CV_RedDown
        colorRects(i).Height = rgbHeight
    Next i
    
    'Next, commit the top positions, which vary by index
    colorRects(CV_RedUp).Top = colorRects(CV_Primary).Top
    colorRects(CV_GreenUp).Top = colorRects(CV_Primary).Top + rgbHeight
    colorRects(CV_BlueUp).Top = colorRects(CV_Primary).Top + rgbHeight * 2
    
    colorRects(CV_RedDown).Top = colorRects(CV_RedUp).Top
    colorRects(CV_GreenDown).Top = colorRects(CV_GreenUp).Top
    colorRects(CV_BlueDown).Top = colorRects(CV_BlueUp).Top
    
    'With the color rects successfully constructed, we can now add them to our master path collection
    For i = CV_Primary To NUM_OF_VARIANTS - 1
        m_ColorRegions(i).addRectangle_RectF colorRects(i)
    Next i

End Sub

Private Sub CreateSubregions_Circular(ByVal ucLeft As Long, ByVal ucTop As Long, ByVal ucBottom As Long, ByVal ucRight As Long)

    'First, make sure our border size is DPI-aware
    Dim dpiAwareBorderSize As Long
    dpiAwareBorderSize = FixDPI(VARIANT_BOX_SIZE)
    
    'Constructing circular sub-regions actually involves less code than rectangular ones, because they're spaced perfectly evenly,
    ' so we can easily construct them in a loop.
    
    'Start by calculating basic arc and circle values
    Dim minDimension As Single
    If ucRight - ucLeft < ucBottom - ucTop Then
        minDimension = ucRight - ucLeft
    Else
        minDimension = ucBottom - ucTop
    End If
    
    Dim centerX As Single, centerY As Single
    centerX = ucLeft + (ucRight - ucLeft) / 2
    centerY = ucTop + (ucBottom - ucTop) / 2
    
    Dim innerRadius As Double, outerRadius As Double
    outerRadius = (minDimension / 2)
    innerRadius = (minDimension / 2) - dpiAwareBorderSize
    
    'The primary circle is the only subregion that receives a special construction method.
    m_ColorRegions(CV_Primary).addCircle centerX, centerY, innerRadius
    
    'All subregions are added uniformly, in a loop
    Dim startAngle As Single, sweepAngle As Single
    startAngle = 180
    sweepAngle = 30
    
    Dim x1 As Double, x2 As Double, x3 As Double, x4 As Double, y1 As Double, y2 As Double, y3 As Double, y4 As Double
    
    Dim i As Long
    For i = 1 To NUM_OF_VARIANTS - 1
    
        'Calculate (x, y) coordinates for the four corners of this subregion.  We use these as the endpoints for the radial lines
        ' marking either side of the "slice".
        Math_Functions.convertPolarToCartesian Math_Functions.DegreesToRadians(startAngle), innerRadius, x1, y1, centerX, centerY
        Math_Functions.convertPolarToCartesian Math_Functions.DegreesToRadians(startAngle), outerRadius, x2, y2, centerX, centerY
        Math_Functions.convertPolarToCartesian Math_Functions.DegreesToRadians(startAngle + sweepAngle), innerRadius, x3, y3, centerX, centerY
        Math_Functions.convertPolarToCartesian Math_Functions.DegreesToRadians(startAngle + sweepAngle), outerRadius, x4, y4, centerX, centerY
        
        'Add the two divider lines to the current path object, and place connecting arcs between them
        m_ColorRegions(i).addLine x1, y1, x2, y2
        m_ColorRegions(i).addArcCircular centerX, centerY, outerRadius, startAngle, sweepAngle
        m_ColorRegions(i).addLine x4, y4, x3, y3
        m_ColorRegions(i).addArcCircular centerX, centerY, innerRadius, startAngle + sweepAngle, -sweepAngle
        m_ColorRegions(i).closeCurrentFigure
        
        'Offset the startAngle for the next slice
        startAngle = startAngle + sweepAngle
        
    Next i

End Sub

'Any time the primary color changes (for whatever reason, external or internal), new variant colors must be calculated.
' Call this function to auto-calculate them, but try to do it only when necessary, as there's a lot of math involved.
Private Sub CalculateVariantColors()
    
    'The primary color serves as the base color for all subsequent calculations.  Retrieve its RGB and HSV quads now.
    Dim rPrimary As Long, gPrimary As Long, bPrimary As Long, hPrimary As Double, sPrimary As Double, vPrimary As Double
    rPrimary = Color_Functions.ExtractR(m_ColorList(CV_Primary))
    gPrimary = Color_Functions.ExtractG(m_ColorList(CV_Primary))
    bPrimary = Color_Functions.ExtractB(m_ColorList(CV_Primary))
    Color_Functions.RGBtoHSV rPrimary, gPrimary, bPrimary, hPrimary, sPrimary, vPrimary
    
    'We now need to calculate new RGB values.  How we do this varies by variant, obviously!
    Dim rNew As Long, gNew As Long, bNew As Long, hNew As Double, sNew As Double, vNew As Double
    Dim rFloat As Double, gFloat As Double, bFloat As Double
    Dim grayNew As Long
    
    'Also, during testing I'm experimenting with different increment amounts for HSV and RGB adjustments
    Dim rgbChange As Long, svChange As Double, hChange As Double
    rgbChange = 20
    svChange = 0.08
    hChange = 0.03
    
    Dim i As COLOR_VARIANTS
    For i = CV_HueUp To CV_RedDown
        
        rNew = rPrimary: gNew = gPrimary: bNew = bPrimary
        rFloat = rNew / 255: gFloat = gNew / 255: bFloat = bNew / 255
        hNew = hPrimary: sNew = sPrimary: vNew = vPrimary
        
        Select Case i
        
            Case CV_HueUp
                hNew = hNew + hChange
                If hNew > 1 Then hNew = 1 - hNew
                Color_Functions.HSVtoRGB hNew, sNew, vNew, rNew, gNew, bNew
                
            Case CV_SaturationUp
                
                'Use a fake saturation calculation
                grayNew = Color_Functions.getHQLuminance(rNew, gNew, bNew)
                rNew = rNew + (rNew - grayNew) * 0.4
                gNew = gNew + (gNew - grayNew) * 0.4
                bNew = bNew + (bNew - grayNew) * 0.4
                rNew = Math_Functions.ClampL(rNew, 0, 255)
                gNew = Math_Functions.ClampL(gNew, 0, 255)
                bNew = Math_Functions.ClampL(bNew, 0, 255)
            
            Case CV_ValueUp
                
                'Use a fake value calculation
                rNew = Math_Functions.ClampL(rNew + rgbChange, 0, 255)
                gNew = Math_Functions.ClampL(gNew + rgbChange, 0, 255)
                bNew = Math_Functions.ClampL(bNew + rgbChange, 0, 255)
                
            Case CV_RedUp
                rNew = Math_Functions.ClampL(rNew + rgbChange, 0, 255)
                
            Case CV_GreenUp
                gNew = Math_Functions.ClampL(gNew + rgbChange, 0, 255)
                
            Case CV_BlueUp
                bNew = Math_Functions.ClampL(bNew + rgbChange, 0, 255)
                
            Case CV_ValueDown
                
                'Use a fake value calculation
                rNew = Math_Functions.ClampL(rNew - rgbChange, 0, 255)
                gNew = Math_Functions.ClampL(gNew - rgbChange, 0, 255)
                bNew = Math_Functions.ClampL(bNew - rgbChange, 0, 255)
            
            Case CV_SaturationDown
                
                'Use a fake saturation calculation
                grayNew = Color_Functions.getHQLuminance(rNew, gNew, bNew)
                rNew = rNew + (grayNew - rNew) * 0.3
                gNew = gNew + (grayNew - gNew) * 0.3
                bNew = bNew + (grayNew - bNew) * 0.3
                rNew = Math_Functions.ClampL(rNew, 0, 255)
                gNew = Math_Functions.ClampL(gNew, 0, 255)
                bNew = Math_Functions.ClampL(bNew, 0, 255)
            
            Case CV_HueDown
                hNew = hNew - hChange
                If hNew < 0 Then hNew = 1 + hNew
                Color_Functions.HSVtoRGB hNew, sNew, vNew, rNew, gNew, bNew
            
            Case CV_BlueDown
                bNew = Math_Functions.ClampL(bNew - rgbChange, 0, 255)
                
            Case CV_GreenDown
                gNew = Math_Functions.ClampL(gNew - rgbChange, 0, 255)
                
            Case CV_RedDown
                rNew = Math_Functions.ClampL(rNew - rgbChange, 0, 255)
        
        End Select
        
        'Cache the new RGB values
        m_ColorList(i) = RGB(rNew, gNew, bNew)
    
    Next i
    
    'After recreating color values, the control must be redrawn, but we leave this to our caller to handle
    
End Sub

'Redraw the UC.  Note that some UI elements must be created prior to calling this function (e.g. the color wheel).
Private Sub DrawUC()

    'Create the back buffer as necessary.  (This is primarily for solving IDE issues.)
    If m_BackBuffer Is Nothing Then m_BackBuffer.createBlank UserControl.ScaleWidth, UserControl.ScaleHeight, 24, RGB(255, 255, 255)
    
    If g_IsProgramRunning Then
    
        'Paint the background.
        GDI_Plus.GDIPlusFillDIBRect m_BackBuffer, 0, 0, m_BackBuffer.getDIBWidth, m_BackBuffer.getDIBHeight, g_Themer.getThemeColor(PDTC_BACKGROUND_DEFAULT), 255
        
        Dim fillColor As Long, borderColor As Long, borderPen As Long
        borderColor = g_Themer.getThemeColor(PDTC_GRAY_DEFAULT)
        
        'We can reuse a single border pen for all sub-paths
        borderPen = GDI_Plus.getGDIPlusPenHandle(borderColor, , , , LineJoinMiter)
        
        'Draw each subregion in turn, filling it first, then tracing its borders.
        Dim i As Long, regionBrush As Long
        For i = CV_Primary To CV_RedDown
            
            regionBrush = GDI_Plus.getGDIPlusSolidBrushHandle(m_ColorList(i), 255)
            m_ColorRegions(i).fillPathToDIB_BareBrush regionBrush, m_BackBuffer
            GDI_Plus.releaseGDIPlusBrush regionBrush
            
            m_ColorRegions(i).strokePathToDIB_BarePen borderPen, m_BackBuffer
            
        Next i
        
        'Draw a special outline around the central primary color, to help it stand out more
        m_ColorRegions(CV_Primary).strokePathToDIB_UIStyle m_BackBuffer, , False, False, LineJoinMiter
        
        'Release any remaining GDI+ objects
        GDI_Plus.releaseGDIPlusPen borderPen
        
        'If a subregion is currently hovered, trace it with a highlight outline.
        If m_MouseInsideRegion >= 0 Then
            borderColor = g_Themer.getThemeColor(PDTC_ACCENT_DEFAULT)
            borderPen = GDI_Plus.getGDIPlusPenHandle(borderColor, , 3, , LineJoinMiter)
            m_ColorRegions(m_MouseInsideRegion).strokePathToDIB_BarePen borderPen, m_BackBuffer
            GDI_Plus.releaseGDIPlusPen borderPen
        End If
        
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

'When the currently hovered color variant changes, we want to assign a new tooltip to the control
Private Sub MakeNewTooltip(ByVal activeIndex As COLOR_VARIANTS)
    
    'Failsafe for compile-time errors when properties are written
    If Not g_IsProgramRunning Then Exit Sub
    
    Dim toolString As String, hexString As String, rgbString As String
    
    If activeIndex >= 0 Then
        hexString = "#" & UCase(Color_Functions.getHexStringFromRGB(m_ColorList(activeIndex)))
        rgbString = Color_Functions.ExtractR(m_ColorList(activeIndex)) & ", " & Color_Functions.ExtractG(m_ColorList(activeIndex)) & ", " & Color_Functions.ExtractB(m_ColorList(activeIndex))
        toolString = hexString & vbCrLf & rgbString
        If activeIndex = CV_Primary Then toolString = toolString & vbCrLf & g_Language.TranslateMessage("Click to enter a full color selection screen.")
    End If
    
    Select Case activeIndex
        
        Case CV_Primary
            Me.AssignTooltip toolString, "Current color"
        
        Case CV_HueUp
            Me.AssignTooltip toolString, "Rotate hue clockwise"
                
        Case CV_SaturationUp
            Me.AssignTooltip toolString, "Increase saturation"
            
        Case CV_ValueUp
            Me.AssignTooltip toolString, "Increase luminance"
            
        Case CV_RedUp
            Me.AssignTooltip toolString, "Increase red"
            
        Case CV_GreenUp
            Me.AssignTooltip toolString, "Increase green"
            
        Case CV_BlueUp
            Me.AssignTooltip toolString, "Increase blue"
            
        Case CV_ValueDown
            Me.AssignTooltip toolString, "Decrease luminance"
            
        Case CV_SaturationDown
            Me.AssignTooltip toolString, "Decrease saturation"
            
        Case CV_HueDown
            Me.AssignTooltip toolString, "Rotate hue counterclockwise"
            
        Case CV_BlueDown
            Me.AssignTooltip toolString, "Decrease blue"
            
        Case CV_GreenDown
            Me.AssignTooltip toolString, "Decrease green"
            
        Case CV_RedDown
            Me.AssignTooltip toolString, "Decrease red"
        
        Case Else
            Me.AssignTooltip ""
                
    End Select
    
End Sub

'Due to complex interactions between user controls and PD's translation engine, tooltips require this dedicated function.
' (IMPORTANT NOTE: the tooltip class will handle translations automatically.  Always pass the original English text!)
Public Sub AssignTooltip(ByVal newTooltip As String, Optional ByVal newTooltipTitle As String, Optional ByVal newTooltipIcon As TT_ICON_TYPE = TTI_NONE)
    toolTipManager.setTooltip Me.hWnd, Me.containerHwnd, newTooltip, newTooltipTitle, newTooltipIcon
End Sub

'External functions can call this to request a redraw.  This is helpful for live-updating theme settings, as in the Preferences dialog,
' and/or retranslating any text against the current language.
Public Sub UpdateAgainstCurrentTheme()
    
    'Update the tooltip, if any
    If g_IsProgramRunning Then toolTipManager.UpdateAgainstCurrentTheme
        
    'Redraw the control (in case anything has changed)
    UpdateControlSize
    
End Sub

