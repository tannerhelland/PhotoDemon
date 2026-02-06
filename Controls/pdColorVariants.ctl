VERSION 5.00
Begin VB.UserControl pdColorVariants 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   1980
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2385
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
'Copyright 2015-2026 by Tanner Helland
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
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Just like PD's old color selector, this control will raise a ColorChanged event after user interactions.
Public Event ColorChanged(ByVal newColor As Long, ByVal srcIsInternal As Boolean)
Public Event GotFocusAPI()
Public Event LostFocusAPI()

'These values help the central renderer know where the mouse is, so we can draw various on-screen indicators.
' If set to -1, the mouse is not inside any box.
Private m_MouseInsideRegion As Long

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

'The same color list as m_ColorList(), but color-managed.  This is used for painting the on-screen appearance *only*.
' Never retrieve these RGB values.
Private m_ColorDisplay() As Long

'Initially, we used a collection of RectF objects to house the coordinates for each subregion, but to increase flexibility,
' these were later moved to generic path objects.  This is how we are able to provide both rectangular and circular appearances,
' with almost no changes to the underlying code.
Private m_ColorRegions() As pd2DPath

'This control supports both rectangular and circular shapes
Public Enum COLOR_WHEEL_SHAPE
    CWS_Circular = 0
    CWS_Rectangular = 1
End Enum

#If False Then
    Private Const CWS_Circular = 0, CWS_Rectangular = 1
#End If

Private m_ControlShape As COLOR_WHEEL_SHAPE

'User control support class.  Historically, many classes (and associated subclassers) were required by each user control,
' but I've since wrapped these into a single central support class.
Private WithEvents ucSupport As pdUCSupport
Attribute ucSupport.VB_VarHelpID = -1

'Local list of themable colors.  This list includes all potential colors used by this class, regardless of state change
' or internal control settings.  The list is updated by calling the UpdateColorList function.
' (Note also that this list does not include variants, e.g. "BorderColor" vs "BorderColor_Hovered".  Variant values are
'  automatically calculated by the color management class, and they are retrieved by passing boolean modifiers to that
'  class, rather than treating every imaginable variant as a separate constant.)
Private Enum PDCV_COLOR_LIST
    [_First] = 0
    PDCV_Background = 0
    PDCV_Border = 1
    [_Last] = 1
    [_Count] = 2
End Enum

'Color retrieval and storage is handled by a dedicated class; this allows us to optimize theme interactions,
' without worrying about the details locally.
Private m_Colors As pdThemeColors

Public Function GetControlType() As PD_ControlType
    GetControlType = pdct_ColorVariants
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
    Color = m_ColorList(0)
End Property

Public Property Let Color(ByVal newColor As Long)
    
    m_ColorList(0) = newColor
    ColorManagement.ApplyDisplayColorManagement_SingleColor m_ColorList(0), m_ColorDisplay(0)
    
    MakeNewTooltip CV_Primary
    
    'Recalculate all color variants, then redraw the control
    CalculateVariantColors
    RedrawBackBuffer
    
    RaiseEvent ColorChanged(m_ColorList(0), False)
    PropertyChanged "Color"
    
End Property

Public Property Get WheelShape() As COLOR_WHEEL_SHAPE
    WheelShape = m_ControlShape
End Property

Public Property Let WheelShape(ByVal newShape As COLOR_WHEEL_SHAPE)
    If (m_ControlShape <> newShape) Then
        m_ControlShape = newShape
        UpdateControlLayout
        PropertyChanged "WheelShape"
    End If
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
    
        'See if the mouse cursor is inside a box
        m_MouseInsideRegion = GetRegionFromPoint(x, y)
        
        If (m_MouseInsideRegion >= 0) Then
        
            'If the primary color box is clicked, raise a full color selection dialog
            If (m_MouseInsideRegion = 0) Then
                DisplayColorSelection
            Else
                m_ColorList(0) = m_ColorList(m_MouseInsideRegion)
                ColorManagement.ApplyDisplayColorManagement_SingleColor m_ColorList(0), m_ColorDisplay(0)
            End If
            
            'Recalculate all color variants to match the new color (if any) and redraw the control
            MakeNewTooltip m_MouseInsideRegion
            CalculateVariantColors
            RedrawBackBuffer
            
            'Raise an event to match
            RaiseEvent ColorChanged(m_ColorList(0), True)
        
        End If
        
    End If
    
End Sub

Private Sub ucSupport_MouseLeave(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    ucSupport.RequestCursor IDC_DEFAULT
    m_MouseInsideRegion = -1
    RedrawBackBuffer
End Sub

Private Sub ucSupport_MouseMoveCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal timeStamp As Long)
    
    'Calculate a new hovered box ID, if any
    Dim oldMouseIndex As Long
    oldMouseIndex = m_MouseInsideRegion
    m_MouseInsideRegion = GetRegionFromPoint(x, y)
    
    'Modify the cursor to match
    If (m_MouseInsideRegion >= 0) Then ucSupport.RequestCursor IDC_HAND Else ucSupport.RequestCursor IDC_DEFAULT
    
    'If the box ID has changed, update the tooltip and redraw the control to match
    If (m_MouseInsideRegion <> oldMouseIndex) Then
        RedrawBackBuffer True
        MakeNewTooltip m_MouseInsideRegion
    End If
    
End Sub

'Given an (x, y) coordinate pair from the mouse, return the index of the containing rect (if any).
' Returns -1 if the point lies outside all rects.
Private Function GetRegionFromPoint(ByVal x As Single, ByVal y As Single) As Long

    GetRegionFromPoint = -1
    
    Dim i As Long
    For i = 0 To NUM_OF_VARIANTS - 1
        
        If m_ColorRegions(i).IsPointInsidePathF(x, y) Then
            GetRegionFromPoint = i
            Exit Function
        End If
        
    Next i

End Function

Private Sub ucSupport_RepaintRequired(ByVal updateLayoutToo As Boolean)
    If updateLayoutToo Then UpdateControlLayout Else RedrawBackBuffer
End Sub

Private Sub UserControl_Initialize()
    
    'Initialize a user control support class
    Set ucSupport = New pdUCSupport
    ucSupport.RegisterControl UserControl.hWnd, True
    ucSupport.RequestExtraFunctionality True
    ucSupport.SubclassCustomMessage WM_PD_COLOR_MANAGEMENT_CHANGE, True
    
    m_MouseInsideRegion = -1
    
    'Prep the color manager and load default colors
    Set m_Colors = New pdThemeColors
    Dim colorCount As PDCV_COLOR_LIST: colorCount = [_Count]
    m_Colors.InitializeColorList "PDColorVariants", colorCount
    If Not PDMain.IsProgramRunning() Then UpdateColorList
    
    'Prep the various color variant lists
    ReDim m_ColorList(0 To NUM_OF_VARIANTS - 1) As Long
    ReDim m_ColorDisplay(0 To NUM_OF_VARIANTS - 1) As Long
    ReDim m_ColorRegions(0 To NUM_OF_VARIANTS - 1) As pd2DPath
    
    Dim i As Long
    For i = 0 To NUM_OF_VARIANTS - 1
        Set m_ColorRegions(i) = New pd2DPath
    Next i
    
    CalculateVariantColors
    
    'Draw the control at least once
    UpdateControlLayout
    
End Sub

Private Sub UserControl_InitProperties()
    Color = RGB(50, 200, 255)
    WheelShape = CWS_Circular
End Sub

'At run-time, painting is handled by the support class.  In the IDE, however, we must rely on VB's internal paint event.
Private Sub UserControl_Paint()
    ucSupport.RequestIDERepaint UserControl.hDC
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Me.Color = PropBag.ReadProperty("Color", RGB(50, 200, 255))
    Me.WheelShape = PropBag.ReadProperty("WheelShape", CWS_Circular)
End Sub

Private Sub UserControl_Resize()
    If (Not PDMain.IsProgramRunning()) Then ucSupport.RequestRepaint True
End Sub
    
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "Color", Me.Color, RGB(50, 200, 255)
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
    If ShowColorDialog(newColor, oldColor, Nothing) Then
        m_ColorList(0) = newColor
    Else
        m_ColorList(0) = oldColor
    End If
    
    ColorManagement.ApplyDisplayColorManagement_SingleColor m_ColorList(0), m_ColorDisplay(0)
    
End Sub

Private Sub NotifyColorManagementChange()
    ColorManagement.ApplyDisplayColorManagement_SingleColor m_ColorList(CV_Primary), m_ColorDisplay(CV_Primary)
    CalculateVariantColors
    RedrawBackBuffer
End Sub

'Call this to recreate all buffers against a changed control size.
Private Sub UpdateControlLayout()
    
    'Retrieve DPI-aware control dimensions from the support class
    Dim bWidth As Long, bHeight As Long
    bWidth = ucSupport.GetBackBufferWidth
    bHeight = ucSupport.GetBackBufferHeight
    
    If PDMain.IsProgramRunning() Then
        
        'Re-calculate all control subregions.  This is a little confusing (okay, a LOT confusing), but basically we want to
        ' create an evenly spaced border around the central color rect, with subdivided regions that provide some dynamic
        ' color variants for the user to choose from.
        Dim i As Long
        For i = 0 To NUM_OF_VARIANTS - 1
            m_ColorRegions(i).ResetPath
        Next i
        
        'Leave a little room around the control, so we can draw chunky borders around hovered sub-regions.
        Dim ucLeft As Long, ucTop As Long, ucBottom As Long, ucRight As Long
        ucLeft = 1
        ucTop = 1
        ucBottom = bHeight - 2
        ucRight = bWidth - 2
        
        'How we actually create the regions varies depending on the current control orientation.
        If (m_ControlShape = CWS_Circular) Then
            CreateSubregions_Circular ucLeft, ucTop, ucBottom, ucRight
        Else
            CreateSubregions_Rectangular ucLeft, ucTop, ucBottom, ucRight
        End If
        
    End If
    
    'With the backbuffer and rects successfully constructed, we can finally redraw the control
    RedrawBackBuffer
    
End Sub

'Create a rectangular grid-based variant control
Private Sub CreateSubregions_Rectangular(ByVal ucLeft As Long, ByVal ucTop As Long, ByVal ucBottom As Long, ByVal ucRight As Long)
    
    'First, make sure our border size is DPI-aware
    Dim dpiAwareBorderSize As Long
    dpiAwareBorderSize = FixDPI(VARIANT_BOX_SIZE)
    
    'For this control layout, we are going to use a temporary rect collection to define the position of all color variants.
    ' This simplifies things a bit, and when we're done, we'll simply copy all rects into the central pd2DPath array.
    Dim colorRects() As RectF
    ReDim colorRects(0 To NUM_OF_VARIANTS - 1) As RectF
    
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
    rgbHeight = colorRects(CV_Primary).Height / 3!
    
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
    
    'With the color rects successfully constructed, we can now add them to our central path collection
    For i = CV_Primary To NUM_OF_VARIANTS - 1
        m_ColorRegions(i).AddRectangle_RectF colorRects(i)
    Next i
    
End Sub

Private Sub CreateSubregions_Circular(ByVal ucLeft As Long, ByVal ucTop As Long, ByVal ucBottom As Long, ByVal ucRight As Long)

    'First, make sure our border size is DPI-aware
    Dim dpiAwareBorderSize As Long
    dpiAwareBorderSize = Interface.FixDPI(VARIANT_BOX_SIZE)
    
    'Constructing circular sub-regions actually involves less code than rectangular ones, because they're spaced perfectly evenly,
    ' so we can easily construct them in a loop.
    
    'Start by calculating basic arc and circle values
    Dim minDimension As Single
    If (ucRight - ucLeft < ucBottom - ucTop) Then
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
    
    'Failsafe check
    If (innerRadius <= 4) Then Exit Sub
    
    'The primary circle is the only subregion that receives a special construction method.
    m_ColorRegions(CV_Primary).AddCircle centerX, centerY, innerRadius
    
    'All subregions are added uniformly, in a loop
    Dim startAngle As Single, sweepAngle As Single
    startAngle = 180
    sweepAngle = 30
    
    Dim x1 As Double, x2 As Double, x3 As Double, x4 As Double, y1 As Double, y2 As Double, y3 As Double, y4 As Double
    
    Dim i As Long
    For i = 1 To NUM_OF_VARIANTS - 1
    
        'Calculate (x, y) coordinates for the four corners of this subregion.  We use these as the endpoints for the radial lines
        ' marking either side of the "slice".
        PDMath.ConvertPolarToCartesian PDMath.DegreesToRadians(startAngle), innerRadius, x1, y1, centerX, centerY
        PDMath.ConvertPolarToCartesian PDMath.DegreesToRadians(startAngle), outerRadius, x2, y2, centerX, centerY
        PDMath.ConvertPolarToCartesian PDMath.DegreesToRadians(startAngle + sweepAngle), innerRadius, x3, y3, centerX, centerY
        PDMath.ConvertPolarToCartesian PDMath.DegreesToRadians(startAngle + sweepAngle), outerRadius, x4, y4, centerX, centerY
        
        'Add the two divider lines to the current path object, and place connecting arcs between them
        m_ColorRegions(i).AddLine x1, y1, x2, y2
        m_ColorRegions(i).AddArcCircular centerX, centerY, outerRadius, startAngle, sweepAngle
        m_ColorRegions(i).AddLine x4, y4, x3, y3
        m_ColorRegions(i).AddArcCircular centerX, centerY, innerRadius, startAngle + sweepAngle, -sweepAngle
        m_ColorRegions(i).CloseCurrentFigure
        
        'Offset the startAngle for the next slice
        startAngle = startAngle + sweepAngle
        
    Next i
    
End Sub

'Any time the primary color changes (for whatever reason, external or internal), new variant colors must be calculated.
' Call this function to auto-calculate them, but try to do it only when necessary, as there's a lot of math involved.
Private Sub CalculateVariantColors()
    
    If (Not PDMain.IsProgramRunning()) Then Exit Sub
    
    'The primary color serves as the base color for all subsequent calculations.  Retrieve its RGB and HSV quads now.
    Dim rPrimary As Long, gPrimary As Long, bPrimary As Long, hPrimary As Double, sPrimary As Double, vPrimary As Double
    rPrimary = Colors.ExtractRed(m_ColorList(CV_Primary))
    gPrimary = Colors.ExtractGreen(m_ColorList(CV_Primary))
    bPrimary = Colors.ExtractBlue(m_ColorList(CV_Primary))
    Colors.RGBtoHSV rPrimary, gPrimary, bPrimary, hPrimary, sPrimary, vPrimary
    
    'We now need to calculate new RGB values.  How we do this varies by variant, obviously!
    Dim rNew As Long, gNew As Long, bNew As Long, hNew As Double, sNew As Double, vNew As Double
    Dim rFloat As Double, gFloat As Double, bFloat As Double
    Dim grayNew As Long
    
    'Also, during testing I'm experimenting with different increment amounts for HSV and RGB adjustments
    Dim rgbChange As Long, svChange As Double, hChange As Double
    rgbChange = 20
    svChange = 0.08
    hChange = 0.03
    
    Const ONE_DIV_255 As Double = 1# / 255#
    
    Dim i As COLOR_VARIANTS
    For i = CV_HueUp To CV_RedDown
        
        rNew = rPrimary: gNew = gPrimary: bNew = bPrimary
        rFloat = rNew * ONE_DIV_255: gFloat = gNew * ONE_DIV_255: bFloat = bNew * ONE_DIV_255
        hNew = hPrimary: sNew = sPrimary: vNew = vPrimary
        
        If (i = CV_HueUp) Then
            hNew = hNew + hChange
            If (hNew > 1#) Then hNew = hNew - 1#
            Colors.HSVtoRGB hNew, sNew, vNew, rNew, gNew, bNew
            
        ElseIf (i = CV_SaturationUp) Then
                
            'Use a fake saturation calculation
            grayNew = Colors.GetHQLuminance(rNew, gNew, bNew)
            rNew = rNew + (rNew - grayNew) * 0.4
            gNew = gNew + (gNew - grayNew) * 0.4
            bNew = bNew + (bNew - grayNew) * 0.4
            rNew = PDMath.ClampL(rNew, 0, 255)
            gNew = PDMath.ClampL(gNew, 0, 255)
            bNew = PDMath.ClampL(bNew, 0, 255)
            
        ElseIf (i = CV_ValueUp) Then
                
            'Use a fake value calculation
            rNew = PDMath.ClampL(rNew + rgbChange, 0, 255)
            gNew = PDMath.ClampL(gNew + rgbChange, 0, 255)
            bNew = PDMath.ClampL(bNew + rgbChange, 0, 255)
            
        ElseIf (i = CV_RedUp) Then
            rNew = PDMath.ClampL(rNew + rgbChange, 0, 255)
                
        ElseIf (i = CV_GreenUp) Then
            gNew = PDMath.ClampL(gNew + rgbChange, 0, 255)
                
        ElseIf (i = CV_BlueUp) Then
            bNew = PDMath.ClampL(bNew + rgbChange, 0, 255)
                
        ElseIf (i = CV_ValueDown) Then
                
            'Use a fake value calculation
            rNew = PDMath.ClampL(rNew - rgbChange, 0, 255)
            gNew = PDMath.ClampL(gNew - rgbChange, 0, 255)
            bNew = PDMath.ClampL(bNew - rgbChange, 0, 255)
            
        ElseIf (i = CV_SaturationDown) Then
                
            'Use a fake saturation calculation
            grayNew = Colors.GetHQLuminance(rNew, gNew, bNew)
            rNew = rNew + (grayNew - rNew) * 0.3
            gNew = gNew + (grayNew - gNew) * 0.3
            bNew = bNew + (grayNew - bNew) * 0.3
            rNew = PDMath.ClampL(rNew, 0, 255)
            gNew = PDMath.ClampL(gNew, 0, 255)
            bNew = PDMath.ClampL(bNew, 0, 255)
            
        ElseIf (i = CV_HueDown) Then
            hNew = hNew - hChange
            If (hNew < 0#) Then hNew = 1# + hNew
            Colors.HSVtoRGB hNew, sNew, vNew, rNew, gNew, bNew
        
        ElseIf (i = CV_BlueDown) Then
            bNew = PDMath.ClampL(bNew - rgbChange, 0, 255)
                
        ElseIf (i = CV_GreenDown) Then
            gNew = PDMath.ClampL(gNew - rgbChange, 0, 255)
                
        ElseIf (i = CV_RedDown) Then
            rNew = PDMath.ClampL(rNew - rgbChange, 0, 255)
        
        End If
        
        'Cache the new RGB values
        m_ColorList(i) = RGB(rNew, gNew, bNew)
        
        'If color management is active, apply it now
        ColorManagement.ApplyDisplayColorManagement_SingleColor m_ColorList(i), m_ColorDisplay(i)
    
    Next i
    
    'After recreating color values, the control must be redrawn, but we leave this to our caller to handle
    
End Sub

'Redraw the UC.  Note that some UI elements must be created prior to calling this function (e.g. the color wheel).
Private Sub RedrawBackBuffer(Optional ByVal redrawImmediately As Boolean = False)
    
    'Request the back buffer DC, and ask the support module to erase any existing rendering for us.
    Dim bufferDC As Long
    bufferDC = ucSupport.GetBackBufferDC(True, m_Colors.RetrieveColor(PDCV_Background, Me.Enabled))
    If (bufferDC = 0) Then Exit Sub
    
    Dim bWidth As Long, bHeight As Long
    bWidth = ucSupport.GetBackBufferWidth
    bHeight = ucSupport.GetBackBufferHeight
    
    If PDMain.IsProgramRunning() Then
    
        Dim borderColor As Long
        borderColor = m_Colors.RetrieveColor(PDCV_Border, Me.Enabled, False, False)
        
        'Prep various painting objects
        Dim cSurface As pd2DSurface, cBrush As pd2DBrush, cPen As pd2DPen
        Dim cPenUIBase As pd2DPen, cPenUITop As pd2DPen
        Drawing2D.QuickCreateSurfaceFromDC cSurface, bufferDC, True
        
        'We can reuse a single border pen for all sub-paths
        Drawing2D.QuickCreateSolidPen cPen, 1#, borderColor, 100#, P2_LJ_Miter
        
        'Draw each subregion in turn, filling it first, then tracing its borders.
        Dim i As Long
        For i = CV_Primary To CV_RedDown
            Drawing2D.QuickCreateSolidBrush cBrush, m_ColorDisplay(i)
            PD2D.FillPath cSurface, cBrush, m_ColorRegions(i)
            PD2D.DrawPath cSurface, cPen, m_ColorRegions(i)
        Next i
        
        'Draw a special outline around the central primary color, to help it stand out more.  (But only do this if
        ' the central primary color is UNSELECTED; if it's selected, we'll paint it in the accent color momentarily.)
        If (m_MouseInsideRegion <> CV_Primary) Then
            Drawing.BorrowCachedUIPens cPenUIBase, cPenUITop
            PD2D.DrawPath cSurface, cPenUIBase, m_ColorRegions(CV_Primary)
            PD2D.DrawPath cSurface, cPenUITop, m_ColorRegions(CV_Primary)
        End If
        
        'If a subregion is currently hovered, trace it with a highlight outline.
        If (m_MouseInsideRegion >= 0) Then
            Drawing.BorrowCachedUIPens cPenUIBase, cPenUITop, True
            PD2D.DrawPath cSurface, cPenUIBase, m_ColorRegions(m_MouseInsideRegion)
            PD2D.DrawPath cSurface, cPenUITop, m_ColorRegions(m_MouseInsideRegion)
        End If
        
        Set cSurface = Nothing: Set cBrush = Nothing: Set cPen = Nothing
        
    End If
    
    'Paint the final result to the screen, as relevant
    ucSupport.RequestRepaint redrawImmediately
    
End Sub

'When the currently hovered color variant changes, we want to assign a new tooltip to the control
Private Sub MakeNewTooltip(ByVal activeIndex As COLOR_VARIANTS)
    
    'Failsafe for compile-time errors when properties are written
    If (Not PDMain.IsProgramRunning()) Then Exit Sub
    
    Dim toolString As String, hexString As String, rgbString As String
    
    If (activeIndex >= 0) Then
        hexString = "#" & UCase$(Colors.GetHexStringFromRGB(m_ColorList(activeIndex)))
        rgbString = Colors.ExtractRed(m_ColorList(activeIndex)) & ", " & Colors.ExtractGreen(m_ColorList(activeIndex)) & ", " & Colors.ExtractBlue(m_ColorList(activeIndex))
        toolString = hexString & vbCrLf & rgbString
        If (activeIndex = CV_Primary) Then toolString = toolString & vbCrLf & g_Language.TranslateMessage("Click to enter a full color selection screen.")
    End If
    
    Select Case activeIndex
        
        Case CV_Primary
            Me.AssignTooltip toolString, "Current color", True
        
        Case CV_HueUp
            Me.AssignTooltip toolString, "Rotate hue clockwise", True
                
        Case CV_SaturationUp
            Me.AssignTooltip toolString, "Increase saturation", True
            
        Case CV_ValueUp
            Me.AssignTooltip toolString, "Increase luminance", True
            
        Case CV_RedUp
            Me.AssignTooltip toolString, "Increase red", True
            
        Case CV_GreenUp
            Me.AssignTooltip toolString, "Increase green", True
            
        Case CV_BlueUp
            Me.AssignTooltip toolString, "Increase blue", True
            
        Case CV_ValueDown
            Me.AssignTooltip toolString, "Decrease luminance", True
            
        Case CV_SaturationDown
            Me.AssignTooltip toolString, "Decrease saturation", True
            
        Case CV_HueDown
            Me.AssignTooltip toolString, "Rotate hue counterclockwise", True
            
        Case CV_BlueDown
            Me.AssignTooltip toolString, "Decrease blue", True
            
        Case CV_GreenDown
            Me.AssignTooltip toolString, "Decrease green", True
            
        Case CV_RedDown
            Me.AssignTooltip toolString, "Decrease red", True
        
        Case Else
            Me.AssignTooltip vbNullString, , False
                
    End Select
    
End Sub

'Before this control does any painting, we need to retrieve relevant colors from PD's primary theming class.  Note that this
' step must also be called if/when PD's visual theme settings change.
Private Sub UpdateColorList()
    m_Colors.LoadThemeColor PDCV_Background, "Background", IDE_WHITE
    m_Colors.LoadThemeColor PDCV_Border, "Border", IDE_BLACK
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

'Due to complex interactions between user controls and PD's translation engine, tooltips require this dedicated function.
' (IMPORTANT NOTE: the tooltip class will handle translations automatically.  Always pass the original English text!)
Public Sub AssignTooltip(ByRef newTooltip As String, Optional ByRef newTooltipTitle As String = vbNullString, Optional ByVal raiseTipsImmediately As Boolean = False)
    ucSupport.AssignTooltip UserControl.ContainerHwnd, newTooltip, newTooltipTitle, raiseTipsImmediately
End Sub
