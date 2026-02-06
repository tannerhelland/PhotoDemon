Attribute VB_Name = "Tools_Gradient"
'***************************************************************************
'PhotoDemon On-Canvas Gradient Tool Manager
'Copyright 2018-2026 by Tanner Helland
'Created: 31/December/18
'Last updated: 26/February/19
'Last update: add progress bar updates for long-running gradient renders
'
'This module interfaces between the gradient tool UI and pd2DGradient backend.  Look in the relevant
' tool panel form for more details on how the UI relays relevant fill data here.
'
'Most gradient shapes are rendered by internal, pure-VB6 renderers.  There are multiple reasons for this;
' 3rd-party libraries (GDI+, Cairo) don't support many of the gradient shapes, and even when they do,
' our internal renderers are often faster.  (This is especially true of Cairo, which is unfortunately
' slow across almost all gradient rendering modes.)
'
'The class is currently designed to use lookup tables for calculating actual gradient colors.
' This provides a ton of flexibility for future enhancements; for example, if in the future we want to
' implement something like blending colors in L*a*b* space, I simply need to change the lookup table
' generation code - the renderers themselves need no adjustment, as they just blindly pick colors
' from the lut.
'
'At present, colors are mixed in sRGB space, which is not really ideal.  Future enhancements will
' likely start here.  Adding jitter dither to the gradients themselves would also be a nice addition,
' although that *would* require some changes inside the renderers themselves.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Development-time only; remove in production
Private Const PROFILE_GRADIENT_PERF As Boolean = False

'Depending on gradient and/or system settings, we may switch between different gradient renderers.
' Many - but not all - gradient functions are implemented against multiple backends, so you'll need
' to look at the ChooseRenderer function to figure out which backend will be used on any given call.
Private Enum PD_GradientRenderer
    gr_Internal = 0
    gr_GDIPlus = 1
End Enum

#If False Then
    Private Const gr_Internal = 0, gr_GDIPlus = 1
#End If

Public Enum PD_GradientAttributes
    GA_Opacity = 0
    GA_BlendMode = 1
    GA_AlphaMode = 2
    GA_Antialiasing = 3
    GA_Repeat = 4
    GA_Shape = 5
End Enum

#If False Then
    Private Const GA_Opacity = 0, GA_BlendMode = 1, GA_AlphaMode = 2, GA_Antialiasing = 3, GA_Repeat = 4, GA_Shape = 5
#End If

Public Enum PD_GradientRepeat
    gr_None = 0
    gr_Clamp = 1
    gr_Wrap = 2
    gr_Reflect = 3
End Enum

#If False Then
    Private Const gr_None = 0, gr_Clamp = 1, gr_Wrap = 2, gr_Reflect = 3
#End If

Public Enum PD_GradientShape
    gs_Linear = 0
    gs_Reflection = 1
    gs_Radial = 2
    gs_Spherical = 3
    gs_Square = 4
    gs_Diamond = 5
    gs_Conical = 6
    gs_Spiral = 7
End Enum

#If False Then
    Private Const gs_Linear = 0, gs_Reflection = 1, gs_Radial = 2, gs_Spherical = 3, gs_Square = 4, gs_Diamond = 5, gs_Conical = 6, gs_Spiral = 7
#End If

'Gradient attributes are stored in these variables
Private m_GradientOpacity As Single
Private m_GradientBlendmode As PD_BlendMode
Private m_GradientAlphamode As PD_AlphaMode
Private m_GradientRepeat As PD_GradientRepeat
Private m_GradientShape As PD_GradientShape

'Uninitialized mouse points (i.e. if the user hasn't clicked the mouse yet) are initialized to an
' "impossible" UI value.
Private Const MOUSE_OOB As Single = -9.99999E+14!

'Current mouse/pen input values.  These are blindly relayed to us by the canvas, and it's up to us to perform any
' special tracking calculations.
Private m_MouseDown As Boolean
Private m_MouseX As Single, m_MouseY As Single

'Start and current x/y coordinates, *in image coordinates* (per convention).  Do not attempt to access
' the array without first checking m_PointsInitialized.
Private m_PointsInitialized As Boolean
Private m_Points() As PointFloat

'A persistent gradient object is used to perform the actual gradient rendering
Private m_GradientGdip As pd2DGradient

'Other gradient parameters, as relevant
Private m_RadialOffset As Single

'To avoid the case where the user releases a shift/ctrl/alt modifier near the same time as
' _MouseUp, we cache key state during MouseMove events and ensure that the final render always
' uses the key state of the last MouseMove event - *not* the keystate at _MouseUp, which may
' be different.
Private m_KeyModifiers As ShiftConstants

'When our internal gradient renderer is uses, we don't want to manually interpolate gradient
' values for every point on-the-fly, as that's crazy slow.  Instead, we pre-generate a gradient
' lookup table.  The size of this table depends on a number of factors, including gradient-specific
' settings like the number of colors in use (e.g. 2-color gradients don't benefit from extremely
' large lookups).  IMPORTANT NOTE: the lookup table resolution must *always* be a power of 2.
' This allows us to use && instead of % on the inner rendering loop, at a large boost to performance.
' Changing it to some value that is NOT a power of 2 will break the renderer.
'
'Note that we also cache the "last" color in the gradient; this accelerates the default "clamped"
' edge mode, as we can bypass the lookup table entirely for points outside the gradient's boundary.
Private Const MAX_LOOKUP_RESOLUTION As Long = 8192
Private m_GradLookup() As Long, m_LookupResolution As Long, m_LastGradColor As Long, m_LastGradColorQuad As RGBQuad

'To improve canvas responsiveness, this module can render specialized "fast" previews during
' UI interactions, then silently switch to full "accurate" rendering on _MouseUp.  When fast previews
' are active (currently determined by the global viewport performance user preference), this temporary
' DIB is used to cache intermediate gradient results.
Private m_PreviewDIB As pdDIB

'Set to TRUE when a gradient is actively being rendered; we use this to skip rendering intermediate
' steps during _MouseMove events.
Private m_GradientRendering As Boolean

'Universal gradient settings
Public Function GetGradientAlphaMode() As PD_AlphaMode
    GetGradientAlphaMode = m_GradientAlphamode
End Function

Public Function GetGradientBlendMode() As PD_BlendMode
    GetGradientBlendMode = m_GradientBlendmode
End Function

Public Function GetGradientOpacity() As Single
    GetGradientOpacity = m_GradientOpacity
End Function

Public Function GetGradientRadialOffset() As Single
    GetGradientRadialOffset = m_RadialOffset
End Function

Public Function GetGradientRepeat() As PD_GradientRepeat
    GetGradientRepeat = m_GradientRepeat
End Function

Public Function GetGradientShape() As PD_GradientShape
    GetGradientShape = m_GradientShape
End Function

'Property set functions.  Note that not all brush properties are used by all styles.
' (e.g. "brush hardness" is not used by "pencil" style brushes, etc)
Public Sub SetGradientAlphaMode(Optional ByVal newAlphaMode As PD_AlphaMode = AM_Normal)
    If (newAlphaMode <> m_GradientAlphamode) Then m_GradientAlphamode = newAlphaMode
End Sub

Public Sub SetGradientBlendMode(Optional ByVal newBlendMode As PD_BlendMode = BM_Normal)
    If (newBlendMode <> m_GradientBlendmode) Then m_GradientBlendmode = newBlendMode
End Sub

Public Sub SetGradientOpacity(ByVal newOpacity As Single)
    If (newOpacity <> m_GradientOpacity) Then m_GradientOpacity = newOpacity
End Sub

Public Sub SetGradientRadialOffset(ByVal newOffset As Single)
    If (newOffset <> m_RadialOffset) Then m_RadialOffset = newOffset
End Sub

Public Sub SetGradientRepeat(ByVal newRepeat As PD_GradientRepeat)
    If (newRepeat <> m_GradientRepeat) Then m_GradientRepeat = newRepeat
End Sub

Public Sub SetGradientShape(ByVal newShape As PD_GradientShape)
    If (newShape <> m_GradientShape) Then m_GradientShape = newShape
End Sub

Public Function GetGradientProperty(ByVal bProperty As PD_GradientAttributes) As Variant
    
    Select Case bProperty
        Case GA_AlphaMode
            GetGradientProperty = GetGradientAlphaMode()
        Case GA_BlendMode
            GetGradientProperty = GetGradientBlendMode()
        Case GA_Opacity
            GetGradientProperty = GetGradientOpacity()
        Case GA_Repeat
            GetGradientProperty = GetGradientRepeat()
        Case GA_Shape
            GetGradientProperty = GetGradientShape()
    End Select
    
End Function

'Notify the gradient engine of the current mouse position.  Coordinates should always be in *image* coordinate space,
' not screen space.  (Translation between spaces will be handled internally.)
Public Sub NotifyToolXY(ByVal mouseButtonDown As Boolean, ByVal Shift As ShiftConstants, ByVal srcX As Single, ByVal srcY As Single, ByVal mouseTimeStamp As Long, ByRef srcCanvas As pdCanvas)
    
    'Failsafe check: no images loaded
    If (Not PDImages.IsImageActive()) Then Exit Sub
    
    'If a render is already in progress, we typically want to exit immediately
    ' (as the previous render needs to finish).  If we don't do this, mouse events raised
    ' while the final gradient is rendering will trigger this sub again, causing reentrancy
    ' nightmares.
    If m_GradientRendering Then Exit Sub
    m_GradientRendering = True
    
    m_MouseX = srcX
    m_MouseY = srcY
    
    Dim isFirstStroke As Boolean, isLastStroke As Boolean
    isFirstStroke = (Not m_MouseDown) And mouseButtonDown
    isLastStroke = m_MouseDown And (Not mouseButtonDown)
    
    'Different backends are used for different gradient settings (e.g. linear gradients can be rendered
    ' very nicely by GDI+, but conical gradients must be manually rendered).
    Dim curBackend As PD_GradientRenderer
    curBackend = GetBestRenderer()
    
    'On first stroke, initialize the point array and store the base point coordinates
    If isFirstStroke Then
    
        InitializePoints
        m_Points(0).x = srcX
        m_Points(0).y = srcY
        
        'Make sure the current scratch layer is properly initialized
        Tools.InitializeToolsDependentOnImage
        PDImages.GetActiveImage.ScratchLayer.SetLayerOpacity m_GradientOpacity
        PDImages.GetActiveImage.ScratchLayer.SetLayerBlendMode m_GradientBlendmode
        PDImages.GetActiveImage.ScratchLayer.SetLayerAlphaMode m_GradientAlphamode
        
        If (curBackend = gr_GDIPlus) Then
        
            Set m_GradientGdip = New pd2DGradient
            m_GradientGdip.CreateGradientFromString toolpanel_Gradient.grdPrimary.Gradient()
            m_GradientGdip.SetGradientShape P2_GS_Linear
            
            'Note: GDI+ does not support wrap modes the same way that our internal renderers do
            Select Case m_GradientRepeat
                Case gr_None
                    'Clamp mode is not supported by GDI+, so we lie and set a functional mode
                    ' and simply overwrite the results later
                    m_GradientGdip.SetGradientWrapMode P2_WM_TileFlipXY
                Case gr_Wrap
                    m_GradientGdip.SetGradientWrapMode P2_WM_Tile
                Case gr_Reflect
                    m_GradientGdip.SetGradientWrapMode P2_WM_TileFlipXY
            End Select
            
        ElseIf (curBackend = gr_Internal) Then
            
            'Our freestanding gradient class can directly produce a lookup table for us.
            Set m_GradientGdip = New pd2DGradient
            m_GradientGdip.CreateGradientFromString toolpanel_Gradient.grdPrimary.Gradient()
            m_GradientGdip.SetGradientShape P2_GS_Linear
            
            If (m_GradientGdip.GetNumOfNodes < 3) Then
                m_LookupResolution = 256
            Else
                m_LookupResolution = m_GradientGdip.GetNumOfNodes * 256
                m_LookupResolution = PDMath.NearestPowerOfTwo(m_LookupResolution)
                If (m_LookupResolution > MAX_LOOKUP_RESOLUTION) Then m_LookupResolution = MAX_LOOKUP_RESOLUTION
            End If
            
            m_GradientGdip.GetLookupTable m_GradLookup, m_LookupResolution
            m_LastGradColor = m_GradientGdip.GetLastColor()
            m_LastGradColorQuad = m_GradientGdip.GetLastColorQuad()
            
        End If
        
    End If
    
    'On any other stroke, update the 2nd set of mouse coordinates
    If m_PointsInitialized Then
        m_Points(1).x = srcX
        m_Points(1).y = srcY
    End If
    
    'To ensure the final _MouseUp event keystate matches the keystate of any previous events,
    ' we cache key state during _MouseMove and reuse the last-known state on _MouseUp.
    Dim ctrlKeyDown As Boolean
    If (Not isLastStroke) Then m_KeyModifiers = Shift
    ctrlKeyDown = ((m_KeyModifiers And vbCtrlMask) = vbCtrlMask)
    
    'If the CTRL key has been pressed, we want to lock the angle between the first point and second point
    ' to an increment of 15 degrees.
    If ctrlKeyDown And m_PointsInitialized Then
    
        Dim curAngle As Double, curDistance As Double
        curDistance = PDMath.DistanceTwoPoints(m_Points(0).x, m_Points(0).y, m_Points(1).x, m_Points(1).y)
        curAngle = PDMath.Atan2(m_Points(1).y - m_Points(0).y, m_Points(1).x - m_Points(0).x)
        
        'Lock the angle to the nearest multiple of 15 degrees.  (I got this idea from GIMP; 15 degrees is
        ' nice because it covers all major rotational values: 0, 15, 30, 45, 60, 75, 90.)
        Dim curAngleDegrees As Double, curAngleModulo As Double
        curAngleDegrees = PDMath.RadiansToDegrees(curAngle)
        curAngleModulo = PDMath.Modulo(curAngleDegrees, 15#)
        
        curAngleDegrees = curAngleDegrees - curAngleModulo
        If (curAngleModulo >= 7.5) Then curAngleDegrees = curAngleDegrees + 15#
        curAngle = PDMath.DegreesToRadians(curAngleDegrees)
        
        'Remap the point in question to the newly calculated angle
        Dim newX As Single, newY As Single
        PDMath.RotatePointAroundPoint m_Points(0).x + curDistance, m_Points(0).y, m_Points(0).x, m_Points(0).y, curAngle, newX, newY
        m_Points(1).x = newX
        m_Points(1).y = newY
        
    End If
    
    'Notify the scratch layer of our updates
    If mouseButtonDown Or isLastStroke Then
        
        PDImages.GetActiveImage.ScratchLayer.GetLayerDIB.ResetDIB 0
        
        If (g_ViewportPerformance >= PD_PERF_BALANCED) And (Not isLastStroke) Then
            PreviewRenderer srcCanvas, m_Points(0), m_Points(1), curBackend
        Else
            
            If PROFILE_GRADIENT_PERF Then
                Dim gradStartTime As Currency
                VBHacks.GetHighResTime gradStartTime
            End If
            
            If (curBackend = gr_GDIPlus) Then
                GdipRenderer m_Points(0), m_Points(1), PDImages.GetActiveImage.ScratchLayer.GetLayerDIB
            ElseIf (curBackend = gr_Internal) Then
                InternalRenderer m_Points(0), m_Points(1), PDImages.GetActiveImage.ScratchLayer.GetLayerDIB
            End If
            
            If PROFILE_GRADIENT_PERF Then PDDebug.LogAction "Gradient rendered by " & GetNameOfRenderer(curBackend) & " in " & VBHacks.GetTimeDiffNowAsString(gradStartTime)
            
        End If
        
        'Notify the target layer of the changes
        PDImages.GetActiveImage.ScratchLayer.NotifyOfDestructiveChanges
        
    End If
    
    'With all drawing tasks complete, update all old state values to match the new state values.
    m_MouseDown = mouseButtonDown
    
    'On last stroke, release the gradient UI elements (as the mouse has been released)
    If isLastStroke Then m_PointsInitialized = False
    
    'Notify the viewport of the need for a redraw
    Dim tmpViewportParams As PD_ViewportParams
    tmpViewportParams = Viewport.GetDefaultParamObject()
    tmpViewportParams.renderScratchLayerIndex = PDImages.GetActiveImage.GetActiveLayerIndex()
    
    'If fast previews are active, we want to inject our own local scratch layer instead of using
    ' the standard (full-image-sized) one.
    If (g_ViewportPerformance >= PD_PERF_BALANCED) And (Not isLastStroke) Then tmpViewportParams.ptrToAlternateScratch = ObjPtr(m_PreviewDIB)
    If mouseButtonDown Then Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), srcCanvas, VarPtr(tmpViewportParams)
    
    'We can mark rendering as "complete" now
    m_GradientRendering = False
    
End Sub

'Return the best renderer for the current gradient job; this varies according to both gradient and system settings.
Private Function GetBestRenderer() As PD_GradientRenderer
    
    If (m_GradientShape = gs_Linear) Then
        GetBestRenderer = gr_GDIPlus
    ElseIf (m_GradientShape = gs_Reflection) Then
        GetBestRenderer = gr_GDIPlus
    
    'All other shapes are only supported by our internal gradient renderer
    Else
        GetBestRenderer = gr_Internal
    End If
    
End Function

Private Function GetNameOfRenderer(ByVal rID As PD_GradientRenderer) As String
    If (rID = gr_Internal) Then
        GetNameOfRenderer = "PhotoDemon"
    ElseIf (rID = gr_GDIPlus) Then
        GetNameOfRenderer = "GDI+"
    End If
End Function

'A new test; attempt to maximize performance by translating the gradient to the current viewport space and only rendering it there.
' At _MouseUp(), a full-size preview will be manually rendered and committed.
Private Sub PreviewRenderer(ByRef srcCanvas As pdCanvas, ByRef firstPoint As PointFloat, ByRef secondPoint As PointFloat, Optional ByVal curBackend As PD_GradientRenderer = gr_Internal)
    
    'Retrieve a copy of the intersected viewport rect; we will use this for clipping
    Dim viewportIntersectRect As RectF
    PDImages.GetActiveImage.ImgViewport.GetIntersectRectCanvas viewportIntersectRect
    
    'Ensure we have a valid preview DIB
    If (m_PreviewDIB Is Nothing) Then Set m_PreviewDIB = New pdDIB
    
    'Initialize to the size of the current viewport.
    With m_PreviewDIB
        If (.GetDIBWidth <> srcCanvas.GetCanvasWidth) Or (.GetDIBHeight <> srcCanvas.GetCanvasHeight) Then
            Dim pDibWidth As Long, pDibHeight As Long
            With viewportIntersectRect
                pDibWidth = Int(.Width + PDMath.Frac(.Left) + 0.9999)
                pDibHeight = Int(.Height + PDMath.Frac(.Top) + 0.9999)
            End With
            m_PreviewDIB.CreateBlank pDibWidth, pDibHeight, 32, 0, 0
        Else
            m_PreviewDIB.ResetDIB 0
        End If
    End With
    
    'With the preview DIB created, we now need to translate the stored gradient endpoints from
    ' image space to viewport space.
    Dim newPoints() As PointFloat
    ReDim newPoints(0 To 1) As PointFloat
    CopyMemoryStrict VarPtr(newPoints(0)), VarPtr(m_Points(0)), 16&
    
    Dim cTransform As pd2DTransform
    Drawing.GetTransformFromImageToCanvas cTransform, srcCanvas, PDImages.GetActiveImage
    cTransform.ApplyTranslation -viewportIntersectRect.Left, -viewportIntersectRect.Top, P2_TO_Append
    cTransform.ApplyTransformToPointFs VarPtr(newPoints(0)), 2
    
    'Call the relevant renderer, which will proceed to draw a miniature version of the current gradient
    If (curBackend = gr_GDIPlus) Then
        GdipRenderer newPoints(0), newPoints(1), m_PreviewDIB
    ElseIf (curBackend = gr_Internal) Then
        InternalRenderer newPoints(0), newPoints(1), m_PreviewDIB, False
    End If
    
End Sub

Private Sub GdipRenderer(ByRef firstPoint As PointFloat, ByRef secondPoint As PointFloat, ByRef dstDIB As pdDIB)

    'Rendering methods are still being debated; cairo and GDI+ both have trade-offs depending
    ' on gradient parameters.
    Dim cSurface As pd2DSurface
    Set cSurface = New pd2DSurface
    cSurface.WrapSurfaceAroundPDDIB dstDIB
    cSurface.SetSurfaceAntialiasing P2_AA_None
    cSurface.SetSurfaceCompositing P2_CM_Overwrite
    cSurface.SetSurfacePixelOffset P2_PO_Normal
    
    'Populate any remaining gradient properties
    If (m_GradientShape = gs_Linear) Then
        m_GradientGdip.SetGradientShape P2_GS_Linear
    ElseIf (m_GradientShape = gs_Reflection) Then
        m_GradientGdip.SetGradientShape P2_GS_Reflection
    End If
    
    Dim gradAngle As Double
    gradAngle = PDMath.Atan2(secondPoint.y - firstPoint.y, secondPoint.x - firstPoint.x)
    m_GradientGdip.SetGradientAngle PDMath.RadiansToDegrees(gradAngle)
    
    'Fill the entire source
    Dim cBrush As pd2DBrush
    Set cBrush = New pd2DBrush
    cBrush.SetBrushMode P2_BM_Gradient
    
    Dim cRadius As Double
    cRadius = PDMath.DistanceTwoPoints(firstPoint.x, firstPoint.y, secondPoint.x, secondPoint.y)
    
    Dim boundsRect As RectF
    With boundsRect
        .Left = PDMath.Min2Float_Single(firstPoint.x, secondPoint.x)
        .Top = PDMath.Min2Float_Single(firstPoint.y, secondPoint.y)
        .Width = Abs(secondPoint.x - firstPoint.x)
        If (.Width < 1!) Then .Width = 1!
        .Height = Abs(secondPoint.y - firstPoint.y)
        If (.Height < 1!) Then .Height = 1!
    End With
    
    'We now have everything we need to render the gradient.  Unfortunately, certain edge-wrap modes
    ' (e.g. clamp) have no direct support in GDI+.  This greatly complicates their rendering, as we
    ' must manually clamp the results.  (Similarly, GDI+ does not allow you to render with any kind
    ' of "non-tiled/wrapped" behavior, so we must sometimes overwrite the result manually.)
    cBrush.SetBoundaryRect boundsRect
    cBrush.SetBrushGradientAllSettings m_GradientGdip.GetGradientAsString()
    
    'Because GDI+ is glitchy with certain wrap modes, we need to manually request custom wrap modes
    ' when natively unsupported modes are used
    If (m_GradientRepeat = gr_None) Then
        cBrush.SetBrushGradientWrapMode P2_WM_TileFlipXY
    ElseIf (m_GradientRepeat = gr_Clamp) Then
        cBrush.SetBrushGradientWrapMode P2_WM_TileFlipXY
    End If
    
    Dim slWidth As Single, slHeight As Single
    slWidth = dstDIB.GetDIBWidth()
    slHeight = dstDIB.GetDIBHeight()
    PD2D.FillRectangleF cSurface, cBrush, 0!, 0!, slWidth, slHeight
    
    'The gradient now covers the entire underlying scratch layer, for better or worse.
    
    'If the wrap mode is a mode unsupported by GDI+ (e.g. "extend/clamp"), we now need to manually
    ' overwrite the gradient in certain areas.
    If (m_GradientRepeat = gr_Clamp) Or (m_GradientRepeat = gr_None) Then
    
        'To overwrite the ends of the gradient (which have been forcibly tiled by GDI+),
        ' we need to perform some manual calculations.
        
        'First, we need to calculate lines that mark the ends of the gradient.  These lines will
        ' be perpendicular to the gradient direction.
        
        'We know the angle of the current line (calculated above).  Add/subtract PI/2 to rotate it
        ' 90 degrees in either direction.
        Dim angPerpendicular As Single, angPerpendicular2 As Single
        angPerpendicular = gradAngle + PI_HALF
        angPerpendicular2 = gradAngle - PI_HALF
        
        'There are two "end lines" for a gradient: one through each gradient end point.
        ' For each end point of the original gradient, calculate two new endpoints for a
        ' perpendicular line (with length equal to 2 * diagonal size of bounding box - since we are
        ' only using this for clipping, we deliberately want to make it extend beyond the edges
        ' of the current bounding box).
        Dim diagLength As Single
        diagLength = Sqr(slWidth * slWidth + slHeight * slHeight)
        
        Dim clipPoly() As PointFloat
        ReDim clipPoly(0 To 3) As PointFloat
        PDMath.ConvertPolarToCartesian_Sng angPerpendicular, diagLength, clipPoly(0).x, clipPoly(0).y, firstPoint.x, firstPoint.y
        PDMath.ConvertPolarToCartesian_Sng angPerpendicular2, diagLength, clipPoly(1).x, clipPoly(1).y, firstPoint.x, firstPoint.y
        
        'We now have two endpoints of the clip polygon we desire.  To generate the next two points,
        ' we can repeat our previous steps: rotate a point 90 degrees around the two points we've
        ' already calculated, which will give us a parallelogram defining a nice clip area.  Cool!
        ' The main thing we need to remember is to rotate the new points in the OPPOSITE direction
        ' of each anchor's previous rotation direction.
        angPerpendicular = angPerpendicular + PI_HALF
        angPerpendicular2 = angPerpendicular2 - PI_HALF
        PDMath.ConvertPolarToCartesian_Sng angPerpendicular, diagLength, clipPoly(3).x, clipPoly(3).y, clipPoly(0).x, clipPoly(0).y
        PDMath.ConvertPolarToCartesian_Sng angPerpendicular2, diagLength, clipPoly(2).x, clipPoly(2).y, clipPoly(1).x, clipPoly(1).y
        
        'Fill the area with the first color in the gradient
        Dim srcColor As RGBQuad
        m_GradientGdip.GetColorAtPosition_RGBA 0!, srcColor
        
        Set cBrush = New pd2DBrush
        If (m_GradientRepeat = gr_None) Then
            srcColor.Red = 0
            srcColor.Green = 0
            srcColor.Blue = 0
            srcColor.Alpha = 0
        End If
        Drawing2D.QuickCreateSolidBrush cBrush, RGB(srcColor.Red, srcColor.Green, srcColor.Blue), srcColor.Alpha / 2.55!
        PD2D.FillPolygonF_FromPtF cSurface, cBrush, 4, VarPtr(clipPoly(0))
        
        'Now we basically repeat all the above steps, but for the second gradient endpoint.
        ' (Naturally, we also swap the order of +/-90 points, to ensure that the polygon lies on the
        ' opposite side of the gradient.)
        angPerpendicular = gradAngle - PI_HALF
        angPerpendicular2 = gradAngle + PI_HALF
        PDMath.ConvertPolarToCartesian_Sng angPerpendicular, diagLength, clipPoly(0).x, clipPoly(0).y, secondPoint.x, secondPoint.y
        PDMath.ConvertPolarToCartesian_Sng angPerpendicular2, diagLength, clipPoly(1).x, clipPoly(1).y, secondPoint.x, secondPoint.y
        
        angPerpendicular = angPerpendicular + PI_HALF
        angPerpendicular2 = angPerpendicular2 - PI_HALF
        PDMath.ConvertPolarToCartesian_Sng angPerpendicular, diagLength, clipPoly(3).x, clipPoly(3).y, clipPoly(0).x, clipPoly(0).y
        PDMath.ConvertPolarToCartesian_Sng angPerpendicular2, diagLength, clipPoly(2).x, clipPoly(2).y, clipPoly(1).x, clipPoly(1).y
        
        'Fill the new clip area with the last color in the gradient
        If (m_GradientShape = gs_Linear) Then m_GradientGdip.GetColorAtPosition_RGBA 1!, srcColor
        If (m_GradientRepeat = gr_None) Then
            srcColor.Red = 0
            srcColor.Green = 0
            srcColor.Blue = 0
            srcColor.Alpha = 0
        End If
        Drawing2D.QuickCreateSolidBrush cBrush, RGB(srcColor.Red, srcColor.Green, srcColor.Blue), srcColor.Alpha / 2.55!
        PD2D.FillPolygonF_FromPtF cSurface, cBrush, 4, VarPtr(clipPoly(0))
        
        'Free our intermediary "fix" brush
        Set cBrush = Nothing
        
    End If
    
    'Free all handles
    Set cBrush = Nothing
    Set cSurface = Nothing
    
End Sub

Private Sub InternalRenderer(ByRef firstPoint As PointFloat, ByRef secondPoint As PointFloat, ByRef dstDIB As pdDIB, Optional ByVal isFullSizeRender As Boolean = True)
    
    'On final renders, display a progress bar (as the render may take up to 500 ms on large images)
    If isFullSizeRender Then ProgressBars.SetProgBarMax dstDIB.GetDIBHeight
    
    'If the repeat mode is "clamp", perform a fast fill of the back buffer now.
    ' (Note that linear mode does *not* require this, as it clamps differently.)
    If (m_GradientRepeat = gr_Clamp) And (m_GradientShape <> gs_Linear) Then
        Dim cBrush As pd2DBrush, cSurface As pd2DSurface
        Drawing2D.QuickCreateSurfaceFromDIB cSurface, dstDIB, False
        cSurface.SetSurfaceCompositing P2_CM_Overwrite
        Drawing2D.QuickCreateSolidBrush cBrush, RGB(m_LastGradColorQuad.Red, m_LastGradColorQuad.Green, m_LastGradColorQuad.Blue), m_LastGradColorQuad.Alpha / 2.55
        PD2D.FillRectangleI cSurface, cBrush, 0, 0, dstDIB.GetDIBWidth, dstDIB.GetDIBHeight
        Set cBrush = Nothing
        Set cSurface = Nothing
    End If
    
    If (m_GradientShape = gs_Radial) Then
        InternalRender_Radial firstPoint, secondPoint, dstDIB, isFullSizeRender
    ElseIf (m_GradientShape = gs_Spherical) Then
        InternalRender_Spherical firstPoint, secondPoint, dstDIB, isFullSizeRender
    ElseIf (m_GradientShape = gs_Square) Then
        InternalRender_Square firstPoint, secondPoint, dstDIB, isFullSizeRender
    ElseIf (m_GradientShape = gs_Diamond) Then
        InternalRender_Diamond firstPoint, secondPoint, dstDIB, isFullSizeRender
    ElseIf (m_GradientShape = gs_Conical) Then
        InternalRender_Conical firstPoint, secondPoint, dstDIB, isFullSizeRender
    ElseIf (m_GradientShape = gs_Spiral) Then
        InternalRender_Spiral firstPoint, secondPoint, dstDIB, isFullSizeRender
    End If
    
    'If we displayed a progress bar, free it now
    If isFullSizeRender Then
        ProgressBars.SetProgBarVal 0
        ProgressBars.ReleaseProgressBar
    End If
    
End Sub

Private Sub InternalRender_Conical(ByRef firstPoint As PointFloat, ByRef secondPoint As PointFloat, ByRef dstDIB As pdDIB, Optional ByVal isFullSizeRender As Boolean = True)

    'Before doing anything else, calculate some helpful geometry shortcuts
    Dim gradAngle As Double
    gradAngle = PDMath.Atan2(secondPoint.y - firstPoint.y, secondPoint.x - firstPoint.x)
    gradAngle = PDMath.Modulo(gradAngle, PI_DOUBLE)
    
    Dim lutBound As Long
    lutBound = m_LookupResolution - 1
    
    Dim oX As Double, oY As Double
    oX = firstPoint.x
    oY = firstPoint.y
    
    'Unlike other gradient modes, conical gradients don't really have need of a "wrap" property.
    ' This is because they are measured by angle, not distance, so they always extend to the
    ' boundary of the visible area.
    '
    'That said, we can "cheat" and make the gradient behave a little differently according to
    ' wrap mode.  This spares us having to create different gradient tools for e.g. symmetrical
    ' vs asymmetrical conical patterns.
    Dim conLimit As Double
    If (m_GradientRepeat = gr_None) Then
        conLimit = PI_DOUBLE
    ElseIf (m_GradientRepeat = gr_Clamp) Then
        conLimit = PI_DOUBLE
    ElseIf (m_GradientRepeat = gr_Wrap) Then
        conLimit = PI
    ElseIf (m_GradientRepeat = gr_Reflect) Then
        conLimit = PI
    End If
    
    'Construct a mapping between gradient distance and lookup table size
    Dim mapAdjust As Double
    mapAdjust = CDbl(lutBound) / conLimit
    
    'Wrap an array around the destination DIB
    Dim x As Long, y As Long, xBound As Long, yBound As Long
    Dim dstPixels() As Long, dstSA As SafeArray1D, dstPtr As Long, dstStride As Long
    dstPtr = dstDIB.GetDIBPointer
    dstStride = dstDIB.GetDIBStride
    dstDIB.WrapLongArrayAroundScanline dstPixels, dstSA, 0
    xBound = dstDIB.GetDIBWidth - 1
    yBound = dstDIB.GetDIBHeight - 1
    
    'If this is a full-size render, it may take a second or two, so prep progress bars notifiers
    Dim progBarInterval As Long
    progBarInterval = ProgressBars.FindBestProgBarValue()
    
    Dim curAngle As Double, curColor As Long, luIndex As Long
    
    For y = 0 To yBound
        dstSA.pvData = dstPtr + dstStride * y
    For x = 0 To xBound
        
        'Calculate angle, and remap it (with rounding) to a lookup table index
        curAngle = PDMath.Atan2_Faster(y - oY, x - oX)
        curAngle = curAngle - gradAngle
        
        'Reflect mode operates a little differently; we want the gradient to be symmetrical across
        ' the axis of the gradient line, so we want to reflect the result across pi
        If (m_GradientRepeat = gr_Reflect) Then
            curAngle = PDMath.Modulo(curAngle, conLimit * 2)
            If (curAngle > PI) Then curAngle = PI_DOUBLE - curAngle
        Else
            curAngle = PDMath.Modulo(curAngle, conLimit)
        End If
        
        luIndex = Int(curAngle * mapAdjust + 0.5)
        curColor = m_GradLookup(luIndex)
        
        'No further interpolation is required.  (In the future, we could add a dithering
        ' element here, but for now let's just focus on getting the gradient itself rendered OK.)
        dstPixels(x) = curColor
        
    Next x
        If isFullSizeRender Then
            If ((y And progBarInterval) = 0) Then ProgressBars.SetProgBarVal y
        End If
    Next y
    
    dstDIB.UnwrapLongArrayFromDIB dstPixels
    
End Sub

Private Sub InternalRender_Diamond(ByRef firstPoint As PointFloat, ByRef secondPoint As PointFloat, ByRef dstDIB As pdDIB, Optional ByVal isFullSizeRender As Boolean = True)

    'Before doing anything else, calculate some helpful geometry shortcuts
    Dim gradDistance As Double
    gradDistance = PDMath.DistanceTwoPoints(firstPoint.x, firstPoint.y, secondPoint.x, secondPoint.y)
    If (gradDistance < 1#) Then gradDistance = 1#
    
    Dim gradAngle As Double
    gradAngle = -1# * PDMath.Atan2(secondPoint.y - firstPoint.y, secondPoint.x - firstPoint.x)
    
    'For performance reasons, we want to cache the cos and sin of the angle in question
    Dim sinAngle As Double, cosAngle As Double
    sinAngle = Sin(gradAngle)
    cosAngle = Cos(gradAngle)
    
    Dim lutBound As Long
    lutBound = m_LookupResolution - 1
    
    'Construct a mapping between gradient distance and lookup table size
    Dim mapAdjust As Double
    mapAdjust = CDbl(lutBound) / gradDistance
    
    'Wrap an array around the destination DIB
    Dim x As Long, y As Long, xBound As Long, yBound As Long
    Dim dstPixels() As Long, dstSA As SafeArray1D, dstPtr As Long, dstStride As Long
    dstPtr = dstDIB.GetDIBPointer
    dstStride = dstDIB.GetDIBStride
    dstDIB.WrapLongArrayAroundScanline dstPixels, dstSA, 0
    xBound = dstDIB.GetDIBWidth - 1
    yBound = dstDIB.GetDIBHeight - 1
    
    'If this is a full-size render, it may take a second or two, so prep progress bars notifiers
    Dim progBarInterval As Long
    progBarInterval = ProgressBars.FindBestProgBarValue()
    
    Dim curDistance As Double, curColor As Long, luIndex As Long, yFast As Double
    Dim centerX As Double, centerY As Double, newX As Double, newY As Double
    centerX = firstPoint.x
    centerY = firstPoint.y
    
    For y = 0 To yBound
        
        dstSA.pvData = dstPtr + dstStride * y
        
        'Because y-terms do not change for the length of this scanline, precalculate what we can
        yFast = CDbl(y) - centerY
        
    For x = 0 To xBound
        
        'Rotate this point around the origin (where the user's first clicked point is the origin) by the
        ' same angle as the gradient line; the "diamond" is then calculated using manhattan distance.
        ' (Note that multiple terms here could be pre-cached into luts for an even bigger perf boost.)
        newX = cosAngle * (CDbl(x) - centerX) - sinAngle * yFast + centerX
        newY = cosAngle * yFast + sinAngle * (CDbl(x) - centerX) + centerY
        
        curDistance = Abs(newX - centerX) + Abs(newY - centerY)
        luIndex = Int(curDistance * mapAdjust + 0.5)
        
        'Pixels beyond the edge of the user-drawn line must be handled according to the current wrap setting
        If (curDistance > gradDistance) Then
            
            'Edge mode handling here
            Select Case m_GradientRepeat
                
                'None: ignore pixel
                Case gr_None
                    GoTo NextPixel
                    
                'Clamp: paint as the terminal gradient color.  (This was handled in a previous step.)
                Case gr_Clamp
                    GoTo NextPixel
                
                'Reflect: shift phase by 1, remap to [-1, 1], and reflect negative values
                Case gr_Reflect
                    luIndex = luIndex - lutBound
                    luIndex = (luIndex And (lutBound * 2)) - lutBound
                    If (luIndex < 0) Then luIndex = -luIndex
                    curColor = m_GradLookup(luIndex)
                
                'Wrap: re-map the calculated index mod the number of items in the lookup table
                Case gr_Wrap
                    curColor = m_GradLookup(luIndex And lutBound)
                
            End Select
            
        Else
            curColor = m_GradLookup(luIndex)
        End If
        
        'No further interpolation is required.  (In the future, we could add a dithering
        ' element here, but for now let's just focus on getting the gradient itself rendered OK.)
        dstPixels(x) = curColor
        
NextPixel:
        
    Next x
        If isFullSizeRender Then
            If ((y And progBarInterval) = 0) Then ProgressBars.SetProgBarVal y
        End If
    Next y
    
    dstDIB.UnwrapLongArrayFromDIB dstPixels
    
End Sub

Private Sub InternalRender_Radial(ByRef firstPoint As PointFloat, ByRef secondPoint As PointFloat, ByRef dstDIB As pdDIB, Optional ByVal isFullSizeRender As Boolean = True)
    
    'Before doing anything else, calculate some helpful geometry shortcuts
    Dim gradDistance As Double
    gradDistance = PDMath.DistanceTwoPoints(firstPoint.x, firstPoint.y, secondPoint.x, secondPoint.y)
    If (gradDistance <= 1#) Then gradDistance = 1#
    
    Dim lutBound As Long
    lutBound = m_LookupResolution - 1
    
    Dim oX As Double, oY As Double
    oX = firstPoint.x
    oY = firstPoint.y
    
    'Construct a mapping between gradient distance and lookup table size
    Dim mapAdjust As Double
    mapAdjust = CDbl(lutBound) / gradDistance
    
    'Wrap an array around the destination DIB
    Dim x As Long, y As Long, xBound As Long, yBound As Long
    Dim dstPixels() As Long, dstSA As SafeArray1D, dstPtr As Long, dstStride As Long
    dstPtr = dstDIB.GetDIBPointer
    dstStride = dstDIB.GetDIBStride
    dstDIB.WrapLongArrayAroundScanline dstPixels, dstSA, 0
    xBound = dstDIB.GetDIBWidth - 1
    yBound = dstDIB.GetDIBHeight - 1
    
    'If this is a full-size render, it may take a second or two, so prep progress bars notifiers
    Dim progBarInterval As Long
    progBarInterval = ProgressBars.FindBestProgBarValue()
    
    Dim curDistance As Double, curColor As Long, luIndex As Long, ySquared As Long
    
    For y = 0 To yBound
        dstSA.pvData = dstPtr + dstStride * y
        ySquared = (y - oY) * (y - oY)
    For x = 0 To xBound
        
        'Calculate distance, and remap it (with rounding) to a lookup table index
        curDistance = Sqr((x - oX) * (x - oX) + ySquared)
        luIndex = Int(curDistance * mapAdjust + 0.5)
        
        'Pixels beyond the edge of the user-drawn line must be handled according to the current wrap setting
        If (curDistance > gradDistance) Then
            
            'Edge mode handling here
            Select Case m_GradientRepeat
                
                'None: ignore pixel
                Case gr_None
                    GoTo NextPixel
                    
                'Clamp: paint as the terminal gradient color.  (This was handled in a previous step.)
                Case gr_Clamp
                    GoTo NextPixel
                
                'Reflect: shift phase by 1, remap to [-1, 1], and reflect negative values
                Case gr_Reflect
                    luIndex = luIndex - lutBound
                    luIndex = (luIndex And (lutBound * 2)) - lutBound
                    If (luIndex < 0) Then luIndex = -luIndex
                    curColor = m_GradLookup(luIndex)
                
                'Wrap: re-map the calculated index mod the number of items in the lookup table
                Case gr_Wrap
                    curColor = m_GradLookup(luIndex And lutBound)
                
            End Select
            
        Else
            curColor = m_GradLookup(luIndex)
        End If
        
        'No further interpolation is required.  (In the future, we could add a dithering
        ' element here, but for now let's just focus on getting the gradient itself rendered OK.)
        dstPixels(x) = curColor
        
NextPixel:
        
    Next x
        If isFullSizeRender Then
            If ((y And progBarInterval) = 0) Then ProgressBars.SetProgBarVal y
        End If
    Next y
    
    dstDIB.UnwrapLongArrayFromDIB dstPixels
    
End Sub

'Math for this offset radial renderer was developed with help from the pixman project (MIT-licensed).
' Pixman's algebraic solution is also used by other major renderers, including Skia. See http://pixman.org/
' for details, especially https://github.com/freedesktop/pixman/blob/master/pixman/pixman-radial-gradient.c
' for an excellent breakdown of the required geometry.
Private Sub InternalRender_Spherical(ByRef firstPoint As PointFloat, ByRef secondPoint As PointFloat, ByRef dstDIB As pdDIB, Optional ByVal isFullSizeRender As Boolean = True)
    
    'Before doing anything else, calculate some helpful geometry shortcuts
    Dim gradDistance As Double
    gradDistance = PDMath.DistanceTwoPoints(firstPoint.x, firstPoint.y, secondPoint.x, secondPoint.y)
    If (gradDistance <= 1#) Then gradDistance = 1#
    
    Dim lutBound As Long
    lutBound = m_LookupResolution - 1
    
    Dim oX As Double, oY As Double
    oX = firstPoint.x
    oY = firstPoint.y
    
    'Construct a mapping between gradient distance and lookup table size
    Dim mapAdjust As Double
    mapAdjust = CDbl(lutBound) / gradDistance
    
    'Wrap an array around the destination DIB
    Dim x As Long, y As Long, xBound As Long, yBound As Long
    Dim dstPixels() As Long, dstSA As SafeArray1D, dstPtr As Long, dstStride As Long
    dstPtr = dstDIB.GetDIBPointer
    dstStride = dstDIB.GetDIBStride
    dstDIB.WrapLongArrayAroundScanline dstPixels, dstSA, 0
    xBound = dstDIB.GetDIBWidth - 1
    yBound = dstDIB.GetDIBHeight - 1
    
    'If this is a full-size render, it may take a second or two, so prep progress bars notifiers
    Dim progBarInterval As Long
    progBarInterval = ProgressBars.FindBestProgBarValue()
    
    Dim curDistance As Double, curColor As Long, luIndex As Long
    Dim dx As Double, dy As Double, cX1 As Double, cY1 As Double
    
    'Gradients are typically specified as two circles, each with independent radii and center points.
    ' To simplify mouse interactions for the user, we automatically produce the internal circle
    ' center point "on the fly" and force its radius to 0; this simplifies some subsequent calculations.
    dx = secondPoint.x - firstPoint.x
    dy = secondPoint.y - firstPoint.y
    cX1 = firstPoint.x + (dx * m_RadialOffset)
    cY1 = firstPoint.y + (dy * m_RadialOffset)
    
    'cX1/cY1 now describe the center point of the first circle; cX2/cY2 describe the center point of
    ' the second (outer) circle, with radius = distance of the user's mouse drag
    Dim cX2 As Double, cY2 As Double
    cX2 = firstPoint.x: cY2 = firstPoint.y
    
    Dim cdx As Double, cdy As Double
    cdx = cX2 - cX1
    cdy = cY2 - cY1
    
    'The following math involves some messy solutions for calculating distance between the two circles
    ' we've constructed.  Thank you to the authors of the MIT-licensed pixman project for their helpful
    ' breakdown of this technique.
    Dim a As Double, invA As Double, b As Double, c As Double, t As Double
    a = cdx * cdx + cdy * cdy - gradDistance * gradDistance
    If (a <> 0#) Then invA = 1# / a
    
    For y = 0 To yBound
        dstSA.pvData = dstPtr + dstStride * y
    For x = 0 To xBound
        
        'The goal here is simple enough: we want to scale the radius by some factor that's
        ' contingent on the angle of the current point.  (Points lying close to the user-drawn
        ' gradient line are scaled very little, while points on the opposite side of the
        ' "circle" are scaled more.)
        
        'Note that the actual required calculations are more complex than the ones used here;
        ' we shortcut certain elements because the radius of the inset circle is always forced
        ' to 0 in PD.
        b = (x - cX1) * cdx + (y - cY1) * cdy
        c = (x - cX1) * (x - cX1) + (y - cY1) * (y - cY1)
        curDistance = (b * b - a * c)
        
        'We need to calculate Sqr, so ensure we have a non-zero value
        If (curDistance < 0.0000001) Then curDistance = 0.0000001
        t = (b - Sqr(curDistance)) * invA
        
        'Failsafe check to prevent OOB errors when indexing into the LUT
        If (t < 0) Then t = 0#
        
        'Convert the lookup to a color in our prebuilt lookup table, then update the distance
        ' tracker to reflect the value of t.  (The distance tracker is used to determine pixel
        ' handling when pixels fall outside the radial boundary; this is contingent on the
        ' user's edge handling mode.)
        luIndex = Int(t * lutBound + 0.5)
        curDistance = t * gradDistance
        
        'Pixels beyond the edge of the user-drawn line must be handled according to the current wrap setting
        If (curDistance > gradDistance) Then
            
            'Edge mode handling here
            Select Case m_GradientRepeat
                
                'None: ignore pixel
                Case gr_None
                    GoTo NextPixel
                    
                'Clamp: paint as the terminal gradient color.  (This was handled in a previous step.)
                Case gr_Clamp
                    GoTo NextPixel
                
                'Reflect: shift phase by 1, remap to [-1, 1], and reflect negative values
                Case gr_Reflect
                    luIndex = luIndex - lutBound
                    luIndex = (luIndex And (lutBound * 2)) - lutBound
                    If (luIndex < 0) Then luIndex = -luIndex
                    curColor = m_GradLookup(luIndex)
                
                'Wrap: re-map the calculated index mod the number of items in the lookup table
                Case gr_Wrap
                    curColor = m_GradLookup(luIndex And lutBound)
                
            End Select
            
        Else
            curColor = m_GradLookup(luIndex)
        End If
        
        'No further interpolation is required.  (In the future, we could add a dithering
        ' element here, but for now let's just focus on getting the gradient itself rendered OK.)
        dstPixels(x) = curColor
        
NextPixel:
        
    Next x
        If isFullSizeRender Then
            If ((y And progBarInterval) = 0) Then ProgressBars.SetProgBarVal y
        End If
    Next y
    
    dstDIB.UnwrapLongArrayFromDIB dstPixels
    
End Sub

Private Sub InternalRender_Spiral(ByRef firstPoint As PointFloat, ByRef secondPoint As PointFloat, ByRef dstDIB As pdDIB, Optional ByVal isFullSizeRender As Boolean = True)

    'Before doing anything else, calculate some helpful geometry shortcuts
    Dim gradAngle As Double
    gradAngle = PDMath.Atan2(secondPoint.y - firstPoint.y, secondPoint.x - firstPoint.x)
    gradAngle = PDMath.Modulo(gradAngle, PI_DOUBLE)
    
    Dim gradDistance As Double
    gradDistance = PDMath.DistanceTwoPoints(firstPoint.x, firstPoint.y, secondPoint.x, secondPoint.y)
    If (gradDistance <= 1#) Then gradDistance = 1#
    
    'Map gradient angle to the same magnitude as gradient distance
    Dim gradAngleMapping As Double
    gradAngleMapping = (gradDistance / PI_DOUBLE)
    
    Dim lutBound As Long
    lutBound = m_LookupResolution - 1
    
    Dim oX As Double, oY As Double
    oX = firstPoint.x
    oY = firstPoint.y
    
    'Construct a mapping between gradient distance and lookup table size
    Dim mapAdjust As Double
    mapAdjust = CDbl(lutBound) / gradDistance
    
    'Wrap an array around the destination DIB
    Dim x As Long, y As Long, xBound As Long, yBound As Long
    Dim dstPixels() As Long, dstSA As SafeArray1D, dstPtr As Long, dstStride As Long
    dstPtr = dstDIB.GetDIBPointer
    dstStride = dstDIB.GetDIBStride
    dstDIB.WrapLongArrayAroundScanline dstPixels, dstSA, 0
    xBound = dstDIB.GetDIBWidth - 1
    yBound = dstDIB.GetDIBHeight - 1
    
    'If this is a full-size render, it may take a second or two, so prep progress bars notifiers
    Dim progBarInterval As Long
    progBarInterval = ProgressBars.FindBestProgBarValue()
    
    Dim curDistance As Double, curAngle As Double, curColor As Long, luIndex As Long
    Dim yFast As Long, ySquared As Long
    
    For y = 0 To yBound
        dstSA.pvData = dstPtr + dstStride * y
        
        'For a minor perf improvement, we can calculate y and y^2 terms outside the inner loop
        yFast = (y - oY)
        ySquared = yFast * yFast
    For x = 0 To xBound
        
        'Calculate distance
        curDistance = Sqr((x - oX) * (x - oX) + ySquared)
        
        'Calculate angle, remap it against the angle of the gradient line, then normalize it
        ' (in radians, remember)
        curAngle = PDMath.Atan2_Faster(yFast, x - oX)
        curAngle = curAngle - gradAngle
        curAngle = PDMath.Modulo(curAngle, PI_DOUBLE)
        
        If (m_GradientRepeat = gr_None) Or (m_GradientRepeat = gr_Clamp) Then
            
            'Calculate the spiral
            curDistance = curDistance + curAngle * gradAngleMapping
            
            'If the point lies outside the first spiral in the pattern, skip this pixel
            If (curDistance > gradDistance) Then
                GoTo NextPixel
            Else
                luIndex = Int(curDistance * mapAdjust + 0.5)
                curColor = m_GradLookup(luIndex)
            End If
        
        ElseIf (m_GradientRepeat = gr_Wrap) Then
            
            'Calculate the spiral, and simply force the result to a multiple of the original
            ' gradient distance.
            curDistance = curDistance + curAngle * gradAngleMapping
            curDistance = PDMath.Modulo(curDistance, gradDistance)
            luIndex = Int(curDistance * mapAdjust + 0.5)
            curColor = m_GradLookup(luIndex)
        
        ElseIf (m_GradientRepeat = gr_Reflect) Then
            
            'Multiply the angle by 2.0 (to ensure the pattern aligns along the axis boundary)
            ' then force the result to a multiple of *double* the gradient distance.  This allows
            ' us to calculate a reflection when the result is on the range [original, 2 * original].
            curDistance = curDistance + curAngle * gradAngleMapping * 2#
            curDistance = PDMath.Modulo(curDistance, gradDistance * 2#)
            If (curDistance > gradDistance) Then curDistance = gradDistance * 2# - curDistance
            luIndex = Int(curDistance * mapAdjust + 0.5)
            curColor = m_GradLookup(luIndex)
            
        End If
        
        'No further interpolation is required.  (In the future, we could add a dithering
        ' element here, but for now let's just focus on getting the gradient itself rendered OK.)
        dstPixels(x) = curColor
        
NextPixel:

    Next x
        If isFullSizeRender Then
            If ((y And progBarInterval) = 0) Then ProgressBars.SetProgBarVal y
        End If
    Next y
    
    dstDIB.UnwrapLongArrayFromDIB dstPixels
    
End Sub

Private Sub InternalRender_Square(ByRef firstPoint As PointFloat, ByRef secondPoint As PointFloat, ByRef dstDIB As pdDIB, Optional ByVal isFullSizeRender As Boolean = True)
    
    'Calculate the angle of the user's current line
    Dim gradAngle As Double
    gradAngle = -1# * PDMath.Atan2(secondPoint.y - firstPoint.y, secondPoint.x - firstPoint.x)
    
    'For performance reasons, we want to cache the cos and sin of the angle in question
    Dim sinAngle As Double, cosAngle As Double
    sinAngle = Sin(gradAngle)
    cosAngle = Cos(gradAngle)
    
    'Rotate the current line to angle (0)
    Dim newX As Single, newY As Single
    PDMath.RotatePointAroundPoint secondPoint.x, secondPoint.y, firstPoint.x, firstPoint.y, gradAngle, newX, newY
    
    'Before doing anything else, calculate some helpful geometry shortcuts
    Dim gradDistance As Single
    gradDistance = PDMath.Max2Int(Abs(firstPoint.x - newX), Abs(firstPoint.y - newY))
    If (gradDistance < 1) Then gradDistance = 1
    
    Dim lutBound As Long
    lutBound = m_LookupResolution - 1
    
    Dim oX As Long, oY As Long
    oX = firstPoint.x
    oY = firstPoint.y
    
    'Construct a mapping between gradient distance and lookup table size
    Dim mapAdjust As Double
    mapAdjust = CDbl(lutBound) / gradDistance
    
    'Wrap an array around the destination DIB
    Dim x As Long, y As Long, xBound As Long, yBound As Long
    Dim dstPixels() As Long, dstSA As SafeArray1D, dstPtr As Long, dstStride As Long
    dstPtr = dstDIB.GetDIBPointer
    dstStride = dstDIB.GetDIBStride
    dstDIB.WrapLongArrayAroundScanline dstPixels, dstSA, 0
    xBound = dstDIB.GetDIBWidth - 1
    yBound = dstDIB.GetDIBHeight - 1
    
    'If this is a full-size render, it may take a second or two, so prep progress bars notifiers
    Dim progBarInterval As Long
    progBarInterval = ProgressBars.FindBestProgBarValue()
    
    Dim curDistance As Double, curColor As Long, luIndex As Long, yFast As Double
    Dim centerX As Double, centerY As Double
    centerX = firstPoint.x
    centerY = firstPoint.y
    
    For y = 0 To yBound
    
        dstSA.pvData = dstPtr + dstStride * y
        
        'Because y-terms do not change for the length of this scanline, precalculate what we can
        yFast = CDbl(y) - centerY
        
    For x = 0 To xBound
        
        'Rotate this point around the origin (where the user's first clicked point is the origin) by the
        ' same angle as the gradient line; the "square" is then calculated using the largest absolute value
        ' of either the x-delta or y-delta of the corresponding line.
        ' (Note that multiple terms here could be pre-cached into luts for an even bigger perf boost.)
        newX = cosAngle * (CDbl(x) - centerX) - sinAngle * yFast + centerX
        newY = cosAngle * yFast + sinAngle * (CDbl(x) - centerX) + centerY
        
        newX = Abs(newX - centerX)
        newY = Abs(newY - centerY)
        
        If newX > newY Then
            curDistance = newX
        Else
            curDistance = newY
        End If
        
        luIndex = Int(CDbl(curDistance) * mapAdjust + 0.5)
        
        'Pixels beyond the edge of the user-drawn line must be handled according to the current wrap setting
        If (curDistance > gradDistance) Then
            
            'Edge mode handling here
            Select Case m_GradientRepeat
                
                'None: ignore pixel
                Case gr_None
                    GoTo NextPixel
                    
                'Clamp: paint as the terminal gradient color.  (This was handled in a previous step.)
                Case gr_Clamp
                    GoTo NextPixel
                
                'Reflect: shift phase by 1, remap to [-1, 1], and reflect negative values
                Case gr_Reflect
                    luIndex = luIndex - lutBound
                    luIndex = (luIndex And (lutBound * 2)) - lutBound
                    If (luIndex < 0) Then luIndex = -luIndex
                    curColor = m_GradLookup(luIndex)
                
                'Wrap: re-map the calculated index mod the number of items in the lookup table
                Case gr_Wrap
                    curColor = m_GradLookup(luIndex And lutBound)
                
            End Select
            
        Else
            curColor = m_GradLookup(luIndex)
        End If
        
        'No further interpolation is required.  (In the future, we could add a dithering
        ' element here, but for now let's just focus on getting the gradient itself rendered OK.)
        dstPixels(x) = curColor
        
NextPixel:

    Next x
        If isFullSizeRender Then
            If ((y And progBarInterval) = 0) Then ProgressBars.SetProgBarVal y
        End If
    Next y
    
    dstDIB.UnwrapLongArrayFromDIB dstPixels
    
End Sub

'Want to commit your current gradient work?  Call this function to make the gradient results permanent.
Public Sub CommitGradientResults()
        
    'This dummy string only exists to ensure that the processor name gets localized properly
    ' (as that text is used for Undo/Redo descriptions).  PD's translation engine will detect
    ' the TranslateMessage() call and produce a matching translation entry.
    Dim strDummy As String
    strDummy = g_Language.TranslateMessage("Gradient tool")
    
    'Note that gradients do not (currently) maintain a changed-area rect.  Pass a rect that matches the
    ' full size of the scratch layer.
    Dim tmpRectF As RectF
    With tmpRectF
        .Left = 0
        .Top = 0
        .Width = PDImages.GetActiveImage.ScratchLayer.GetLayerDIB.GetDIBWidth
        .Height = PDImages.GetActiveImage.ScratchLayer.GetLayerDIB.GetDIBHeight
    End With
    
    Layers.CommitScratchLayer "Gradient tool", tmpRectF
    
End Sub

Public Sub RenderGradientUI(ByRef targetCanvas As pdCanvas)
    
    If (Not PDImages.IsImageActive()) Then Exit Sub
    
    'Clone a pair of UI pens from the main rendering module.  (Note that we clone unique pens instead
    ' of simply borrowing the shared UI pens as we may need to modify rendering properties, and we don't
    ' want to fuck up pens that are shared across other places in PD.)
    Dim basePenInactive As pd2DPen, topPenInactive As pd2DPen
    Dim basePenActive As pd2DPen, topPenActive As pd2DPen
    Drawing.CloneCachedUIPens basePenInactive, topPenInactive, False
    Drawing.CloneCachedUIPens basePenActive, topPenActive, True
    
    'Create other required pd2D drawing tools (a surface)
    Dim cSurface As pd2DSurface
    Drawing2D.QuickCreateSurfaceFromDC cSurface, targetCanvas.hDC, True
    cSurface.SetSurfacePixelOffset P2_PO_Normal
        
    'Mousedown/up obviously affects the UI elements that we render
    If m_MouseDown Then
    
        'Ensure we actually have points to operate on
        If (Not m_PointsInitialized) Then InitializePoints
        If (m_Points(0).x = MOUSE_OOB) Then Exit Sub
        If (m_Points(1).x = MOUSE_OOB) Then Exit Sub
        
        'Start by converting the original mouse positions from image coords to canvas coords
        Dim canvasCoordsX() As Double, canvasCoordsY() As Double
        ReDim canvasCoordsX(0 To 1) As Double
        ReDim canvasCoordsY(0 To 1) As Double
        
        Dim i As Long
        For i = 0 To 1
            Drawing.ConvertImageCoordsToCanvasCoords targetCanvas, PDImages.GetActiveImage(), m_Points(i).x, m_Points(i).y, canvasCoordsX(i), canvasCoordsY(i)
        Next i
        
        'Specify rounded line edges for our pens; this looks better for this particular tool
        basePenInactive.SetPenStartCap P2_LC_Round
        topPenInactive.SetPenStartCap P2_LC_Round
        basePenInactive.SetPenEndCap P2_LC_ArrowAnchor
        topPenInactive.SetPenEndCap P2_LC_ArrowAnchor
        
        basePenActive.SetPenLineCap P2_LC_Round
        topPenActive.SetPenLineCap P2_LC_Round
        
        basePenInactive.SetPenLineJoin P2_LJ_Round
        topPenInactive.SetPenLineJoin P2_LJ_Round
        basePenActive.SetPenLineJoin P2_LJ_Round
        topPenActive.SetPenLineJoin P2_LJ_Round
        
        'Stroke an arrow in the direction of the current gradient mouse-drag
        PD2D.DrawLineF cSurface, basePenInactive, canvasCoordsX(0), canvasCoordsY(0), canvasCoordsX(1), canvasCoordsY(1)
        PD2D.DrawLineF cSurface, topPenInactive, canvasCoordsX(0), canvasCoordsY(0), canvasCoordsX(1), canvasCoordsY(1)
        
    Else
    
        'Convert the current stored mouse coordinates from image coordinate space to viewport coordinate space
        Dim cursX As Double, cursY As Double
        Drawing.ConvertImageCoordsToCanvasCoords targetCanvas, PDImages.GetActiveImage(), m_MouseX, m_MouseY, cursX, cursY
        
        'Paint a target cursor
        Dim crossLength As Single, outerCrossBorder As Single
        crossLength = 5#
        outerCrossBorder = 0.5
        
        PD2D.DrawLineF cSurface, basePenInactive, cursX, cursY - crossLength - outerCrossBorder, cursX, cursY + crossLength + outerCrossBorder
        PD2D.DrawLineF cSurface, basePenInactive, cursX - crossLength - outerCrossBorder, cursY, cursX + crossLength + outerCrossBorder, cursY
        PD2D.DrawLineF cSurface, topPenInactive, cursX, cursY - crossLength, cursX, cursY + crossLength
        PD2D.DrawLineF cSurface, topPenInactive, cursX - crossLength, cursY, cursX + crossLength, cursY
    
    End If
    
    Set cSurface = Nothing
    Set basePenInactive = Nothing: Set topPenInactive = Nothing
    Set basePenActive = Nothing: Set topPenActive = Nothing
    
End Sub
    
Private Sub InitializePoints()
    m_PointsInitialized = True
    ReDim m_Points(0 To 1) As PointFloat
    m_Points(0).x = MOUSE_OOB
    m_Points(0).y = MOUSE_OOB
    m_Points(1).x = MOUSE_OOB
    m_Points(1).y = MOUSE_OOB
End Sub
