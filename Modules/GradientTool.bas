Attribute VB_Name = "Tools_Gradient"
'***************************************************************************
'PhotoDemon On-Canvas Gradient Tool Manager
'Copyright 2018-2019 by Tanner Helland
'Created: 31/December/18
'Last updated: 08/January/19
'Last update: finish up initial prototype
'
'This module interfaces between the gradient tool UI and pd2DGradient backend.  Look in the relevant
' tool panel form for more details on how the UI relays relevant fill data here.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Development-time parameter only, remove in production
Private Const USE_FAST_PREVIEW As Boolean = True

'Development-time parameter only, remove in production
Private Const USE_CAIRO_RENDERER As Boolean = False

Public Enum PD_GradientAttributes
    GA_Opacity = 0
    GA_BlendMode = 1
    GA_AlphaMode = 2
    GA_Antialiasing = 3
    GA_Repeat = 4
End Enum

#If False Then
    Private Const GA_Opacity = 0, GA_BlendMode = 1, GA_AlphaMode = 2, GA_Antialiasing = 3, GA_Repeat = 4
#End If

Public Enum PD_GradientRepeat
    gr_None = 0
    gr_Wrap = 1
    gr_Reflect = 2
End Enum

#If False Then
    Private Const gr_None = 0, gr_Wrap = 1, gr_Reflect = 2
#End If

'Gradient attributes are stored in these variables
Private m_GradientOpacity As Single
Private m_GradientBlendmode As PD_BlendMode
Private m_GradientAlphamode As PD_AlphaMode
Private m_GradientAntialiasing As PD_2D_Antialiasing
Private m_GradientRepeat As PD_GradientRepeat

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
Private m_GradientGdip As pd2DGradient, m_GradientCairo As pd2DGradientCairo

'Other gradient parameters, as relevant
Private m_Angle As Single

'TESTING ONLY: components for a fast-preview mode
Private m_PreviewDIB As pdDIB

'Universal gradient settings
Public Function GetGradientAlphaMode() As PD_AlphaMode
    GetGradientAlphaMode = m_GradientAlphamode
End Function

Public Function GetGradientAntialiasing() As PD_2D_Antialiasing
    GetGradientAntialiasing = m_GradientAntialiasing
End Function

Public Function GetGradientBlendMode() As PD_BlendMode
    GetGradientBlendMode = m_GradientBlendmode
End Function

Public Function GetGradientOpacity() As Single
    GetGradientOpacity = m_GradientOpacity
End Function

Public Function GetGradientRepeat() As PD_GradientRepeat
    GetGradientRepeat = m_GradientRepeat
End Function

'Property set functions.  Note that not all brush properties are used by all styles.
' (e.g. "brush hardness" is not used by "pencil" style brushes, etc)
Public Sub SetGradientAlphaMode(Optional ByVal newAlphaMode As PD_AlphaMode = LA_NORMAL)
    If (newAlphaMode <> m_GradientAlphamode) Then m_GradientAlphamode = newAlphaMode
End Sub

Public Sub SetGradientAntialiasing(Optional ByVal newAntialiasing As PD_2D_Antialiasing = P2_AA_HighQuality)
    If (newAntialiasing <> m_GradientAntialiasing) Then m_GradientAntialiasing = newAntialiasing
End Sub

Public Sub SetGradientBlendMode(Optional ByVal newBlendMode As PD_BlendMode = BL_NORMAL)
    If (newBlendMode <> m_GradientBlendmode) Then m_GradientBlendmode = newBlendMode
End Sub

Public Sub SetGradientOpacity(ByVal newOpacity As Single)
    If (newOpacity <> m_GradientOpacity) Then m_GradientOpacity = newOpacity
End Sub

Public Sub SetGradientRepeat(ByVal newRepeat As PD_GradientRepeat)
    If (newRepeat <> m_GradientRepeat) Then m_GradientRepeat = newRepeat
End Sub

Public Function GetGradientProperty(ByVal bProperty As PD_GradientAttributes) As Variant
    
    Select Case bProperty
        Case GA_AlphaMode
            GetGradientProperty = GetGradientAlphaMode()
        Case GA_Antialiasing
            GetGradientProperty = GetGradientAntialiasing()
        Case GA_BlendMode
            GetGradientProperty = GetGradientBlendMode()
        Case GA_Opacity
            GetGradientProperty = GetGradientOpacity()
        Case GA_Repeat
            GetGradientProperty = GetGradientRepeat()
    End Select
    
End Function

Public Sub SetBrushProperty(ByVal bProperty As PD_BrushAttributes, ByVal newPropValue As Variant)
    
    Select Case bProperty
        Case GA_AlphaMode
            SetGradientAlphaMode newPropValue
        Case GA_Antialiasing
            SetGradientAntialiasing newPropValue
        Case GA_BlendMode
            SetGradientBlendMode newPropValue
        Case GA_Opacity
            SetGradientOpacity newPropValue
        Case GA_Repeat
            SetGradientRepeat newPropValue
    End Select
    
End Sub

'Notify the gradient engine of the current mouse position.  Coordinates should always be in *image* coordinate space,
' not screen space.  (Translation between spaces will be handled internally.)
Public Sub NotifyToolXY(ByVal mouseButtonDown As Boolean, ByVal Shift As ShiftConstants, ByVal srcX As Single, ByVal srcY As Single, ByVal mouseTimeStamp As Long, ByRef srcCanvas As pdCanvas)
    
    If (Not PDImages.IsImageActive()) Then Exit Sub
    
    m_MouseX = srcX
    m_MouseY = srcY
    
    Dim isFirstStroke As Boolean, isLastStroke As Boolean
    isFirstStroke = (Not m_MouseDown) And mouseButtonDown
    isLastStroke = m_MouseDown And (Not mouseButtonDown)
    
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
        
        If USE_CAIRO_RENDERER Then
        
            Set m_GradientCairo = New pd2DGradientCairo
            m_GradientCairo.CreateGradientFromGdipGradientString toolpanel_Gradient.grdPrimary.Gradient()
            m_GradientCairo.SetGradientShape P2_GS_Linear
            
            Select Case m_GradientRepeat
                Case gr_None
                    m_GradientCairo.SetGradientExtend ce_ExtendPad
                Case gr_Wrap
                    m_GradientCairo.SetGradientExtend ce_ExtendRepeat
                Case gr_Reflect
                    m_GradientCairo.SetGradientExtend ce_ExtendReflect
            End Select
            
        Else
            Set m_GradientGdip = New pd2DGradient
            m_GradientGdip.CreateGradientFromString toolpanel_Gradient.grdPrimary.Gradient()
            m_GradientGdip.SetGradientShape P2_GS_Linear
            
            Select Case m_GradientRepeat
                Case gr_None
                    'Clamp mode is not supported by GDI+, so we lie and set a functional mode
                    ' and simply overwrite the results later
                    'm_GradientGdip.SetGradientWrapMode P2_WM_Clamp
                    m_GradientGdip.SetGradientWrapMode P2_WM_TileFlipXY
                Case gr_Wrap
                    m_GradientGdip.SetGradientWrapMode P2_WM_Tile
                Case gr_Reflect
                    m_GradientGdip.SetGradientWrapMode P2_WM_TileFlipXY
            End Select
            
        End If
        
    End If
    
    'On any other stroke, update the 2nd set of mouse coordinates
    If m_PointsInitialized Then
        m_Points(1).x = srcX
        m_Points(1).y = srcY
    End If
    
    'Notify the scratch layer of our updates
    If mouseButtonDown Or isLastStroke Then
        
        PDImages.GetActiveImage.ScratchLayer.layerDIB.ResetDIB 0
        
        If USE_FAST_PREVIEW And (Not isLastStroke) Then
            PreviewRenderer srcCanvas
        Else
            If USE_CAIRO_RENDERER Then
                CairoRenderer
            Else
                GdipRenderer
            End If
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
    tmpViewportParams = ViewportEngine.GetDefaultParamObject()
    tmpViewportParams.renderScratchLayerIndex = PDImages.GetActiveImage.GetActiveLayerIndex()
    If USE_FAST_PREVIEW And (Not isLastStroke) Then tmpViewportParams.ptrToAlternateScratch = ObjPtr(m_PreviewDIB)
    If mouseButtonDown Then ViewportEngine.Stage2_CompositeAllLayers PDImages.GetActiveImage(), srcCanvas, VarPtr(tmpViewportParams)
    
End Sub

'A new test; attempt to maximize performance by translating the gradient to the current viewport space and only rendering it there.
' At _MouseUp(), a full-size preview will be manually rendered and committed.
Private Sub PreviewRenderer(ByRef srcCanvas As pdCanvas)
    
    'Retrieve a copy of the intersected viewport rect; we will use this for clipping
    Dim viewportIntersectRect As RectF
    PDImages.GetActiveImage.ImgViewport.GetIntersectRectCanvas viewportIntersectRect
    
    'Ensure we have a valid preview DIB (TODO: don't re-initialize on every call, obviously)
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
    
    'Test only: fill the selected area
    Dim cSurface As pd2DSurface
    Set cSurface = New pd2DSurface
    cSurface.WrapSurfaceAroundPDDIB m_PreviewDIB
    
    Dim cBrush As pd2DBrush
    Drawing2D.QuickCreateSolidBrush cBrush, vbRed, 50!
    
    PD2D.FillRectangleF cSurface, cBrush, 0, 0, viewportIntersectRect.Width, viewportIntersectRect.Height
    
    Set cBrush = Nothing
    Set cSurface = Nothing
    
End Sub

Private Sub CairoRenderer()

    'Rendering methods are still being debated; cairo and GDI+ both have trade-offs depending
    ' on gradient parameters.
    Dim cSurface As pd2DSurfaceCairo
    Set cSurface = New pd2DSurfaceCairo
    cSurface.SetAntialias ca_FAST
    cSurface.SetOperator co_Source
    cSurface.WrapAroundPDDIB PDImages.GetActiveImage.ScratchLayer.layerDIB
    
    'Populate any remaining gradient properties
    m_GradientCairo.SetGradientPoint1 m_Points(0)
    m_GradientCairo.SetGradientPoint2 m_Points(1)
    m_GradientCairo.SetGradientShape P2_GS_Linear
    'm_GradientCairo.SetGradientRadii 0!, PDMath.DistanceTwoPoints(m_Points(0).x, m_Points(0).y, m_Points(1).x, m_Points(1).y)
    
    'Select the pattern into the destination source
    Dim hPattern As Long
    hPattern = m_GradientCairo.GetPatternHandle()
    Plugin_Cairo.Context_SetSourcePattern cSurface.GetContextHandle, hPattern
    
    Dim cairoStartTime As Currency
    VBHacks.GetHighResTime cairoStartTime
    
    'Fill the entire source
    Plugin_Cairo.Context_Rectangle cSurface.GetContextHandle, 0#, 0#, PDImages.GetActiveImage.ScratchLayer.layerDIB.GetDIBWidth, PDImages.GetActiveImage.ScratchLayer.layerDIB.GetDIBHeight
    Plugin_Cairo.Context_Fill cSurface.GetContextHandle
    
    Debug.Print VBHacks.GetTimeDiffNowAsString(cairoStartTime)
    
    'Free all handles and notify the scratch layer of our changes
    Plugin_Cairo.FreeCairoPattern hPattern
    Set cSurface = Nothing
    
End Sub

Private Sub GdipRenderer()

    'Rendering methods are still being debated; cairo and GDI+ both have trade-offs depending
    ' on gradient parameters.
    Dim cSurface As pd2DSurface
    Set cSurface = New pd2DSurface
    cSurface.WrapSurfaceAroundPDDIB PDImages.GetActiveImage.ScratchLayer.layerDIB
    cSurface.SetSurfaceAntialiasing P2_AA_None
    cSurface.SetSurfaceCompositing P2_CM_Overwrite
    cSurface.SetSurfacePixelOffset P2_PO_Normal
    
    'Populate any remaining gradient properties
    m_GradientGdip.SetGradientShape P2_GS_Linear
    
    Dim gradAngle As Double
    gradAngle = PDMath.Atan2(m_Points(1).y - m_Points(0).y, m_Points(1).x - m_Points(0).x)
    m_GradientGdip.SetGradientAngle PDMath.RadiansToDegrees(gradAngle)
    
    Dim gdipStartTime As Currency
    VBHacks.GetHighResTime gdipStartTime
    
    'Fill the entire source
    Dim cBrush As pd2DBrush
    Set cBrush = New pd2DBrush
    cBrush.SetBrushMode P2_BM_Gradient
    
    Dim cRadius As Double
    cRadius = PDMath.DistanceTwoPoints(m_Points(0).x, m_Points(0).y, m_Points(1).x, m_Points(1).y)
    
    Dim boundsRect As RectF
    With boundsRect
        .Left = PDMath.Min2Float_Single(m_Points(0).x, m_Points(1).x)
        .Top = PDMath.Min2Float_Single(m_Points(0).y, m_Points(1).y)
        .Width = Abs(m_Points(1).x - m_Points(0).x)
        If (.Width < 1!) Then .Width = 1!
        .Height = Abs(m_Points(1).y - m_Points(0).y)
        If (.Height < 1!) Then .Height = 1!
    End With
    
    'We now have everything we need to render the gradient.  Unfortunately, certain edge-wrap modes
    ' (e.g. clamp) have no direct support in GDI+.  This greatly complicates their rendering, as we
    ' must manually clamp the results.  (Similarly, GDI+ does not allow you to render with any kind
    ' of "non-tiled/wrapped" behavior, so we must sometimes overwrite the result manually.)
    cBrush.SetBoundaryRect boundsRect
    cBrush.SetBrushGradientAllSettings m_GradientGdip.GetGradientAsString()
    
    Dim slWidth As Single, slHeight As Single
    slWidth = PDImages.GetActiveImage.ScratchLayer.layerDIB.GetDIBWidth()
    slHeight = PDImages.GetActiveImage.ScratchLayer.layerDIB.GetDIBHeight()
    PD2D.FillRectangleF cSurface, cBrush, 0!, 0!, slWidth, slHeight
    
    'The gradient now covers the entire underlying scratch layer, for better or worse.
    
    'If the wrap mode is a mode unsupported by GDI+ (e.g. "extend/clamp"), we now need to manually
    ' overwrite the gradient in certain areas.
    If (m_GradientRepeat = gr_None) Then
    
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
        PDMath.ConvertPolarToCartesian_Sng angPerpendicular, diagLength, clipPoly(0).x, clipPoly(0).y, m_Points(0).x, m_Points(0).y
        PDMath.ConvertPolarToCartesian_Sng angPerpendicular2, diagLength, clipPoly(1).x, clipPoly(1).y, m_Points(0).x, m_Points(0).y
        
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
        Drawing2D.QuickCreateSolidBrush cBrush, RGB(srcColor.Red, srcColor.Green, srcColor.Blue), srcColor.Alpha / 2.55!
        PD2D.FillPolygonF_FromPtF cSurface, cBrush, 4, VarPtr(clipPoly(0))
        
        'Now we basically repeat all the above steps, but for the second gradient endpoint.
        ' (Naturally, we also swap the order of +/-90 points, to ensure that the polygon lies on the
        ' opposite side of the gradient.)
        angPerpendicular = gradAngle - PI_HALF
        angPerpendicular2 = gradAngle + PI_HALF
        PDMath.ConvertPolarToCartesian_Sng angPerpendicular, diagLength, clipPoly(0).x, clipPoly(0).y, m_Points(1).x, m_Points(1).y
        PDMath.ConvertPolarToCartesian_Sng angPerpendicular2, diagLength, clipPoly(1).x, clipPoly(1).y, m_Points(1).x, m_Points(1).y
        
        angPerpendicular = angPerpendicular + PI_HALF
        angPerpendicular2 = angPerpendicular2 - PI_HALF
        PDMath.ConvertPolarToCartesian_Sng angPerpendicular, diagLength, clipPoly(3).x, clipPoly(3).y, clipPoly(0).x, clipPoly(0).y
        PDMath.ConvertPolarToCartesian_Sng angPerpendicular2, diagLength, clipPoly(2).x, clipPoly(2).y, clipPoly(1).x, clipPoly(1).y
        
        'Fill the new clip area with the last color in the gradient
        m_GradientGdip.GetColorAtPosition_RGBA 1!, srcColor
        Drawing2D.QuickCreateSolidBrush cBrush, RGB(srcColor.Red, srcColor.Green, srcColor.Blue), srcColor.Alpha / 2.55!
        PD2D.FillPolygonF_FromPtF cSurface, cBrush, 4, VarPtr(clipPoly(0))
        
        'Free our intermediary "fix" brush
        Set cBrush = Nothing
        
    End If
    
    'Profiling will be turned off in final builds, obviously
    Debug.Print VBHacks.GetTimeDiffNowAsString(gdipStartTime)
    
    'Free all handles
    Set cBrush = Nothing
    Set cSurface = Nothing
    
End Sub

'Want to commit your current gradient work?  Call this function to make the gradient results permanent.
Public Sub CommitGradientResults()
    
    'If relevant, make a local copy of the gradient's bounding rect, and clip it to the layer's boundaries
    'Dim tmpRectF As RectF
    'tmpRectF = m_TotalModifiedRectF
   '
   ' With tmpRectF
   '     If (.Left < 0) Then .Left = 0
   '     If (.Top < 0) Then .Top = 0
   '     If (.Width > PDImages.GetActiveImage.ScratchLayer.layerDIB.GetDIBWidth) Then .Width = PDImages.GetActiveImage.ScratchLayer.layerDIB.GetDIBWidth
   '     If (.Height > PDImages.GetActiveImage.ScratchLayer.layerDIB.GetDIBHeight) Then .Height = PDImages.GetActiveImage.ScratchLayer.layerDIB.GetDIBHeight
   ' End With
    
    'Committing gradient results is actually pretty easy!
    
    'First, if the layer beneath the gradient is a raster layer, we simply want to merge the scratch
    ' layer onto it.
    If PDImages.GetActiveImage.GetActiveLayer.IsLayerRaster Then
        
        Dim bottomLayerFullSize As Boolean
        With PDImages.GetActiveImage.GetActiveLayer
            bottomLayerFullSize = ((.GetLayerOffsetX = 0) And (.GetLayerOffsetY = 0) And (.layerDIB.GetDIBWidth = PDImages.GetActiveImage.Width) And (.layerDIB.GetDIBHeight = PDImages.GetActiveImage.Height))
        End With
        
        PDImages.GetActiveImage.MergeTwoLayers PDImages.GetActiveImage.ScratchLayer, PDImages.GetActiveImage.GetActiveLayer, bottomLayerFullSize, True  ', VarPtr(tmpRectF)
        PDImages.GetActiveImage.NotifyImageChanged UNDO_Layer, PDImages.GetActiveImage.GetActiveLayerIndex
        
        'Ask the central processor to create Undo/Redo data for us
        Processor.Process "Gradient tool", , , UNDO_Layer, g_CurrentTool
        
        'Reset the scratch layer
        PDImages.GetActiveImage.ScratchLayer.layerDIB.ResetDIB 0
    
    'If the layer beneath this one is *not* a raster layer, let's add the gradient as a new layer, instead.
    Else
        
        'Before creating the new layer, check for an active selection.  If one exists, we need to preprocess
        ' the paint layer against it.
        If PDImages.GetActiveImage.IsSelectionActive Then
            
            'A selection is active.  Pre-mask the paint scratch layer against it.
            Dim cBlender As pdPixelBlender
            Set cBlender = New pdPixelBlender
            cBlender.ApplyMaskToTopDIB PDImages.GetActiveImage.ScratchLayer.layerDIB, PDImages.GetActiveImage.MainSelection.GetMaskDIB  ', VarPtr(tmpRectF)
            
        End If
        
        Dim newLayerID As Long
        newLayerID = PDImages.GetActiveImage.CreateBlankLayer(PDImages.GetActiveImage.GetActiveLayerIndex)
        
        'Point the new layer index at our scratch layer
        PDImages.GetActiveImage.PointLayerAtNewObject newLayerID, PDImages.GetActiveImage.ScratchLayer
        PDImages.GetActiveImage.GetLayerByID(newLayerID).SetLayerName g_Language.TranslateMessage("Gradient layer")
        Set PDImages.GetActiveImage.ScratchLayer = Nothing
        
        'Activate the new layer
        PDImages.GetActiveImage.SetActiveLayerByID newLayerID
        
        'Notify the parent image of the new layer
        PDImages.GetActiveImage.NotifyImageChanged UNDO_Image_VectorSafe
        
        'Redraw the layer box, and note that thumbnails need to be re-cached
        toolbar_Layers.NotifyLayerChange
        
        'Ask the central processor to create Undo/Redo data for us
        Processor.Process "Gradient tool", , , UNDO_Image_VectorSafe, g_CurrentTool
        
        'Create a new scratch layer
        Tools.InitializeToolsDependentOnImage
        
    End If
    
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
