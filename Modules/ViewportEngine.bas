Attribute VB_Name = "Viewport"
'***************************************************************************
'Viewport Handler - builds and draws the image viewport and associated scroll bars
'Copyright 2001-2026 by Tanner Helland
'Created: 4/15/01
'Last updated: 20/November/24
'Last update: add switch for crop tool UI rendering
'
'Module for handling the image viewport.  The render pipeline works as follows:
'
' - Viewport.Stage1_InitializeBuffer: calculate all viewport position and overlay rects
'    (required only on first image load, when image size is changed, or when viewport zoom changes)
' - Viewport.Stage2_CompositeAllLayers: composite all layers in the active image, while accounting for
'    things like non-destructive modifications.  Because this function only composites the subset of
'    the image relevant to the target viewport, it must be called on canvas scrollbar changes.
' - Viewport.Stage3_CompositeCanvas: composite the image with any underlying/overlying canvas UI elements.
'    At present, this stage handles color management and selection tool compositing, when active.
' - Viewport.Stage4_FlipBufferAndDrawUI: composite any interactive UI elements (transform nodes,
'    paint tool outlines, etc) to the canvas, then flip everything to the screen.
'
'If you need to draw something to the screen, you need to call the *latest possible pipeline stage*.
' Stages are sorted in rough order of their "processing time" requirements, and unnecessarily calling
' early pipeline stages will harm viewport performance.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Due to the complexity of viewport rendering, some viewport pipeline calculations are stored at module-level,
' while others are stored inside the source pdImage or destination pdCanvas used for a given rendering.
' Module-level variables are typically used to improve efficiency over passing objects as parameters,
' because PD's viewport pipeline is as an "out of order" pipeline. Different user interations trigger execution
' of the pipeline at different stages, which is crucial for maximizing viewport performance.
'
'As such, pipeline functions must be *very cautious* about modifying module-level values, or viewport-related
' values stored within source pdImage or destination pdCanvas objects.  CONSIDER YOURSELF WARNED.

'If external functions require special scroll bar treatment in the initial pipeline stage, they must pass one
' of these enums as the first entry in the associated ParamArray().
Public Enum PD_VIEWPORT_SPECIAL_REQUEST
    VSR_ResetToZero = 0
    VSR_ResetToCustom = 1
    VSR_PreservePointPosition = 2
End Enum

#If False Then
    Private Const VSR_ResetToZero = 0, VSR_ResetToCustom = 1, VSR_PreservePointPosition = 2
#End If

'The ZoomVal value is the actual coefficient for the current zoom value (i.e. 0.50 for "50% zoom").
' Multiple pipeline stages use this, so it's cached by the first pipeline staged and simply reused after that.
Private m_ZoomRatio As Double

'm_FrontBuffer holds the final composited image, including any non-interactive overlays
' (like selection tool effects)
Private m_FrontBuffer As pdDIB

'In most cases, viewport rendering is automatically triggered as underlying actions require it,
' but if a bunch of requests need to be batched, it's useful to forcibly delay automatic redraws.
' This variable tracks forcible viewport suspensions; interact with it via the safe Enable/Disable
' wrapper functions, below.
Private m_DisableViewportRendering As Boolean

'As part of continued viewport optimizations, we track the amount of time spent in each viewport stage.
' Note that stage 1 is ignored because it is only called under specific circumstances that are very
' difficult to profile accurately (e.g. changes in zoom values or switching between images).
Private m_TimeStage2 As Currency, m_TimeStage3 As Currency, m_TimeStage4 As Currency
Private m_TotalTime As Currency, m_TotalTimeStage2 As Currency, m_TotalTimeStage3 As Currency, m_TotalTimeStage4 As Currency

'The last POI ("point of interest") passed to this class.  When a marching ant selection outline is active,
' it operates on a timer, sending us redraw requests every (n) ms.  Rather than ask *it* to remember any
' active POIs (e.g. interactive bits on the active canvas that need to be drawn differently), we cache the
' last POI we received and automatically render it on marching-ant timer-initiated redraws.
Private m_LastPOI As PD_PointOfInterest

'The viewport engine supports a number of optional parameters.  Rather than exponentially increase the parameter list
' of each function, new viewport parameters are added to this struct.  Calls into the viewport engine do *not* have
' to supply this struct; most calls will simply want to use default parameters.
Public Type PD_ViewportParams
    curPOI As PD_PointOfInterest
    renderScratchLayerIndex As Long
    ptrToAlternateScratch As Long
End Type

Public Function GetDefaultParamObject() As PD_ViewportParams
    GetDefaultParamObject.curPOI = poi_Undefined
    GetDefaultParamObject.renderScratchLayerIndex = -1
    GetDefaultParamObject.ptrToAlternateScratch = 0&
End Function

'Stage4_FlipBufferAndDrawUI is the final stage of the viewport pipeline.  It will flip the composited
' canvas image to the destination pdCanvas object, and apply any final UI elements as well - control nodes,
' custom cursors, etc.  This step is very fast, and should be used whenever full compositing is unnecessary.
'
'Note also that this stage is the only one to make use of the optional POI ID parameter.  If supplied,
' it will forward this to any UI functions that accept POI identifiers.  (Because the viewport is agnostic
' to underlying UI complexities, by design, it is up to the caller to optimize POI-based requests,
' e.g. not forwarding them unless the POI has changed, etc.)
Public Sub Stage4_FlipBufferAndDrawUI(ByRef srcImage As pdImage, ByRef dstCanvas As pdCanvas, Optional ByVal ptrToViewportParams As Long = 0, Optional ByVal fullPipelineCall As Boolean = False)
    
    'If an outside function has invoked this pipeline stage directly, we need to make sure we're even allowed to render
    Dim allowedToRender As Boolean: allowedToRender = True
    If (Not fullPipelineCall) Then allowedToRender = ViewportRenderingAllowed(srcImage, dstCanvas)
    
    If allowedToRender Then
    
        'If no images have been loaded, clear the canvas and exit
        If (Not PDImages.IsImageNonNull()) Then
            dstCanvas.ClearCanvas
        Else
            
            Dim startTime As Currency
            VBHacks.GetHighResTime startTime
            
            'If the user passed in viewport parameters, retrieve them now
            Dim localViewportParams As PD_ViewportParams
            If (ptrToViewportParams <> 0) Then
                CopyMemoryStrict VarPtr(localViewportParams), ptrToViewportParams, LenB(localViewportParams)
            Else
                localViewportParams = Viewport.GetDefaultParamObject()
            End If
            
            'If the "reuse last POI" indicator is passed, substitute our last-saved POI for the one passed in
            If (localViewportParams.curPOI = poi_ReuseLast) Then localViewportParams.curPOI = m_LastPOI Else m_LastPOI = localViewportParams.curPOI
            
            'Flip the viewport buffer over to the canvas control.  Any additional rendering must now happen there.
            GDI.BitBltWrapper dstCanvas.hDC, 0, 0, m_FrontBuffer.GetDIBWidth, m_FrontBuffer.GetDIBHeight, m_FrontBuffer.GetDIBDC, 0, 0, vbSrcCopy
            
            'Some canvas UI options are universal, rendered atop *any* tool
            If Drawing.Get_ShowLayerEdges() Then Drawing.DrawLayerBoundaries dstCanvas, srcImage, srcImage.GetActiveLayer
            
            'Lastly, do any tool-specific rendering directly onto the canvas itself.
            
            'Various tools do their own custom UI rendering atop the canvas
            If (g_CurrentTool = NAV_MOVE) Then
                Tools_Move.DrawCanvasUI dstCanvas, srcImage, localViewportParams.curPOI, False
                
            ElseIf (g_CurrentTool = NAV_ZOOM) Then
                Tools_Zoom.DrawCanvasUI dstCanvas, srcImage
            
            ElseIf (g_CurrentTool = COLOR_PICKER) Then
                If FormMain.MainCanvas(0).IsMouseOverCanvas Then Tools_ColorPicker.RenderColorPickerCursor dstCanvas
            
            ElseIf (g_CurrentTool = ND_MEASURE) Then
                Tools_Measure.RenderMeasureUI dstCanvas
                
            ElseIf (g_CurrentTool = ND_CROP) Then
                Tools_Crop.DrawCanvasUI dstCanvas, srcImage
                
            'Selections are always rendered onto the canvas.  If a selection is active AND a selection tool is active, we can also
            ' draw transform nodes around the selection area.  (Note that lasso selections are currently an exception to this rule;
            ' they only support the "move" interaction, which is applied by click-dragging anywhere in the lasso region.)
            ElseIf (g_CurrentTool = SELECT_RECT) Or (g_CurrentTool = SELECT_CIRC) Or (g_CurrentTool = SELECT_POLYGON) Or (g_CurrentTool = SELECT_WAND) Or (g_CurrentTool = SELECT_LASSO) Then
                If srcImage.IsSelectionActive Then srcImage.MainSelection.RenderTransformNodes srcImage, dstCanvas, g_CurrentTool
                    
            'Text tools currently draw layer boundaries at all times; I'm working on letting the user control this (TODO!)
            ElseIf (g_CurrentTool = TEXT_BASIC) Or (g_CurrentTool = TEXT_ADVANCED) Then
                If PDImages.GetActiveImage.GetActiveLayer.IsLayerText Then Tools_Move.DrawCanvasUI dstCanvas, srcImage, localViewportParams.curPOI, True
                
            'Pencil and brush tools use the brush engine to paint a custom brush outline at the current mouse position
            ElseIf (g_CurrentTool = PAINT_PENCIL) Then
                If FormMain.MainCanvas(0).IsMouseOverCanvas Then Tools_Pencil.RenderBrushOutline dstCanvas
                
            ElseIf (g_CurrentTool = PAINT_SOFTBRUSH) Or (g_CurrentTool = PAINT_ERASER) Then
                If FormMain.MainCanvas(0).IsMouseOverCanvas Then Tools_Paint.RenderBrushOutline dstCanvas
            
            ElseIf (g_CurrentTool = PAINT_CLONE) Then
                If FormMain.MainCanvas(0).IsMouseOverCanvas Then Tools_Clone.RenderBrushOutline dstCanvas
            
            'Fill tools also render a custom cursor
            ElseIf (g_CurrentTool = PAINT_FILL) Then
                If FormMain.MainCanvas(0).IsMouseOverCanvas Then Tools_Fill.RenderFillCursor dstCanvas
                
            ElseIf (g_CurrentTool = PAINT_GRADIENT) Then
                If FormMain.MainCanvas(0).IsMouseOverCanvas Then Tools_Gradient.RenderGradientUI dstCanvas
                
            End If
            
            'With all rendering complete, we are finally ready to request a screen refresh
            dstCanvas.RequestViewportRedraw True
            
            'Before exiting, calculate the time spent in this stage
            m_TimeStage4 = VBHacks.GetTimerDifferenceNow(startTime)
            If fullPipelineCall Then m_TotalTimeStage4 = m_TotalTimeStage4 + m_TimeStage4
            
        End If
    
    End If
    
End Sub

'Stage3_CompositeCanvas takes the current canvas (which has a checkerboard and fully layered image drawn atop it) and applies a few
' other frills to it, including things like canvas decorations (e.g. drop-shadows, a fix for the space between scroll bars), and the
' current selection, if one is active.  This stage is the final stage before color-management is applied, so it's important to render
' any color-specific bits now, as the next stage will apply color-management processing to whatever is contained in the front buffer.
'
'When this stage is finished, the srcImage.m_FrontBuffer object will contain a screen-ready copy of the canvas, with the fully
' composited image drawn atop a checkerboard in the viewport section of the canvas.  Standard canvas decorations will also be present,
' provided that the user's performance settings allow them.
'
'After this stage, the only things that should be rendered onto the canvas are uncolored UI elements, like custom-drawn cursors or
' control nodes.
Public Sub Stage3_CompositeCanvas(ByRef srcImage As pdImage, ByRef dstCanvas As pdCanvas, Optional ByVal ptrToViewportParams As Long = 0, Optional ByVal fullPipelineCall As Boolean = False)
    
    'If an outside function has invoked this pipeline stage directly, we need to make sure we're even allowed to render
    Dim allowedToRender As Boolean: allowedToRender = True
    If (Not fullPipelineCall) Then allowedToRender = ViewportRenderingAllowed(srcImage, dstCanvas)
    
    If allowedToRender Then
        
        Dim startTime As Currency
        VBHacks.GetHighResTime startTime
            
        'If the user passed in viewport parameters, retrieve them now
        Dim localViewportParams As PD_ViewportParams
        If (ptrToViewportParams <> 0) Then
            CopyMemoryStrict VarPtr(localViewportParams), ptrToViewportParams, LenB(localViewportParams)
        Else
            localViewportParams = Viewport.GetDefaultParamObject()
        End If
        
        'If no images have been loaded, clear the canvas and exit
        If (Not PDImages.IsImageNonNull()) Then
            
            dstCanvas.ClearCanvas
            
            'Before exiting, calculate the time spent in this stage
            m_TimeStage3 = VBHacks.GetTimerDifferenceNow(startTime)
            If fullPipelineCall Then m_TotalTimeStage3 = m_TotalTimeStage3 + m_TimeStage3
            
        Else
            
            'Create the front buffer as necessary
            If (m_FrontBuffer Is Nothing) Then Set m_FrontBuffer = New pdDIB
            If (m_FrontBuffer.GetDIBWidth <> srcImage.CanvasBuffer.GetDIBWidth) Or (m_FrontBuffer.GetDIBHeight <> srcImage.CanvasBuffer.GetDIBHeight) Then
                m_FrontBuffer.CreateFromExistingDIB srcImage.CanvasBuffer
            Else
                GDI.BitBltWrapper m_FrontBuffer.GetDIBDC, 0, 0, srcImage.CanvasBuffer.GetDIBWidth, srcImage.CanvasBuffer.GetDIBHeight, srcImage.CanvasBuffer.GetDIBDC, 0, 0, vbSrcCopy
            End If
            
            'Retrieve a copy of the intersected viewport rect; subsequent rendering ops may use this to optimize their operations
            Dim viewportIntersectRect As RectF
            srcImage.ImgViewport.GetIntersectRectCanvas viewportIntersectRect
            
            '*Now* is when we want to apply color management to the front buffer.  (For performance reasons, UI elements drawn atop
            ' the canvas are not color-managed - only the image itself is.)  Note also that although the front buffer is 32-bpp,
            ' it is always fully opaque, so we notify the color management engine that alpha bytes can be safely ignored.
            ColorManagement.ApplyDisplayColorManagement_RectF m_FrontBuffer, viewportIntersectRect, , False
            
            'Check to see if a selection is active.  If it is, we want to render it now, directly atop the front buffer.  This allows any
            ' subsequent overlays (e.g. brush outlines) to appear "on top" of the selection, without us needing to redraw the selection outline
            ' on every overlay render.
            If srcImage.IsSelectionActive Then srcImage.MainSelection.RenderSelectionToViewport m_FrontBuffer, srcImage, dstCanvas
            
            'Before exiting, calculate the time spent in this stage
            m_TimeStage3 = VBHacks.GetTimerDifferenceNow(startTime)
            If fullPipelineCall Then m_TotalTimeStage3 = m_TotalTimeStage3 + m_TimeStage3
            
            'Pass the completed front buffer to the final stage of the pipeline, which will flip everything to the screen and render any
            ' remaining UI elements!
            Stage4_FlipBufferAndDrawUI srcImage, dstCanvas, ptrToViewportParams, True
            
        End If
        
    End If
            
End Sub

'Stage2_CompositeAllLayers is used to composite a viewport-specific representation of the current layer stack.  The composited
' image is then placed in the source pdImage's back buffer.  Note that this function does not perform any initialization or
' pre-rendering checks, so you cannot use it if zoom is changed, or if the viewport area has changed due to a main window resize.
' (When that happens, you need to call Stage1_InitializeBuffer().)
'
'This function should be called whenever changes are made to individual layers - e.g. from paint tools, adjustments, effects, etc -
' or when viewport scrollbars are used.
'
'The optional fullPipelineCall parameter lets this function know if its been called by a previous pipeline stage.  If it has,
' a full viewport cache purge is required (because things like zoom or window size may have changed).  If this function is called
' directly by another portion of the program, existing caches can be safely reused - but the function *must* check to make sure
' viewport rendering is allowed (as it can't assume a parent pipeline stage has performed this check on its behalf).
Public Sub Stage2_CompositeAllLayers(ByRef srcImage As pdImage, ByRef dstCanvas As pdCanvas, Optional ByVal ptrToViewportParams As Long = 0, Optional ByVal fullPipelineCall As Boolean = False)
    
    'If an outside function has invoked this pipeline stage directly, we need to make sure we're even allowed to render
    Dim allowedToRender As Boolean: allowedToRender = True
    If (Not fullPipelineCall) Then allowedToRender = ViewportRenderingAllowed(srcImage, dstCanvas)
    
    If allowedToRender Then
        
        'If the user passed in viewport parameters, retrieve them now
        Dim localViewportParams As PD_ViewportParams
        If (ptrToViewportParams <> 0) Then
            CopyMemoryStrict VarPtr(localViewportParams), ptrToViewportParams, LenB(localViewportParams)
        Else
            localViewportParams = Viewport.GetDefaultParamObject()
        End If
        
        'This function can return timing reports if desired; at present, this is automatically activated in PRE-ALPHA and ALPHA builds,
        ' but disabled for BETA and PRODUCTION builds; see the LoadTheProgram() function for details.
        Dim startTime As Currency
        VBHacks.GetHighResTime startTime
        
        'Stage 1 of the pipeline (Stage1_InitializeBuffer) prepared srcImage.compositeBuffer for us.  The goal of this stage
        ' is simple: fill the compositeBuffer object with a fully composited copy of the current image, cropped and zoomed to
        ' match the target viewport settings.
        
        'Regardless of the pipeline branch we follow, we need local copies of the relevant region rects calculated by stage 1 of the pipeline.
        Dim imageRect_CanvasCoords As RectF, canvasRect_ImageCoords As RectF, canvasRect_ActualPixels As RectF
        With srcImage.ImgViewport
            .GetCanvasRectActualPixels canvasRect_ActualPixels
            .GetCanvasRectImageCoords canvasRect_ImageCoords
            .GetImageRectCanvasCoords imageRect_CanvasCoords
        End With
        
        'We also need to wipe the back buffer
        GDI_Plus.GDIPlusFillDIBRect srcImage.CanvasBuffer, 0, 0, srcImage.CanvasBuffer.GetDIBWidth, srcImage.CanvasBuffer.GetDIBHeight, UserPrefs.GetCanvasColor(), 255, GP_CM_SourceCopy
        
        'Stage 1 of the pipeline (Stage1_InitializeBuffer) prepared srcImage.BackBuffer for us.  If the user's preferences are "BEST QUALITY",
        ' Stage 2 composited a full-sized version of the image.  The goal of this stage (3) is two-fold:
        ' 1) Fill the viewport area of the canvas with a checkerboard pattern
        ' 2) Render the fully composited image atop the checkerboard pattern
        
        'If the user is not using "BEST QUALITY", the imgCompositor class will be used to dynamically render only the portion of the image
        ' relevant for the current viewport.
        
        'The first thing we need to do is find the intersection rect between two things: the source image, and the canvas rect,
        ' in both the image and canvas coordinate spaces.  These are used to construct a StretchBlt-like set of (x, y) and
        ' (width, height) pairs, which the compositor uses to snip out a portion of the composited image.
        
        'Because the original function doesn't deal with scroll bar values at all, let's calculate the offsets the scroll bars apply.
        Dim xScroll_Canvas As Single, xScroll_Image As Single, yScroll_Canvas As Single, yScroll_Image As Single
        
        'Scroll bar values always represent pixel measurements *in the image coordinate space*.
        xScroll_Image = dstCanvas.GetScrollValue(pdo_Horizontal)
        yScroll_Image = dstCanvas.GetScrollValue(pdo_Vertical)
        
        'Next, let's calculate these *in the canvas coordinate space* (e.g. with zoom applied)
        If (m_ZoomRatio = 0#) Then m_ZoomRatio = Zoom.GetZoomRatioFromIndex(srcImage.GetZoomIndex())
        xScroll_Canvas = xScroll_Image * m_ZoomRatio
        yScroll_Canvas = yScroll_Image * m_ZoomRatio
        
        'Translate the image rect (ImageRect_CanvasCoords) by the scroll bar values (which can be zero; that's fine).
        ' Remember that ImageRect_CanvasCoords gives us the pixel values of where the image appears on the canvas,
        ' when the scroll bars are at (0, 0).
        Dim translatedImageRect As RectF
        With translatedImageRect
            .Left = imageRect_CanvasCoords.Left - xScroll_Canvas
            .Top = imageRect_CanvasCoords.Top - yScroll_Canvas
            .Width = imageRect_CanvasCoords.Width
            .Height = imageRect_CanvasCoords.Height
        End With
        
        'This translated rect allows us to shortcut a lot of coordinate math, so cache a copy inside the source image.
        srcImage.ImgViewport.SetImageRectTranslated translatedImageRect
        
        'We now know where the full image lies, with zoom applied, relative to the canvas coordinate space.  Think of the canvas as
        ' a tiny window, and the image as a huge poster behind the window.  What we're going to do now is find the intersect rect
        ' between the window rect (which is easy - just the size of the canvas itself) and the image rect we've now calculated.
        Dim viewportRect As RectF
        srcImage.ImgViewport.SetIntersectState GDI_Plus.IntersectRectF(viewportRect, canvasRect_ActualPixels, translatedImageRect)
        
        If srcImage.ImgViewport.GetIntersectState Then
            
            'The intersection between the canvas and image is now stored in viewportRect.  Cool!  This is the destination rect of
            ' our viewport StretchBlt function.
            srcImage.ImgViewport.SetIntersectRectCanvas viewportRect
            
            'What we need to do now is reverse-map that rect back onto the image itself.  How do we do this?
            ' Well, we need two key pieces of information:
            ' 1) What's the relationship between (0, 0) on the canvas and (0, 0) on the image.  This value has already been determined
            '    for us, courtesy of the (Left, Top) values of ImageRect_CanvasCoords.
            ' 2) What is the scale between width/height on the canvas and width/height on the image?  This value is simply the
            '    zoom ratio, e.g. a zoom of 200% means that width/height measurements are twice as long on the canvas!
            
            'Start by mapping the (Top, Left) of this rect back onto the image.
            Dim srcLeft As Double, srcTop As Double
            Drawing.ConvertCanvasCoordsToImageCoords dstCanvas, srcImage, viewportRect.Left, viewportRect.Top, srcLeft, srcTop, False
            
            'Width and height are easy - just the width/height of the viewport, divided by the current zoom!
            Dim srcRectF As RectF
            srcRectF.Left = srcLeft
            srcRectF.Top = srcTop
            srcRectF.Width = viewportRect.Width / m_ZoomRatio
            srcRectF.Height = viewportRect.Height / m_ZoomRatio
            
            'We have now mapped the relevant viewport rect back into source coordinates, giving us everything we need for our render.
            
            'Before rendering the image, apply a checkerboard pattern to the viewport region of the source image's back buffer.
            ' TODO: cache g_CheckerboardPattern persistently, in GDI+ format, so we don't have to recreate it on every draw.
            With viewportRect
                GDI_Plus.GDIPlusFillDIBRect_Pattern srcImage.CanvasBuffer, .Left, .Top, .Width, .Height, g_CheckerboardPattern, , True
            End With
            
            'We can now use PD's rect-specific compositor to retrieve only the relevant section of the current viewport.
            ' Note that we request our own interpolation mode, and we determine this based on the user's viewport performance preference.
            
            'When we've been asked to maximize performance, use nearest neighbor for all zoom modes
            Dim vpInterpolation As GP_InterpolationMode
            If (g_ViewportPerformance = PD_PERF_FASTEST) Then
                vpInterpolation = GP_IM_NearestNeighbor
            Else
                
                'If we're zoomed-in, use nearest-neighbor regardless of the current settings
                If (m_ZoomRatio > 1#) Then
                    vpInterpolation = GP_IM_NearestNeighbor
                Else
                    If (g_ViewportPerformance = PD_PERF_BALANCED) Then vpInterpolation = GP_IM_Bilinear Else vpInterpolation = GP_IM_HighQualityBicubic
                End If
                
            End If
            
            srcImage.GetCompositedRect srcImage.CanvasBuffer, viewportRect, srcRectF, vpInterpolation, fullPipelineCall, CLC_Viewport, localViewportParams.renderScratchLayerIndex, localViewportParams.ptrToAlternateScratch
                
            'Cache the relevant section of the image, in case outside functions require it.
            srcImage.ImgViewport.SetIntersectRectImage srcRectF
            
        'The canvas and image do not overlap.  That's okay!  It means we don't have to do any compositing.  Exit now.
        Else
        
        End If
        
        'Before exiting, calculate the time spent in this stage
        m_TimeStage2 = VBHacks.GetTimerDifferenceNow(startTime)
        m_TotalTimeStage2 = m_TotalTimeStage2 + m_TimeStage2
        
        'Note that calls to this function may need to be relayed to other UI elements.  (For example, viewport rulers need to
        ' be repositioned, and if the navigator panel is open, it needs to reflect the new scroll position, if any.)
        
        'Such relays are not handled here, but if you're calling this pipeline function directly, be aware of the UI repercussions.
        ' Examining the pdCanvas class, particularly its scrollbars, is a good place to start for seeing what needs to be notified.
        
        'Pass control to the next stage of the pipeline.
        Stage3_CompositeCanvas srcImage, dstCanvas, ptrToViewportParams, True
        
        'If timing reports are enabled, we report them after the rest of the pipeline has finished.
        If g_DisplayTimingReports Then
            m_TotalTime = m_TotalTime + VBHacks.GetTimerDifferenceNow(startTime)
            'Debug.Print "Viewport render timing by stage (net, 2, 3, 4): " & VBHacks.GetTimeDiffNowAsString(startTime) & ", " & Format$(m_TimeStage2 * 1000#, "#0") & " ms, " & Format$(m_TimeStage3 * 1000#, "#0") & " ms, " & Format$(m_TimeStage4 * 1000#, "#0") & " ms"
        End If
    
    End If
    
End Sub

'Per its name, Stage1_InitializeBuffer() prepares a bunch of math related to viewport rendering.
' This includes:
    '1) Calculating all zoom-related math
    '2) Using (1) to determine limits of canvas scroll bars
    '3) Using (2) to determine canvas offsets, if the image is zoomed out far enough that dead space
    '    is present in the viewport.
    '4) (optionally) Using (1), (2), (3) to calculate new scroll bar values if the user requests it
    '    (e.g. for preserving cursor position during mousewheel-to-zoom)
    
'All subsequent pipeline operations operate on the values and regions calculated by this function.

'Because this function don't do any rendering, it should only be called under specific conditions;
' specifically, after:
    '1) an image is first loaded
    '2) a viewport's zoom value changes
    '3) the main PhotoDemon application window is resized
    '4) edits are applied that modify image size (resize, rotate, etc.)

'This function is very fast, but it always invokes a full rendering pipeline flush when it finishes.
' This is required because the canvas's front and back buffers are forcibly rebuilt by this function
' (because their size has changed), so *don't call this function* if a later pipeline stage suffices
' - you'll get better performance that way.

'This function also accepts a ParamArray.  This should only be used if "zoom to coordinate" behavior
' is required (e.g. when mousewheel zoom is invoked while the mouse is over a viewport).  To preserve
' the hovered point, you must pass two sets of target coordinates: one set in *canvas space*, and
' another set in *image space*.  (Both are required, because this function doesn't track past zoom
' values.  So when a viewport's zoom level changes, as by mousewheel, this function has no idea what
' the prior zoom level was, and a single set of coordinates isn't enough to maintain positioning.)

'So when using "zoom to coordinate" behavior, you must handle zoom changes in the following order:
' 1) cache x/y coordinate values in two coordinate spaces: image and canvas
' 2) disable automatic canvas redraws
' 3) change the zoom value; this allows the zoom engine to reconstruct conditional values, like "fit to window"
' 4) re-enable automatic canvas redraws (this can happen now, or after step 5 - just don't forget to do it!)
' 5) request a manual redraw via this function, and supply your previously cached x/y values
Public Sub Stage1_InitializeBuffer(ByRef srcImage As pdImage, ByRef dstCanvas As pdCanvas, ParamArray ExtraSettings() As Variant)
    
    'Errors have never been encountered in this version of PD, so this exists as a failsafe, only.
    On Error GoTo ViewportPipeline_Stage1_Error
    
    'Some conditions (e.g. program is shutting down) prevent viewport rendering
    If ViewportRenderingAllowed(srcImage, dstCanvas) Then
        
        'If no images are loaded, render a placeholder and exit.
        If (Not PDImages.IsImageNonNull()) Then
            FormMain.MainCanvas(0).ClearCanvas
        Else
            
            'The fundamental problem this first pipeline stage must solve is: how much screen real-estate
            ' does the target canvas have available, and how do we fit the current image - while considering
            ' transforms like scroll/zoom - into that real-estate.
            
            'First, an important clarification on my use of the terms "viewport" and "canvas".
            
            ' Viewport = the area of the screen dedicated to *just the image*
            ' Canvas = the area of the screen dedicated to *the full canvas*, including surrounding dead space
            '          (relevant when zoomed out, or when scrolling beyond the edge of an image)
            
            'Sometimes the viewport and canvas will have identical sizes.  Sometimes they will not.
            
            'If the canvas area is larger than the viewport area, the viewport will typically need to be
            ' centered within the canvas.
            
            'As noted above, callers can request special behavior via the ExtraSettings param array.  In most
            ' cases, these values aren't dealt with until the end of the function, but for "preserve center point"
            ' requests, we need to determine the current image+viewport center points in advance (because they'll
            ' change once new viewport and canvas rects are calculated).
            '
            'To that end, let's note any special requests before starting "serious" work.
            Dim specialRequestActive As Boolean, specialRequestID As PD_VIEWPORT_SPECIAL_REQUEST
            If (UBound(ExtraSettings) >= LBound(ExtraSettings)) Then
                specialRequestActive = True
                specialRequestID = CLng(ExtraSettings(0))
            End If
            
            'Because a full pipeline execution is time-consuming, this function has been manually profiled
            ' many times.  The following line is regularly enabled/disabled during alpha builds.
            ' pdDebug.LogAction "Preparing viewport: " & reasonForRedraw & " | (" & srcImage.imageID & ") "
            
            'This crucial value is the mathematical ratio of the current zoom value: 1 for 100%, 0.5 for 50%,
            ' 2 for 200%, etc.  We can't generate this automatically, because specialty zoom values (like
            ' "fit to window") must be externally generated by PD's zoom engine.
            m_ZoomRatio = Zoom.GetZoomRatioFromIndex(srcImage.GetZoomIndex())
            If (m_ZoomRatio = 0#) Then Exit Sub
            
            'Next, we must calculate a bunch of rects in various coordinate spaces.  Because PD 7.0 added the
            ' ability to scroll past the edge of the image ("overscroll"), these rects are crucial for figuring
            ' out the overlap between a zoomed image and the available canvas area.
            '
            '(In almost all cases, the width/height of the rect is calculated first, and the top/left comes later.)
            
            'First calculate an image rect, translated to the canvas coordinate space (e.g. multiplied by zoom).
            Dim imageRect_CanvasCoords As RectF
            With imageRect_CanvasCoords
                .Width = Int(srcImage.Width * m_ZoomRatio)
                .Height = Int(srcImage.Height * m_ZoomRatio)
            End With
            
            'Before querying the canvas object for sizes, make sure scroll bars are visible.
            ' (As of v7.0, viewport scrollbars are *always* visible - this simplifies things a bit
            ' vs previous versions, where we constantly toggled visibility based upon reaching the
            ' edge of the canvas.)
            FormMain.MainCanvas(0).SetScrollVisibility pdo_Both, True
            
            'Before we can position the image rect, we need to know the size of the canvas.  pdCanvas is
            ' responsible for determining this, as it knows things like scroll bar visibility, status bar
            ' placement, rulers, and any other relevant UI elements.
            Dim canvasRect_ActualPixels As RectF
            With canvasRect_ActualPixels
                .Left = 0
                .Top = 0
                .Width = dstCanvas.GetCanvasWidth()
                .Height = dstCanvas.GetCanvasHeight()
            End With
            
            'While here, we want to calculate a second rect for the canvas: its size, in *image* coordinates.
            Dim canvasRect_ImageCoords As RectF
            With canvasRect_ImageCoords
                .Left = 0
                .Top = 0
                .Width = canvasRect_ActualPixels.Width / m_ZoomRatio
                .Height = canvasRect_ActualPixels.Height / m_ZoomRatio
            End With
            
            'We now want to center the zoomed image relative to the canvas space.  The top-left of the centered
            ' image gives us a baseline for all scroll bar behavior, if the zoomed image is smaller than the
            ' available canvas space.
            With imageRect_CanvasCoords
                .Left = Int((canvasRect_ActualPixels.Width * 0.5) - (.Width * 0.5))
                .Top = Int((canvasRect_ActualPixels.Height * 0.5) - (.Height * 0.5))
            End With
            
            'imageRect_CanvasCoords now contains a RECTF of the image, with zoom applied, centered over the
            ' canvas.  Its (.Top, .Left) coordinate pair represents the (0, 0) position of the image, when the
            ' scrollbars are at (0, 0).  If these values lie outside the canvas rect, we want to reset them to
            ' (0, 0) position (so that (0, 0) in actual pixels represents pixel (0, 0) of the image, if the image
            ' is larger than the canvas).
            With imageRect_CanvasCoords
                If (.Left < 0) Then .Left = 0
                If (.Top < 0) Then .Top = 0
            End With
            
            'Before v7.0, scroll bars were only displayed when absolutely necessary.  In v7.0, I added "overscroll"
            ' capabilities (scroll bars are always available) to match trends with other image editors.  This makes
            ' painting near image edges much simpler and is a greatly improved UI experience.  As such, PD now
            ' assumes that scroll bars are always visible and enabled, regardless of zoom or image size - which also
            ' means we need to always calculate max/min scroll bar limits, regardless of image or canvas size.
            
            'Note that at present, scroll bars only move in single-pixel increments (in the image coordinate space),
            ' which makes life easier.  We basically want to allow the user to scroll long enough that they can
            ' create a "mostly empty" canvas.  How many pixels this requires depends on the size of the image relative
            ' to the current canvas.
            
            'Start by calculating the *required* scroll bar maximum: the amount of the image that cannot physically
            ' fit inside the canvas.
            Dim hScrollMin As Long, hScrollMax As Long, vScrollMin As Long, vScrollMax As Long
            hScrollMax = (srcImage.Width - canvasRect_ImageCoords.Width)
            vScrollMax = (srcImage.Height - canvasRect_ImageCoords.Height)
            
            'Minimum values are easy to calculate; always let the user scroll the image halfway off the screen
            hScrollMin = -1 * (canvasRect_ImageCoords.Width * 0.5)
            vScrollMin = -1 * (canvasRect_ImageCoords.Height * 0.5)
            
            'If hScrollMax or vScrollMax are negative, it means the canvas is larger (in that dimension) than the
            ' zoomed image.  When this happens, rely solely on the "halfway off screen" scroll measurement.
            If (hScrollMax > 0) Then hScrollMax = hScrollMax - hScrollMin Else hScrollMax = -hScrollMin
            If (vScrollMax > 0) Then vScrollMax = vScrollMax - vScrollMin Else vScrollMax = -vScrollMin
            
            'We now have scroll bar max/min values.  Forward them to the destination pdCanvas object.
            With dstCanvas
                .SetRedrawSuspension True
                .SetScrollMin pdo_Horizontal, hScrollMin
                .SetScrollMax pdo_Horizontal, hScrollMax
                .SetScrollMin pdo_Vertical, vScrollMin
                .SetScrollMax pdo_Vertical, vScrollMax
                .SetRedrawSuspension False
            End With
            
            'As a convenience to the user, we also make each scroll bar's LargeChange parameter proportional to the
            ' scroll bar's maximum value.
            If (hScrollMax > 15) And (m_ZoomRatio <= 1#) Then
                dstCanvas.SetScrollLargeChange pdo_Horizontal, hScrollMax \ 16
            Else
                dstCanvas.SetScrollLargeChange pdo_Horizontal, PDMath.Max2Int(64# / m_ZoomRatio, 1)
            End If
            
            If (vScrollMax > 15) And (m_ZoomRatio <= 1#) Then
                dstCanvas.SetScrollLargeChange pdo_Vertical, vScrollMax \ 16
            Else
                dstCanvas.SetScrollLargeChange pdo_Vertical, PDMath.Max2Int(64# / m_ZoomRatio, 1)
            End If
            
            'Scroll bars are now prepped and ready!
            
            'With all scroll bar data assembled, we have enough information to create a target back buffer.
            If (srcImage.CanvasBuffer.GetDIBWidth <> canvasRect_ActualPixels.Width) Or (srcImage.CanvasBuffer.GetDIBHeight <> canvasRect_ActualPixels.Height) Then
                srcImage.CanvasBuffer.CreateBlank canvasRect_ActualPixels.Width, canvasRect_ActualPixels.Height, 32, g_Themer.GetGenericUIColor(UI_CanvasElement), 255
            Else
                GDI_Plus.GDIPlusFillDIBRect srcImage.CanvasBuffer, 0, 0, canvasRect_ActualPixels.Width, canvasRect_ActualPixels.Height, g_Themer.GetGenericUIColor(UI_CanvasElement), 255, GP_CM_SourceCopy
            End If
            
            'Because subsequent stages of the pipeline need all the data we've assembled, store a copy of all
            ' calculated rects inside the passed pdImage object.
            With srcImage.ImgViewport
                .SetCanvasRectActualPixels canvasRect_ActualPixels
                .SetCanvasRectImageCoords canvasRect_ImageCoords
                .SetImageRectCanvasCoords imageRect_CanvasCoords
            End With
            
            'Want to debug viewport things?  Knowing rect contents is a good place to start:
            'With canvasRect_ActualPixels
            '    Debug.Print "canvasRect_ActualPixels", .Left, .Top, .Width, .Height
            'End With
            'With canvasRect_ImageCoords
            '    Debug.Print "canvasRect_ImageCoords", .Left, .Top, .Width, .Height
            'End With
            'With imageRect_CanvasCoords
            '    Debug.Print "imageRect_CanvasCoords", .Left, .Top, .Width, .Height
            'End With
            
            'The final step of this pipeline is optional.  If the user wants us to calculate specific scroll
            ' bar values, they must pass a special request enum via the function ParamArray().  At present,
            ' this function is capable of three different auto-calculations, which correspond to the three
            ' enum values of PD_VIEWPORT_SPECIAL_REQUEST:
            ' - VSR_ResetToZero: reset the scroll bar to (0, 0), which also centers the image when in "zoom-to-fit" mode
            ' - VSR_ResetToCustom: reset the scroll bar to two values supplied by the user (in (x, y) order)
            ' - VSR_AutoCenter: forcibly center the image, regardless of zoom
            ' - VSR_PreservePointPosition: given a point (typically the point under the mouse cursor), preserve its
            '    before-and-after position, even though zoom has changed!  This makes mousewheel scrolling intuitive.
            
            'Check for a param array now, and if nothing's found, skip straight to the next pipeline stage.
            If specialRequestActive Then
                
                'Regardless of what type of scroll bar setting we're applying, we need to disable automatic
                ' viewport redraws.  (Otherwise, changing scroll bar values will trigger viewport pipeline requests,
                ' wreaking havoc.)
                dstCanvas.SetRedrawSuspension True
                
                'The first extra setting defines the type of scroll bar handling request
                Select Case specialRequestID
                
                    Case VSR_ResetToZero
                        dstCanvas.SetScrollValue pdo_Both, 0
                        
                    Case VSR_ResetToCustom
                        dstCanvas.SetScrollValue pdo_Horizontal, CLng(ExtraSettings(1))
                        dstCanvas.SetScrollValue pdo_Vertical, CLng(ExtraSettings(2))
                    
                    'If the user has a point they want us to preserve, they must pass two sets of coordinates:
                    ' 1) The literal (x, y) of the mouse on the current canvas (e.g. the coordinates returned by a mouse event)
                    ' 2) The corresponding (x, y) of that mouse position *in the image coordinate space*
                    '
                    'Our goal is to make that same (x, y) point on the canvas correlate to the same (x, y) on the image,
                    ' regardless of any zoom/viewport/other changes we made earlier in this function.
                    Case VSR_PreservePointPosition
                        
                        Dim oldXCanvas As Single, oldYCanvas As Single, targetXImage As Single, targetYImage As Single
                        oldXCanvas = CSng(ExtraSettings(1))
                        oldYCanvas = CSng(ExtraSettings(2))
                        targetXImage = CSng(ExtraSettings(3))
                        targetYImage = CSng(ExtraSettings(4))
                        
                        'From the supplied coordinates, we know that image coordinate targetXImage was originally
                        ' located on the canvas at position oldXCanvas.  Our goal is to make targetXImage *remain*
                        ' at oldXCanvas position.
                        
                        'Start by converting targetX/Y/Image to the current canvas space.  This will give us
                        ' NewCanvasX/Y values that describe where the coordinates lie on the *new* canvas.
                        
                        '...then set a fake, "translated" image rect, that is correct for the case of h/v/scroll = 0.
                        ' (Normally stage 3 of the pipeline creates a translated rect, but we have to provide one now
                        ' because the canvas/image coordinate translation code relies on that rect!)
                        srcImage.ImgViewport.SetImageRectTranslated imageRect_CanvasCoords
                        
                        'With those values successfully set, we can now translate the target image coords into
                        ' canvas coords, for the case of h/v/scroll = 0.
                        Dim newXCanvas As Double, newYCanvas As Double
                        Drawing.ConvertImageCoordsToCanvasCoords dstCanvas, srcImage, targetXImage, targetYImage, newXCanvas, newYCanvas, False
                        
                        'Use the difference between newCanvasX and oldCanvasX (while accounting for zoom) to determine new
                        ' scroll bar values.
                        dstCanvas.SetScrollValue pdo_Horizontal, (newXCanvas - oldXCanvas) / m_ZoomRatio
                        dstCanvas.SetScrollValue pdo_Vertical, (newYCanvas - oldYCanvas) / m_ZoomRatio
                        
                End Select
                
                'Restore scroll-bar-originating viewport redraw requests
                dstCanvas.SetRedrawSuspension False
                
            End If
            
            'With our work here complete, we can pass control to the next pipeline stage.
            Stage2_CompositeAllLayers srcImage, dstCanvas, 0&, True
        
        End If
        
    'If viewport rendering is disallowed, attempt to render a placeholder UI before exiting
    Else
        
        'Because dstCanvas may not yet exist, forcibly invoke the default canvas
        FormMain.MainCanvas(0).ClearCanvas
        
    End If
    
    Exit Sub

'Note: error handling has never triggered
ViewportPipeline_Stage1_Error:
    PDDebug.LogAction "Viewport rendering paused due to unexpected error (#" & Err.Number & " - " & Err.Description & ")"
    
End Sub

'Before executing a pipeline step, this function needs to be called to see if viewport rendering is even allowed.
' (If the current viewport stage was directly invoked by a *previous* pipeline step, this check can be skipped,
' as it's assumed the parent already handled it.)
Private Function ViewportRenderingAllowed(ByRef srcImage As pdImage, ByRef dstCanvas As pdCanvas) As Boolean
    
    'First, see if viewport rendering has been forcibly disabled.
    ' (Detailed explanation: viewport redraws are automatically triggered by the main window's resize notifications.  When new images
    '  are loaded, the image tabstrip will likely appear, which in turn changes the available viewport space, just like a resize
    '  event.  To prevent this behavior from triggering multiple viewport render requests, m_DisableViewportRendering exists.)
    ViewportRenderingAllowed = (Not m_DisableViewportRendering)
    If ViewportRenderingAllowed Then
        
        'Make sure the source and destination rendering targets are valid
        ViewportRenderingAllowed = (Not dstCanvas Is Nothing) And (Not srcImage Is Nothing)
        
        'Finally, if the source image is inactive (e.g. it has been unloaded at some point in the past), do not redraw.
        ' (For performance reasons, PD may not release unused image sobject right away; their .IsActive flag will always
        ' be up-to-date, however.)
        If ViewportRenderingAllowed Then ViewportRenderingAllowed = srcImage.IsActive
        
    End If
    
End Function

'Call this function to disable *all* viewport pipeline stages.  Note that viewport rendering will remain disabled until
' EnableRendering() is called - so do not forget to call it when you're done!
Public Sub DisableRendering()
    m_DisableViewportRendering = True
End Sub

'Rendering is enabled by default.  This function only needs to be called after DisableRendering() has been forcibly invoked.
Public Sub EnableRendering()
    m_DisableViewportRendering = False
End Sub

Public Function IsRenderingEnabled() As Boolean
    IsRenderingEnabled = (Not m_DisableViewportRendering)
End Function

'When all images have been unloaded, the temporary front buffer can also be erased to keep memory usage as low as possible.
' While not actually part of the viewport pipeline, I find it intuitive to store this function here.
Public Sub EraseViewportBuffers()
    If (Not m_FrontBuffer Is Nothing) Then
        m_FrontBuffer.EraseDIB
        Set m_FrontBuffer = Nothing
    End If
End Sub

'After interacting with this module, you can call this sub to relay any new viewport changes
' to UI elements across the program.  For example, if zoom has been changed, everything from
' on-canvas scrollbars to the top-right navigator window needs to be notified of the change.
' Rather than worry about who needs to be notified of what, just call this sub and it will
' take care of the rest.
Public Sub NotifyEveryoneOfViewportChanges()
    
    'Let the right-hand toolbar know that the viewport has changed (this will relay changes
    ' to child windows, like the navigator)
    toolbar_Layers.NotifyViewportChange
    
    'Notify the canvas itself of any changes
    FormMain.MainCanvas(0).NotifyViewportChanges
    
End Sub

'Report the current viewport performance profiling data to pdDebug.  Useless in non-debug builds.
Public Sub ReportViewportProfilingData()

    If (m_TotalTime <> 0#) Then
        PDDebug.LogAction "Final viewport perf data, by stage:"
        PDDebug.LogAction "2: " & Format$((m_TotalTimeStage2 / m_TotalTime) * 100, "00.0") & "%"
        PDDebug.LogAction "3: " & Format$((m_TotalTimeStage3 / m_TotalTime) * 100, "00.0") & "%"
        PDDebug.LogAction "4: " & Format$((m_TotalTimeStage4 / m_TotalTime) * 100, "00.0") & "%"
    End If

End Sub
