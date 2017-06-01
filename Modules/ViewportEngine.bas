Attribute VB_Name = "ViewportEngine"
'***************************************************************************
'Viewport Handler - builds and draws the image viewport and associated scroll bars
'Copyright 2001-2017 by Tanner Helland
'Created: 4/15/01
'Last updated: 29/November/16
'Last update: reinstate all color management code under LittleCMS (instead of the Windows ICM engine, which is a hot mess)
'
'Module for handling the image viewport.  The render pipeline works as follows:
' - ViewportEngine.Stage1_InitializeBuffer: for recalculating all viewport variables and controls (done only when the zoom value is changed,
'                                             or the user switches to a different image or loads a new picture)
' - ViewportEngine.Stage2_CompositeAllLayers: (re)assemble the layered image, with all non-destructive changes taken into account
' - ViewportEngine.Stage3_ExtractRelevantRegion: copy the relevant portion of the image out of the composite and into its own buffer.  This stage
'                                                 is skipped for the accelerated version of the pipeline.
' - ViewportEngine.Stage4_CompositeCanvas: blend the image and any canvas UI elements together.  At present, this includes shadow underlay and
'                                           selection highlight/lightboxing, if active.
' - ViewportEngine.Stage5_FlipBufferAndDrawUI: render the finalized image+canvas, and overlay any interactive UI elements (transform nodes, etc)
'
'PhotoDemon tries to be intelligent about calling the lowest routine in the pipeline, so it can render the viewport quickly
' regardless of zoom or scroll values.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Due to the complexity of viewport rendering, some viewport pipeline calculations are stored at module-level, while others are
' stored inside the source pdImage or destination pdCanvas used for a given rendering.  The module-level variables in particular are
' used to improve efficiency over passing objects as parameters, because PD's viewport pipeline is as an "out of order" pipeline.
' Different user interations trigger execution of the pipeline at different stages, which is crucial for maximizing viewport performance.
'
'As such, pipeline functions must be *very cautious* about modifying module-level values, or viewport-related values stored within
' source pdImage or destination pdCanvas objects.  CONSIDER YOURSELF WARNED.

'If external functions require special scroll bar treatment in the initial pipeline stage, they must pass one of these enums
' as the first entry in the associated ParamArray().
Public Enum PD_VIEWPORT_SPECIAL_REQUEST
    VSR_ResetToZero = 0
    VSR_ResetToCustom = 1
    VSR_PreservePointPosition = 2
End Enum

#If False Then
    Private Const VSR_ResetToZero = 0, VSR_ResetToCustom = 1, VSR_PreservePointPosition = 2
#End If

'The ZoomVal value is the actual coefficient for the current zoom value.  (For example, 0.50 for "50% zoom")  Multiple pipeline stages
' use this, so it's cached by the first pipeline staged and simply reused after that.
Private m_ZoomRatio As Double

'm_FrontBuffer holds the final composited image, including any non-interactive overlays (like selection highlight/lightbox effects)
Private m_FrontBuffer As pdDIB

'As part of continued viewport optimizations, we track the amount of time spent in each viewport stage.  Note that stage 5 is ignored because
' it only affects changes in zoom values (or switching between images).
Private m_TimeStage2 As Currency, m_TimeStage3 As Currency, m_TimeStage4 As Currency, m_TimeStage5 As Currency

'Stage5_FlipBufferAndDrawUI is the final stage of the viewport pipeline.  It will flip the composited canvas image to the
' destination pdCanvas object, and apply any final UI elements as well - control nodes, custom cursors, etc.  This step is
' very fast, and should be used whenever full compositing is unnecessary.
'
'As part of the buffer flip, this stage will also activate and apply color management to the completed front buffer.
' (Still TODO is fixing the canvas to not rely on AutoRedraw, which will spare us having to re-activate color management on every draw.)
'
'Note also that this stage is the only one to make use of the optional POI ID parameter.  If supplied, it will forward this to any
' UI functions that accept POI identifiers.  (Because the viewport is agnostic to underlying UI complexities, by design, it is up to
' the caller to optimize POI-based requests, e.g. not forwarding them unless the POI has changed, etc)
Public Sub Stage5_FlipBufferAndDrawUI(ByRef srcImage As pdImage, ByRef dstCanvas As pdCanvas, Optional ByVal curPOI As PD_PointOfInterest = poi_Undefined)

    Dim startTime As Currency
    VBHacks.GetHighResTime startTime
    
    'If no images have been loaded, clear the canvas and exit
    If (g_OpenImageCount <= 0) Then
        FormMain.mainCanvas(0).ClearCanvas
        Exit Sub
    Else
        FormMain.mainCanvas(0).SetScrollVisibility PD_BOTH, True
    End If
    
    'Apply a few failsafe checks to make sure all required rendering objects exist
    If (dstCanvas Is Nothing) Then Exit Sub
    If (srcImage Is Nothing) Then Exit Sub
    If (Not srcImage.IsActive) Then Exit Sub
    
    'Flip the viewport buffer over to the canvas control.  Any additional rendering must now happen there.
    BitBlt dstCanvas.hDC, 0, 0, m_FrontBuffer.GetDIBWidth, m_FrontBuffer.GetDIBHeight, m_FrontBuffer.GetDIBDC, 0, 0, vbSrcCopy
    
    'Lastly, do any tool-specific rendering directly onto the form.
    Select Case g_CurrentTool
    
        'The nav tool provides two render options at present: draw layer borders, and draw layer transform nodes
        Case NAV_MOVE
        
            'If the user has requested visible layer borders, draw them now
            If CBool(toolpanel_MoveSize.chkLayerBorder) Then Drawing.DrawLayerBoundaries dstCanvas, srcImage, srcImage.GetActiveLayer
            
            'If the user has requested visible transformation nodes, draw them now.
            ' (TODO: cache these values in either public variables, or inside this module via some kind of setViewportProperties
            '        function - either way, that will let us access drawing settings much more quickly!)
            If CBool(toolpanel_MoveSize.chkLayerNodes) Then Drawing.DrawLayerCornerNodes dstCanvas, srcImage, srcImage.GetActiveLayer, curPOI
            
            'Same as above, but for the current rotation node
            If CBool(toolpanel_MoveSize.chkRotateNode) Then Drawing.DrawLayerRotateNode dstCanvas, srcImage, srcImage.GetActiveLayer, curPOI
            
        'Selections are always rendered onto the canvas.  If a selection is active AND a selection tool is active, we can also
        ' draw transform nodes around the selection area.  (Note that lasso selections are currently an exception to this rule;
        ' they only support the "move" interaction, which is applied by click-dragging anywhere in the lasso region.)
        Case SELECT_RECT, SELECT_CIRC, SELECT_LINE, SELECT_POLYGON, SELECT_WAND ', SELECT_LASSO
            If srcImage.IsSelectionActive Then srcImage.mainSelection.RenderTransformNodes srcImage, dstCanvas
            
        'Text tools currently draw layer boundaries at all times; I'm working on this (TODO!)
        Case VECTOR_TEXT, VECTOR_FANCYTEXT
            
            If pdImages(g_CurrentImage).GetActiveLayer.IsLayerText Then
                Drawing.DrawLayerBoundaries dstCanvas, srcImage, srcImage.GetActiveLayer
                Drawing.DrawLayerCornerNodes dstCanvas, srcImage, srcImage.GetActiveLayer, curPOI
                Drawing.DrawLayerRotateNode dstCanvas, srcImage, srcImage.GetActiveLayer, curPOI
            End If
            
        'Paintbrush tools use the brush engine to paint a custom brush outline at the current mouse position
        Case PAINT_BASICBRUSH, PAINT_SOFTBRUSH
            
            If (Paintbrush.IsBrushActive Or FormMain.mainCanvas(0).IsMouseOverCanvas) Then
                Paintbrush.RenderBrushOutline dstCanvas
            End If
            
            
    End Select
    
    'FYI, in the future, any additional UI compositing can be handled here.
    
    'With all rendering complete, copy the form's image into the .Picture (e.g. render it on-screen) and refresh
    dstCanvas.RequestViewportRedraw True
    
    'Before exiting, calculate the time spent in this stage
    m_TimeStage5 = VBHacks.GetTimerDifferenceNow(startTime)
    
End Sub

'Stage4_CompositeCanvas takes the current canvas (which has a checkerboard and fully layered image drawn atop it) and applies a few
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
Public Sub Stage4_CompositeCanvas(ByRef srcImage As pdImage, ByRef dstCanvas As pdCanvas, Optional ByVal curPOI As PD_PointOfInterest = poi_Undefined)
    
    Dim startTime As Currency
    VBHacks.GetHighResTime startTime
    
    'If no images have been loaded, clear the canvas and exit
    If (g_OpenImageCount = 0) Then
        FormMain.mainCanvas(0).ClearCanvas
        Exit Sub
    End If

    'Make sure the canvas is valid
    If (dstCanvas Is Nothing) Then Exit Sub
    
    'If the image associated with this form is inactive, ignore this request
    If (srcImage Is Nothing) Then Exit Sub
    If (Not srcImage.IsActive) Then Exit Sub

    'Create the front buffer as necessary
    If (m_FrontBuffer Is Nothing) Then Set m_FrontBuffer = New pdDIB
    If (m_FrontBuffer.GetDIBWidth <> srcImage.canvasBuffer.GetDIBWidth) Or (m_FrontBuffer.GetDIBHeight <> srcImage.canvasBuffer.GetDIBHeight) Then
        m_FrontBuffer.CreateFromExistingDIB srcImage.canvasBuffer
    Else
        BitBlt m_FrontBuffer.GetDIBDC, 0, 0, srcImage.canvasBuffer.GetDIBWidth, srcImage.canvasBuffer.GetDIBHeight, srcImage.canvasBuffer.GetDIBDC, 0, 0, vbSrcCopy
    End If
    
    'Retrieve a copy of the intersected viewport rect; subsequent rendering ops may use this to optimize their operations
    Dim viewportIntersectRect As RECTF
    srcImage.imgViewport.GetIntersectRectCanvas viewportIntersectRect
    
    '*Now* is when we want to apply color management to the front buffer.  (For performance reasons, UI elements drawn atop the canvas are not
    ' color-managed - only the image itself is.)  Note also that although the front buffer is 32-bpp, it is always fully opaque, so we can
    ' notify the color management engine that alpha bytes can be completely ignored.
    ColorManagement.ApplyDisplayColorManagement_RectF m_FrontBuffer, viewportIntersectRect, , False
    
    'Check to see if a selection is active.  If it is, we want to render it now, directly atop the front buffer.  This allows any
    ' subsequent overlays (e.g. brush outlines) to appear "on top" of the selection, without us needing to redraw the selection outline
    ' on every overlay render.
    If srcImage.IsSelectionActive Then
        
        'The selection engine handles the actual rendering process
        srcImage.mainSelection.RenderSelectionToViewport m_FrontBuffer, srcImage, dstCanvas
    
    End If
    
    'Before exiting, calculate the time spent in this stage
    m_TimeStage4 = VBHacks.GetTimerDifferenceNow(startTime)
    
    'Pass the completed front buffer to the final stage of the pipeline, which will flip everything to the screen and render any
    ' remaining UI elements!
    Stage5_FlipBufferAndDrawUI srcImage, dstCanvas, curPOI
    
End Sub

'Stage3_ExtractRelevantRegion is used to composite the current image onto the source pdImage's back buffer.  This function does not
' perform any initialization or pre-rendering checks, so you cannot use it if zoom is changed, or if the viewport area has changed.
' (Stage1_InitializeBuffer is used for that.)
'
'The most common use-case for this function is the use of scrollbars, or non-destructive layer changes that require a recomposite
' of the image, but not a full re-creation and calculation of the viewport and canvas buffers.
'
'The optional pipelineOriginatedAtStageOne parameter lets this function know if a full pipeline purge is required.  Some caches
' may need to be regenerated from scratch after zoom changes.
Public Sub Stage3_ExtractRelevantRegion(ByRef srcImage As pdImage, ByRef dstCanvas As pdCanvas, Optional ByVal pipelineOriginatedAtStageOne As Boolean = False, Optional ByVal curPOI As PD_PointOfInterest = poi_Undefined, Optional ByVal renderScratchLayerIndex As Long = -1)
    
    Dim startTime As Currency
    VBHacks.GetHighResTime startTime
    
    'Regardless of the pipeline branch we follow, we need local copies of the relevant region rects calculated by stage 1 of the pipeline.
    Dim imageRect_CanvasCoords As RECTF, canvasRect_ImageCoords As RECTF, canvasRect_ActualPixels As RECTF
    With srcImage.imgViewport
        .GetCanvasRectActualPixels canvasRect_ActualPixels
        .GetCanvasRectImageCoords canvasRect_ImageCoords
        .GetImageRectCanvasCoords imageRect_CanvasCoords
    End With
    
    'We also need to wipe the back buffer
    GDI_Plus.GDIPlusFillDIBRect srcImage.canvasBuffer, 0, 0, srcImage.canvasBuffer.GetDIBWidth, srcImage.canvasBuffer.GetDIBHeight, g_Themer.GetGenericUIColor(UI_CanvasElement), 255, GP_CM_SourceCopy
    
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
    xScroll_Image = dstCanvas.GetScrollValue(PD_HORIZONTAL)
    yScroll_Image = dstCanvas.GetScrollValue(PD_VERTICAL)
    
    'Next, let's calculate these *in the canvas coordinate space* (e.g. with zoom applied)
    If m_ZoomRatio = 0 Then m_ZoomRatio = g_Zoom.GetZoomValue(srcImage.GetZoom)
    xScroll_Canvas = xScroll_Image * m_ZoomRatio
    yScroll_Canvas = yScroll_Image * m_ZoomRatio
    
    'Translate the image rect (ImageRect_CanvasCoords) by the scroll bar values (which can be zero; that's fine).
    ' Remember that ImageRect_CanvasCoords gives us the pixel values of where the image appears on the canvas,
    ' when the scroll bars are at (0, 0).
    Dim translatedImageRect As RECTF
    With translatedImageRect
        .Left = imageRect_CanvasCoords.Left - xScroll_Canvas
        .Top = imageRect_CanvasCoords.Top - yScroll_Canvas
        .Width = imageRect_CanvasCoords.Width
        .Height = imageRect_CanvasCoords.Height
    End With
    
    'This translated rect allows us to shortcut a lot of coordinate math, so cache a copy inside the source image.
    srcImage.imgViewport.SetImageRectTranslated translatedImageRect
    
    'We now know where the full image lies, with zoom applied, relative to the canvas coordinate space.  Think of the canvas as
    ' a tiny window, and the image as a huge poster behind the window.  What we're going to do now is find the intersect rect
    ' between the window rect (which is easy - just the size of the canvas itself) and the image rect we've now calculated.
    Dim viewportRect As RECTF
    srcImage.imgViewport.SetIntersectState GDI_Plus.IntersectRectF(viewportRect, canvasRect_ActualPixels, translatedImageRect)
    
    If srcImage.imgViewport.GetIntersectState Then
        
        'The intersection between the canvas and image is now stored in viewportRect.  Cool!  This is the destination rect of
        ' our viewport StretchBlt function.
        srcImage.imgViewport.SetIntersectRectCanvas viewportRect
        
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
        Dim srcRectF As RECTF
        srcRectF.Left = srcLeft
        srcRectF.Top = srcTop
        srcRectF.Width = viewportRect.Width / m_ZoomRatio
        srcRectF.Height = viewportRect.Height / m_ZoomRatio
        
        'We have now mapped the relevant viewport rect back into source coordinates, giving us everything we need for our render.
        
        'Before rendering the image, apply a checkerboard pattern to the viewport region of the source image's back buffer.
        ' TODO: cache g_CheckerboardPattern persistently, in GDI+ format, so we don't have to recreate it on every draw.
        With viewportRect
            GDI_Plus.GDIPlusFillDIBRect_Pattern srcImage.canvasBuffer, .Left, .Top, .Width, .Height, g_CheckerboardPattern, , True
        End With
        
        'As a failsafe, perform a GDI+ check.  PD probably won't work at all without GDI+, so I could look at dropping this check
        ' in the future... but for now, we leave it, just in case.
        If Drawing2D.IsRenderingEngineActive(P2_GDIPlusBackend) Then
                
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
            
            srcImage.GetCompositedRect srcImage.canvasBuffer, viewportRect, srcRectF, vpInterpolation, pipelineOriginatedAtStageOne, CLC_Viewport, renderScratchLayerIndex
                    
        'This is an emergency fallback, only.  PD won't work without GDI+, so rendering the viewport is pointless.
        Else
            Message "WARNING!  GDI+ could not be found.  (PhotoDemon requires GDI+ for proper program operation.)"
        End If
        
        'Cache the relevant section of the image, in case outside functions require it.
        srcImage.imgViewport.SetIntersectRectImage srcRectF
        
    'The canvas and image do not overlap.  That's okay!  It means we don't have to do any compositing.  Exit now.
    Else
    
    End If
    
    'Before exiting, calculate the time spent in this stage
    m_TimeStage3 = VBHacks.GetTimerDifferenceNow(startTime)
    
    'Note that calls to this function may need to be relayed to other UI elements.  (For example, viewport rulers need to
    ' be repositioned, and if the navigator panel is open, it needs to reflect the new scroll position, if any.)
    
    'Such relays are not handled here, but if you're calling this pipeline function directly, be aware of the UI repercussions.
    ' Examining the pdCanvas class, particularly its scrollbars, is a good place to start for seeing what needs to be notified.
    
    'Pass control to the next stage of the pipeline.
    Stage4_CompositeCanvas srcImage, dstCanvas, curPOI

End Sub

'Stage2_CompositeAllLayers is used to composite the current layer stack (while accounting for all non-destructive modifications)
' into a single, final image.  This image is stored in the source pdImage's back buffer.  This function does not perform any
' initialization or pre-rendering checks, so you cannot use it if zoom is changed, or if the viewport area has changed.
' (Stage1_InitializeBuffer is used for that.)
'
'Similarly, depending on the active render mode, this stage may be redundant for things like viewport scrollbars, as those only
' require Stage3_ExtractRelevantRegion, as the image has already been assembled.
'
'The most common use-case for this function is changes made to individual layers, including non-destructive layer changes that
' require a recomposite of the image, but not a full re-calculation of canvas measurements and the like.
'
'The optional pipelineOriginatedAtStageOne parameter lets this function know if a full pipeline purge is required.  Some caches
' may need to be regenerated from scratch after zoom changes.
Public Sub Stage2_CompositeAllLayers(ByRef srcImage As pdImage, ByRef dstCanvas As pdCanvas, Optional ByVal pipelineOriginatedAtStageOne As Boolean = False, Optional ByVal curPOI As PD_PointOfInterest = poi_Undefined, Optional ByVal renderScratchLayerIndex As Long = -1)
    
    'This function can return timing reports if desired; at present, this is automatically activated in PRE-ALPHA and ALPHA builds,
    ' but disabled for BETA and PRODUCTION builds; see the LoadTheProgram() function for details.
    Dim startTime As Currency
    VBHacks.GetHighResTime startTime
    
    'Like the previous stage of the pipeline, we start by performing a number of "do not render the viewport at all" checks.
    
    'First, and most obvious, is to exit now if the public g_AllowViewportRendering parameter has been forcibly disabled.
    If (Not g_AllowViewportRendering) Then Exit Sub
    
    'Make sure the target canvas is valid
    If (dstCanvas Is Nothing) Then Exit Sub
    
    'If the pdImage object associated with this form is inactive, ignore this request
    If (srcImage Is Nothing) Then Exit Sub
    If (Not srcImage.IsActive) Then Exit Sub
    
    'Stage 1 of the pipeline (Stage1_InitializeBuffer) prepared srcImage.compositeBuffer for us.  The goal of this stage
    ' is simple: fill the compositeBuffer object with a fully composited copy of the current image.
    
    'Note that only the "Best Quality" viewport mode actually requires this step.  The accelerated viewport pipeline assembles a
    ' shrunken version of the image, so it never requires a full-sized composite for the viewport
    
    '****************************************
    'NOTE!  As of v7.0, this option has been disabled.  Look at viewport pipeline stage 3 for additional details.
    'If (g_ViewportPerformance = PD_PERF_BESTQUALITY) Then
    '
    '    'Notify the parent object that a prepared composite buffer is required.  If the buffer is dirty, the parent will regenerate
    '    ' the composite for us.
    '    srcImage.RebuildCompositeBuffer renderScratchLayerIndex
    '
    ''Other viewport performance settings can automatically proceed to stage 3
    'End If
    '****************************************
    
    'Before exiting, calculate the time spent in this stage
    m_TimeStage2 = VBHacks.GetTimerDifferenceNow(startTime)
    
    'Pass control to the next stage of the pipeline.
    Stage3_ExtractRelevantRegion srcImage, dstCanvas, pipelineOriginatedAtStageOne, curPOI, renderScratchLayerIndex
    
    'If timing reports are enabled, we report them after the rest of the pipeline has finished.
    If g_DisplayTimingReports Then
        'Debug.Print "Viewport render timing: " & Format(CStr(VBHacks.GetTimerDifferenceNow(startTime) * 1000), "0000.00") & " ms"
        Debug.Print "Viewport render timing by stage (net, 2, 3, 4, 5): " & Format(CStr(VBHacks.GetTimerDifferenceNow(startTime) * 1000), "###0.00") & " ms, " & Format(CStr(m_TimeStage2 * 1000), "###0.00") & " ms, " & Format(CStr(m_TimeStage3 * 1000), "###0.00") & " ms, " & Format(CStr(m_TimeStage4 * 1000), "###0.00") & " ms, " & Format(CStr(m_TimeStage5 * 1000), "###0.00") & " ms"
    End If
    
End Sub

'Per its name, Stage1_InitializeBuffer is responsible for preparing a bunch of math related to viewport rendering.  Its duties include:
    '1) Calculating all zoom-related math
    '2) Determining max/min values of scroll bars
    '3) Canvas offsets, if the image is zoomed out far enough that dead space is present in the viewport.
    '4) (optionally) Calculating new scroll bar values if the user requests it (e.g. for preserving cursor position during mousewheel-to-zoom)
    
'This function is crucial, because all subsequent pipeline operations operate on the rectangles determined by this function.

'Because this function does no actual rendering - only preparation math - it only needs to be executed under specific conditions,
' namely when:
    '1) an image is first loaded
    '2) the viewport's zoom value is changed
    '3) the main PhotoDemon window is resized
    '4) edits that modify an image's size (resizing, rotating, etc - basically anything that changes the relationship between image size
    '   and the canvas buffer(s))

'Because the full rendering pipeline must be executed when this function is called, it is considered highly expensive, even though
' the math it performs is relatively quick.  The main issue caused by this function is that the current canvas's front and back buffers
' will need to be recreated (instead of just reused, as their size has likely changed), so whenever you need to call the viewport to
' request a redraw, try to figure out how late in the pipeline you can call - performance will improve accordingly.

'While this function is primarily concerned with the math required to handle zoom and scroll operations correctly, there are a few
' additional parameters that are occasionally necessary, which is why a ParamArray is used.  For details on these, please refer to the
' "Zoom to Coordinate" behavior, which is used when the mousewheel is invoked while over the current viewport.  That function will pass
' additional targetX and targetY parameters to this function, which are then used to automatically set matching scroll bar values after
' viewport calculations are complete, in a way that preserves the on-screen position of the passed coordinate.  (Note that the function
' does this as closely as it can, but some zoom changes make this impossible, such as zooming out to a point that scroll bars cannot
' physically reach).

'As an important follow-up note, two sets of target coordinates must be passed for this capability to work: one set of coordinates
' in *canvas space*, and one set in *image space*.  Both are required, because Stage1_InitializeBuffer doesn't keep track of past zoom values.
' This means that once the viewport's zoom level has been changed (as will likely happen prior to calling this function, by mousewheel),
' this function doesn't know what the prior zoom level was - and a single set of coordinates is not enough to maintain image positioning.

'Thus, when making use of "zoom to coordinate" behavior, you must handle zoom changes in the following order:
' 1) cache x/y coordinate values in two coordinate spaces: image and canvas
' 2) disable automatic canvas redraws
' 3) change the zoom value; this allows the zoom engine to reconstruct conditional values, like "fit to window"
' 4) re-enable automatic canvas redraws (this can happen now, or after step 5 - just don't forget to do it!)
' 5) request a manual redraw via Stage1_InitializeBuffer, and be sure supply your previously cached x/y values
Public Sub Stage1_InitializeBuffer(ByRef srcImage As pdImage, ByRef dstCanvas As pdCanvas, ParamArray ExtraSettings() As Variant)
    
    
    On Error GoTo ViewportPipeline_Stage1_Error
    
    
    'This initial stage of the pipeline contains a number of "do not render the viewport at all" checks.
    
    'First, and most obvious, is to exit now if the public g_AllowViewportRendering parameter has been forcibly disabled.
    ' (Detailed explanation: this routine is automatically triggered by the main window's resize notifications.  When new images
    '  are loaded, the image tabstrip will likely appear, which in turn changes the available viewport space, just like a resize
    '  event.  To prevent this behavior from triggering multiple Stage1_InitializeBuffer requests, g_AllowViewportRendering exists.)
    If (Not g_AllowViewportRendering) Then Exit Sub
    
    'Second, exit if the destination canvas has not been initialized yet; this may happen during program initialization.
    If (dstCanvas Is Nothing) Then Exit Sub
    
    'Third, exit if no images have been loaded.  The canvas object itself will render a blank placeholder image.
    If (g_OpenImageCount = 0) Then
        FormMain.mainCanvas(0).ClearCanvas
        Exit Sub
    End If
    
    'Fourth, exit if the source image doesn't exist.  (This check exists as an emergency failsafe, only.)
    If (srcImage Is Nothing) Then Exit Sub
    
    'Fifth, if the source image is inactive (e.g. it has been unloaded at some point in the past), do not execute a redraw.
    ' For performance reasons, PD does not shrink its primary pdImages() array unless required due to memory pressure.
    ' Instead, it just deactivates entries by marking the .IsActive property - so that property must be considered
    ' *prior* to executing image events.
    If (Not srcImage.IsActive) Then Exit Sub
    
    
    'If we made it all the way here, the full viewport pipeline needs to be executed.
    
    'The fundamental problem this first pipeline stage must solve is: how much screen real-estate do we have to work with, and how
    ' will we fit the image into that real-estate.
    
    'Potentially problematic is future feature additions, like rulers, which may interfere with our available viewport real-estate.
    ' To try and preempt changes from such such features, you'll notice various calls into the main pdCanvas object.
    ' My goal is to let pdCanvas calculate the positioning of such features, so only minimal changes are required here.
    ' (I mention this because at present, those functions likely return "0" - but they still serve a purpose!)
    
    'Finally, an important clarification is our use of the terms "viewport" and "canvas".
    
    ' Viewport = the area of the screen dedicated to *just the image*
    ' Canvas = the area of the screen dedicated to *the full canvas*, including any surrounding dead space (relevant when zoomed out,
    '           or scrolled past the edge of the image)
    
    'Sometimes the viewport and canvas rects will be identical.  Sometimes they will not.  If the canvas rect is larger than
    ' the viewport rect, the viewport rect will typically be centered within the viewport area.
    
    'The caller can request special behavior via the ExtraSettings param array.  In most cases, we don't deal with these results until
    ' the end of the function, but for the "preserve center point" request, we need to determine the current image+viewport center points
    ' in advance (as we'll change them once we calculate all the new viewport rects).
    '
    'To that end, note any special requests now.
    Dim specialRequestActive As Boolean, specialRequestID As PD_VIEWPORT_SPECIAL_REQUEST
    If (UBound(ExtraSettings) >= LBound(ExtraSettings)) Then
        specialRequestActive = True
        specialRequestID = CLng(ExtraSettings(0))
    End If
    
    'Because a full pipeline execution is time-consuming, I carefully track hits to this initial function to try and minimize how frequently
    ' it's called.  Feel free to comment out this line if you don't find such updates helpful.
    ' Debug.Print "Preparing viewport: " & reasonForRedraw & " | (" & srcImage.imageID & ") "
    
    'This crucial value is the mathematical ratio of the current zoom value: 1 for 100%, 0.5 for 50%, 2 for 200%, etc.
    ' We can't generate this automatically, because specialty zoom values (like "fit to window") must be externally generated
    ' by PD's zoom handler.
    m_ZoomRatio = g_Zoom.GetZoomValue(srcImage.GetZoom)
    
    'Next, we're going to calculate a bunch of rects in various coordinate spaces.  Because PD 7.0 added the ability to scroll past the
    ' edge of the image (at any zoom), these rects are crucial for figuring out the overlap between the zoomed image, and the available
    ' canvas area.
    '
    'In almost all cases, the width/height of the rect is calculated first, and the top/left comes later.
    
    'First is the image, translated to the canvas coordinate space (e.g. multiplied by zoom).
    Dim imageRect_CanvasCoords As RECTF
    With imageRect_CanvasCoords
        .Width = (srcImage.Width * m_ZoomRatio)
        .Height = (srcImage.Height * m_ZoomRatio)
    End With
    
    'Before we can position the image rect, we need to know the size of the canvas.  pdCanvas is responsible for determining this, as it must
    ' account for the positioning of scroll bars, a status bar, rulers, and whatever else the user has enabled.
    Dim canvasRect_ActualPixels As RECTF
    With canvasRect_ActualPixels
        .Left = 0
        .Top = 0
        .Width = dstCanvas.GetCanvasWidth()
        .Height = dstCanvas.GetCanvasHeight()
    End With
    
    'While here, we want to calculate a second rect for the canvas: its size, in image coordinates.
    Dim canvasRect_ImageCoords As RECTF
    With canvasRect_ImageCoords
        .Left = 0
        .Top = 0
        .Width = canvasRect_ActualPixels.Width / m_ZoomRatio
        .Height = canvasRect_ActualPixels.Height / m_ZoomRatio
    End With
    
    'We now want to center the zoomed image relative to the canvas space.  The top-left of the centered image gives us a baseline
    ' for all scroll bar behavior, if the image is smaller than the available canvas space.
    With imageRect_CanvasCoords
        .Left = (canvasRect_ActualPixels.Width * 0.5) - (.Width * 0.5)
        .Top = (canvasRect_ActualPixels.Height * 0.5) - (.Height * 0.5)
    End With
    
    'NEW IN 7.0: convert our calculated RectFs to their nearest integer-only estimates.  This should solve some obnoxious,
    ' persistent issues with edge handling during rendering.
    ' (I have suspended this feature pending further testing.)
    'PDMath.GetNearestIntRectF canvasRect_ActualPixels
    'PDMath.GetNearestIntRectF canvasRect_ImageCoords
    'PDMath.GetNearestIntRectF imageRect_CanvasCoords
    
    'imageRect_CanvasCoords now contains a RECTF of the image, with zoom applied, centered over the canvas.  The (.Top, .Left)
    ' coordinate pair of this rect represents the (0, 0) position of the image, when the scrollbars are (0, 0).  As such, if they
    ' lie outside the canvas rect, we want to reset them to (0, 0) position (so that (0, 0) in actual pixels represents pixel (0, 0)
    ' of the image, if the image is larger than the canvas).
    With imageRect_CanvasCoords
        If (.Left < 0) Then .Left = 0
        If (.Top < 0) Then .Top = 0
    End With
    
    'Pre-7.0, scroll bars were only displayed if absolutely necessary.  With the addition of paint tools, this is longer practical,
    ' so we now assume that scroll bars are always visible and enabled, regardless of zoom or image size - which also means we need
    ' to always calculate max/min scroll bar limits, regardless of the current image or canvas size.
    
    'Note that at present, scroll bars only move in single-pixel increments (in the image coordinate space), which makes our life
    ' somewhat easier.  We basically want to allow the user to scroll long enough that they can create a "mostly empty" canvas.
    ' How many pixels this requires depends on the size of the image, relative to the current canvas.
    
    'Start by calculating the *required* scroll bar maximum: the amount of the image that cannot physically fit inside the canvas.
    Dim hScrollMin As Long, hScrollMax As Long, vScrollMin As Long, vScrollMax As Long
    hScrollMax = (srcImage.Width - canvasRect_ImageCoords.Width)
    vScrollMax = (srcImage.Height - canvasRect_ImageCoords.Height)
    
    'Minimum values are easy to calculate; let the user scroll the image halfway off the screen
    hScrollMin = -1 * (canvasRect_ImageCoords.Width * 0.5)
    vScrollMin = -1 * (canvasRect_ImageCoords.Height * 0.5)
    
    'If hScrollMax or vScrollMax are negative, it means the canvas is larger (in that dimension) than the zoomed image.  When this happens,
    ' rely solely on the "halfway off screen" scroll measurement.
    If (hScrollMax > 0) Then hScrollMax = hScrollMax - hScrollMin Else hScrollMax = -hScrollMin
    If (vScrollMax > 0) Then vScrollMax = vScrollMax - vScrollMin Else vScrollMax = -vScrollMin
    
    'We now have scroll bar max/min values.  Forward them to the destination pdCanvas object.
    With dstCanvas
        .SetRedrawSuspension True
        .SetScrollMin PD_HORIZONTAL, hScrollMin
        .SetScrollMax PD_HORIZONTAL, hScrollMax
        .SetScrollMin PD_VERTICAL, vScrollMin
        .SetScrollMax PD_VERTICAL, vScrollMax
        .SetRedrawSuspension False
    End With
    
    'As a convenience to the user, we also make each scroll bar's LargeChange parameter proportional to the scroll bar's maximum value.
    If (hScrollMax > 15) And (g_Zoom.GetZoomValue(srcImage.GetZoom) <= 1) Then
        dstCanvas.SetScrollLargeChange PD_HORIZONTAL, hScrollMax \ 16
    Else
        dstCanvas.SetScrollLargeChange PD_HORIZONTAL, 1
    End If
    
    If (vScrollMax > 15) And (g_Zoom.GetZoomValue(srcImage.GetZoom) <= 1) Then
        dstCanvas.SetScrollLargeChange PD_VERTICAL, vScrollMax \ 16
    Else
        dstCanvas.SetScrollLargeChange PD_VERTICAL, 1
    End If
    
    'Scroll bars are now prepped and ready!
    
    'With all scroll bar data assembled, we have enough information to create a target back buffer.
    If (srcImage.canvasBuffer.GetDIBWidth <> canvasRect_ActualPixels.Width) Or (srcImage.canvasBuffer.GetDIBHeight <> canvasRect_ActualPixels.Height) Then
        srcImage.canvasBuffer.CreateBlank canvasRect_ActualPixels.Width, canvasRect_ActualPixels.Height, 32, g_Themer.GetGenericUIColor(UI_CanvasElement), 255
    Else
        GDI_Plus.GDIPlusFillDIBRect srcImage.canvasBuffer, 0, 0, canvasRect_ActualPixels.Width, canvasRect_ActualPixels.Height, g_Themer.GetGenericUIColor(UI_CanvasElement), 255, GP_CM_SourceCopy
    End If
    
    'Because subsequent stages of the pipeline may need all the data we've assembled, store a copy of all relevant rects
    ' inside the source pdImage object.
    With srcImage.imgViewport
        .SetCanvasRectActualPixels canvasRect_ActualPixels
        .SetCanvasRectImageCoords canvasRect_ImageCoords
        .SetImageRectCanvasCoords imageRect_CanvasCoords
    End With
    
    'The final step of this pipeline is optional.  If the user wants us to calculate specific scroll bar values, they must pass
    ' a special request enum via the function ParamArray().  At present, this class is capable of three different auto-calculations,
    ' which correspond to the three enum values of PD_VIEWPORT_SPECIAL_REQUEST
    ' VSR_ResetToZero: reset the scroll bar to (0, 0), which also centers the image when in "zoom-to-fit" mode
    ' VSR_ResetToCustom: reset the scroll bar to two values supplied by the user (in (x, y) order)
    ' VSR_AutoCenter: forcibly center the image, regardless of zoom
    ' VSR_PreservePointPosition: given a point (typically the point under the mouse cursor), preserve its before-and-after position,
    '                            even though zoom has changed!  This makes mousewheel scrolling much more intuitive.
    
    'Check for a param array now, and if none is found, skip straight to the next pipeline stage
    If specialRequestActive Then
        
        'Regardless of what type of scroll bar setting we're applying, we need to disable automatic viewport redraws.
        ' (Otherwise, changing the scroll bar value will trigger a viewport pipeline request, wreaking havoc)
        dstCanvas.SetRedrawSuspension True
        
        'The first extra setting defines the type of scroll bar handling request
        Select Case specialRequestID
        
            Case VSR_ResetToZero
                dstCanvas.SetScrollValue PD_BOTH, 0
                
            Case VSR_ResetToCustom
                dstCanvas.SetScrollValue PD_HORIZONTAL, CLng(ExtraSettings(1))
                dstCanvas.SetScrollValue PD_VERTICAL, CLng(ExtraSettings(2))
            
            'If the user has a point they want us to preserve, they will have passed two sets of coordinates:
            ' 1) The literal (x, y) of the mouse on the current canvas (e.g. the coordinates returned by a mouse event)
            ' 2) The corresponding (x, y) of that mouse position *in the image coordinate space*
            '
            'Our goal is to make that same (x, y) point on the canvas correlate to the same (x, y) on the image, regardless of any
            ' zoom/viewport/other changes we have just made in this function.
            Case VSR_PreservePointPosition
                
                Dim oldXCanvas As Single, oldYCanvas As Single, targetXImage As Single, targetYImage As Single
                oldXCanvas = CSng(ExtraSettings(1))
                oldYCanvas = CSng(ExtraSettings(2))
                targetXImage = CSng(ExtraSettings(3))
                targetYImage = CSng(ExtraSettings(4))
                
                'From the supplied coordinates, we know that image coordinate targetXImage was originally located on the canvas
                ' at position oldXCanvas.  Our goal is to make targetXImage *remain* at oldXCanvas position.
                
                'Start by converting targetX/Y/Image to the current canvas space.  This will give us NewCanvasX/Y values that describe
                ' where the coordinates lie on the *new* canvas.
                
                '...then set a fake, "translated" image rect, that is correct for the case of h/v/scroll = 0.  (Normally stage 3 of the
                ' pipeline creates a translated rect, but we have to provide one now because the canvas/image coordinate translation code
                ' relies on that rect!)
                srcImage.imgViewport.SetImageRectTranslated imageRect_CanvasCoords
                
                'With those values successfully set, we can now translate the target image coords into canvas coords, for the case of
                ' h/v/scroll = 0.
                Dim newXCanvas As Double, newYCanvas As Double
                Drawing.ConvertImageCoordsToCanvasCoords dstCanvas, srcImage, targetXImage, targetYImage, newXCanvas, newYCanvas, False
                
                'Use the difference between newCanvasX and oldCanvasX (while accounting for zoom) to determine new scroll bar values.
                dstCanvas.SetScrollValue PD_HORIZONTAL, (newXCanvas - oldXCanvas) / m_ZoomRatio
                dstCanvas.SetScrollValue PD_VERTICAL, (newYCanvas - oldYCanvas) / m_ZoomRatio
                
        End Select
        
        'Restore scroll-bar-originating viewport redraw requests
        dstCanvas.SetRedrawSuspension False
        
    End If
    
    'With our work here complete, we can pass control to the next pipeline stage.
    Stage2_CompositeAllLayers srcImage, dstCanvas, True
    
    'This stage of the pipeline has completed successfully!
    Exit Sub



ViewportPipeline_Stage1_Error:

    Select Case Err.Number
    
        'Out of memory
        Case 480
            PDMsgBox "There is not enough memory available to continue this operation.  Please free up system memory (RAM) and try again.  If the problem persists, reduce the zoom value and try again.", vbExclamation + vbOKOnly, "Out of memory"
            SetProgBarVal 0
            ReleaseProgressBar
            Message "Operation halted."
            
        'Anything else.  (Never encountered; failsafe only.)
        Case Else
            Message "Viewport rendering paused due to unexpected error (#%1)", Err.Number
            
    End Select

End Sub

'When all images have been unloaded, the temporary front buffer can also be erased to keep memory usage as low as possible.
' While not actually part of the viewport pipeline, I find it intuitive to store this function here.
Public Sub EraseViewportBuffers()
    If (Not m_FrontBuffer Is Nothing) Then
        m_FrontBuffer.EraseDIB
        Set m_FrontBuffer = Nothing
    End If
End Sub
