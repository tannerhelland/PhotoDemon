Attribute VB_Name = "Viewport_Engine"
'***************************************************************************
'Viewport Handler - builds and draws the image viewport and associated scroll bars
'Copyright 2001-2014 by Tanner Helland
'Created: 4/15/01
'Last updated: 30/May/14
'Last update: add support for "preserve relative canvas position under cursor while mousewheel zooming"
'
'Module for handling the image viewport.  The render pipeline works as follows:
' - Viewport_Engine.Stage1_InitializeBuffer: for recalculating all viewport variables and controls (done only when the zoom value is changed or a new picture is loaded)
' - Viewport_Engine.Stage2_CompositeAllLayers: when the viewport is scrolled (minimal redrawing is done, since the zoom value hasn't changed)
' - Viewport_Engine.Stage3_CompositeCanvas: perform any final compositing, such as the Selection Tool effect, then draw the viewport on-screen
'
'PhotoDemon is intelligent about calling the lowest routine in the pipeline, which helps it render the viewport quickly
' regardless of zoom or scroll values.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Some viewport pipeline calculations are stored at module-level.  This is more efficient than passing them as parameters, because PD's
' viewport pipeline has been deliberately created as an "out of order" pipeline.  Different user interations can trigger execution
' of the pipeline at different stages, which is crucial for maximizing viewport performance.

'As such, it is important that pipeline functions are *very cautious* about whether they read or actually modify these values.
' INTERACT WITH CAUTION.

'Width and height values of the image AFTER zoom has been applied.  (For example, if the image is 100x100
' and the zoom value is 200%, m_ImageWidthZoomed and m_ImageHeightZoomed will be 200.)
Private m_ImageWidthZoomed As Double, m_ImageHeightZoomed As Double

'These variables represent the source width - e.g. the size of the viewable picture box, divided by the zoom coefficient
Private srcWidth As Double, srcHeight As Double

'The ZoomVal value is the actual coefficient for the current zoom value.  (For example, 0.50 for "50% zoom")
Private m_ZoomRatio As Double

'These variables are the offset, as determined by the scroll bar values
Private srcX As Long, srcY As Long

'frontBuffer holds the final composited image, including any overlays (like selections)
Private frontBuffer As pdDIB

'cornerFix holds a small gray box that is copied over the corner between the horizontal and vertical scrollbars, if they exist
Private cornerFix As pdDIB

'Stage4_FlipBufferAndDrawUI is the final stage of the viewport pipeline.  It will flip the composited canvas image to the
' destination pdCanvas object, and apply any final UI elements as well - control nodes, custom cursors, etc.  This step is
' very fast, and should be used whenever full compositing is unnecessary.
'
'As part of the buffer flip, this stage will also activate and apply color management to the completed front buffer.
' (Still TODO is fixing the canvas to not rely on AutoRedraw, which will spare us having to re-activate color management on every draw.)
Public Sub Stage4_FlipBufferAndDrawUI(ByRef srcImage As pdImage, ByRef dstCanvas As pdCanvas)

    'If no images have been loaded, clear the canvas and exit
    If g_OpenImageCount = 0 Then
        FormMain.mainCanvas(0).clearCanvas
        Exit Sub
    End If

    'Make sure the canvas is valid
    If dstCanvas Is Nothing Then Exit Sub
    
    'If the image associated with this form is inactive, ignore this request
    If Not srcImage.IsActive Then Exit Sub
    
    'Because AutoRedraw can cause the form's DC to change without warning, we must re-apply color management settings any time
    ' we redraw the screen.  I do not like this any more than you do, but we risk losing our DC's settings otherwise.
    If Not (g_UseSystemColorProfile And g_IsSystemColorProfileSRGB) Then
        assignDefaultColorProfileToObject dstCanvas.hWnd, dstCanvas.hDC
        turnOnColorManagementForDC dstCanvas.hDC
    End If
    
    'Finally, flip the front buffer to the screen
    BitBlt dstCanvas.hDC, 0, srcImage.imgViewport.getTopOffset, frontBuffer.getDIBWidth, frontBuffer.getDIBHeight, frontBuffer.getDIBDC, 0, 0, vbSrcCopy
    
    
    'Finally, we can do some tool-specific rendering directly onto the form.
    Select Case g_CurrentTool
    
        'The nav tool provides two render options at present: draw layer borders, and draw layer transform nodes
        Case NAV_MOVE
        
            'If the user has requested visible layer borders, draw them now
            If CBool(toolbar_Options.chkLayerBorder) Then
                
                'Draw layer borders
                Drawing.drawLayerBoundaries srcImage.getActiveLayerIndex
                
            End If
            
            'If the user has requested visible transformation nodes, draw them now
            If CBool(toolbar_Options.chkLayerNodes) Then
                
                'Draw layer nodes
                Drawing.drawLayerNodes srcImage.getActiveLayerIndex
                
            End If
        
        'Selections are always rendered onto the canvas.  If a selection is active AND a selection tool is active, we can also
        ' draw transform nodes around the selection area.
        Case SELECT_RECT, SELECT_CIRC, SELECT_LINE, SELECT_POLYGON, SELECT_WAND
            
            'Next, check to see if a selection is active and transformable.  If it is, draw nodes around the selected area.
            If srcImage.selectionActive Then
                srcImage.mainSelection.renderTransformNodes srcImage, dstCanvas, srcImage.imgViewport.targetLeft, srcImage.imgViewport.targetTop
            End If
        
    End Select
    
    'FYI, in the future, any additional UI compositing can be handled here.
    
    'With all rendering complete, copy the form's image into the .Picture (e.g. render it on-screen) and refresh
    dstCanvas.requestBufferSync

End Sub

'Stage3_CompositeCanvas takes the current canvas (which has a checkerboard and fully layered image drawn atop it) and applies a few
' other frills to it, including things like canvas decorations (e.g. drop-shadows, a fix for the space between scroll bars), and the
' current selection, if one is active.  This stage is the final stage before color-management is applied, so it's important to render
' any color-specific bits now, as the next stage will apply color-management processing to whatever is contained in the front buffer.
'
'When this stage is finished, the srcImage.frontBuffer object will contain a screen-ready copy of the canvas, with the fully
' composited image drawn atop a checkerboard in the viewport section of the canvas.  Standard canvas decorations will also be present,
' provided that the user's performance settings allow them.
'
'After this stage, the only things that should be rendered onto the canvas are uncolored UI elements, like custom-drawn cursors or
' control nodes.
Public Sub Stage3_CompositeCanvas(ByRef srcImage As pdImage, ByRef dstCanvas As pdCanvas)

    'If no images have been loaded, clear the canvas and exit
    If g_OpenImageCount = 0 Then
        FormMain.mainCanvas(0).clearCanvas
        Exit Sub
    End If

    'Make sure the canvas is valid
    If dstCanvas Is Nothing Then Exit Sub
    
    'If the image associated with this form is inactive, ignore this request
    If Not srcImage.IsActive Then Exit Sub

    'Create the front buffer as necessary
    If frontBuffer Is Nothing Then Set frontBuffer = New pdDIB
        
    If (frontBuffer.getDIBWidth <> srcImage.backBuffer.getDIBWidth) Or (frontBuffer.getDIBHeight <> srcImage.backBuffer.getDIBHeight) Then
        frontBuffer.createFromExistingDIB srcImage.backBuffer
    Else
        BitBlt frontBuffer.getDIBDC, 0, 0, srcImage.backBuffer.getDIBWidth, srcImage.backBuffer.getDIBHeight, srcImage.backBuffer.getDIBDC, 0, 0, vbSrcCopy
    End If
        
    
    'If the user's performance preferences allow for interface decorations, render them next
    If g_InterfacePerformance <> PD_PERF_FASTEST Then
    
        'We'll handle this in two steps; first, render the horizontal shadows
        If Not dstCanvas.getScrollVisibility(PD_VERTICAL) Then
            
            'Make sure the image isn't snugly fit inside the viewport; if it is, rendering drop shadows is a waste of time
            If srcImage.imgViewport.targetTop <> 0 Then
                'Top edge
                StretchBlt frontBuffer.getDIBDC, srcImage.imgViewport.targetLeft, srcImage.imgViewport.targetTop - PD_CANVASSHADOWSIZE, srcImage.imgViewport.targetWidth, PD_CANVASSHADOWSIZE, g_CanvasShadow.getShadowDC(0), 0, 0, 1, PD_CANVASSHADOWSIZE, vbSrcCopy
                'Bottom edge
                StretchBlt frontBuffer.getDIBDC, srcImage.imgViewport.targetLeft, srcImage.imgViewport.targetTop + srcImage.imgViewport.targetHeight, srcImage.imgViewport.targetWidth, PD_CANVASSHADOWSIZE, g_CanvasShadow.getShadowDC(1), 0, 0, 1, PD_CANVASSHADOWSIZE, vbSrcCopy
            End If
        
        End If
        
        'Second, the vertical shadows
        If Not dstCanvas.getScrollVisibility(PD_HORIZONTAL) Then
                    
            'Make sure the image isn't snugly fit inside the viewport; if it is, this is a waste of time
            If srcImage.imgViewport.targetLeft <> 0 Then
                'Left edge
                StretchBlt frontBuffer.getDIBDC, srcImage.imgViewport.targetLeft - PD_CANVASSHADOWSIZE, srcImage.imgViewport.targetTop, PD_CANVASSHADOWSIZE, srcImage.imgViewport.targetHeight, g_CanvasShadow.getShadowDC(2), 0, 0, PD_CANVASSHADOWSIZE, 1, vbSrcCopy
                'Right edge
                StretchBlt frontBuffer.getDIBDC, srcImage.imgViewport.targetLeft + srcImage.imgViewport.targetWidth, srcImage.imgViewport.targetTop, PD_CANVASSHADOWSIZE, srcImage.imgViewport.targetHeight, g_CanvasShadow.getShadowDC(3), 0, 0, PD_CANVASSHADOWSIZE, 1, vbSrcCopy
            End If
        
        End If
        
        'Finally, the corners, which are only drawn if both scroll bars are invisible
        If (Not dstCanvas.getScrollVisibility(PD_HORIZONTAL)) And (Not dstCanvas.getScrollVisibility(PD_VERTICAL)) Then
        
            'NW corner
            StretchBlt frontBuffer.getDIBDC, srcImage.imgViewport.targetLeft - PD_CANVASSHADOWSIZE, srcImage.imgViewport.targetTop - PD_CANVASSHADOWSIZE, PD_CANVASSHADOWSIZE, PD_CANVASSHADOWSIZE, g_CanvasShadow.getShadowDC(4), 0, 0, PD_CANVASSHADOWSIZE, PD_CANVASSHADOWSIZE, vbSrcCopy
            'NE corner
            StretchBlt frontBuffer.getDIBDC, srcImage.imgViewport.targetLeft + srcImage.imgViewport.targetWidth, srcImage.imgViewport.targetTop - PD_CANVASSHADOWSIZE, PD_CANVASSHADOWSIZE, PD_CANVASSHADOWSIZE, g_CanvasShadow.getShadowDC(5), 0, 0, PD_CANVASSHADOWSIZE, PD_CANVASSHADOWSIZE, vbSrcCopy
            'SW corner
            StretchBlt frontBuffer.getDIBDC, srcImage.imgViewport.targetLeft - PD_CANVASSHADOWSIZE, srcImage.imgViewport.targetTop + srcImage.imgViewport.targetHeight, PD_CANVASSHADOWSIZE, PD_CANVASSHADOWSIZE, g_CanvasShadow.getShadowDC(6), 0, 0, PD_CANVASSHADOWSIZE, PD_CANVASSHADOWSIZE, vbSrcCopy
            'SE corner
            StretchBlt frontBuffer.getDIBDC, srcImage.imgViewport.targetLeft + srcImage.imgViewport.targetWidth, srcImage.imgViewport.targetTop + srcImage.imgViewport.targetHeight, PD_CANVASSHADOWSIZE, PD_CANVASSHADOWSIZE, g_CanvasShadow.getShadowDC(7), 0, 0, PD_CANVASSHADOWSIZE, PD_CANVASSHADOWSIZE, vbSrcCopy
        
        End If
    
    End If
    
    'If both scrollbars are active, copy a gray square over the small space between them
    If dstCanvas.getScrollVisibility(PD_HORIZONTAL) And dstCanvas.getScrollVisibility(PD_VERTICAL) Then
        
        'Only initialize the corner fix image once
        If cornerFix Is Nothing Then
            Set cornerFix = New pdDIB
            cornerFix.createBlank dstCanvas.getScrollWidth(PD_VERTICAL), dstCanvas.getScrollHeight(PD_HORIZONTAL), 24, vbButtonFace
        End If
        
        'Draw the square over any exposed parts of the image in the bottom-right of the image, between the scroll bars
        BitBlt dstCanvas.hDC, dstCanvas.getScrollLeft(PD_VERTICAL), dstCanvas.getScrollTop(PD_HORIZONTAL), cornerFix.getDIBWidth, cornerFix.getDIBHeight, cornerFix.getDIBDC, 0, 0, vbSrcCopy
        
    End If
    
    'Check to see if a selection is active.
    If srcImage.selectionActive Then
    
        'If it is, composite the selection against the front buffer
        srcImage.mainSelection.renderCustom frontBuffer, srcImage, dstCanvas, srcImage.imgViewport.targetLeft, srcImage.imgViewport.targetTop, srcImage.imgViewport.targetWidth, srcImage.imgViewport.targetHeight, toolbar_Options.cboSelRender.ListIndex, toolbar_Options.csSelectionHighlight.Color
    
    End If
        
    'Pass the completed front buffer to the final stage of the pipeline, which will flip everything to the screen and render any
    ' remaining UI elements!
    Stage4_FlipBufferAndDrawUI srcImage, dstCanvas
    
    
End Sub

'Stage2_CompositeAllLayers is used to composite the current image onto the source pdImage's back buffer.  This function does not
' perform any initialization or pre-rendering checks, so you cannot use it if zoom is changed, or if the viewport area has changed.
' (Stage1_InitializeBuffer is used for that.)  The most common use-case for this function is the use of scrollbars, or non-destructive
' layer changes that require a recomposite of the image, but not a full recreation calculation of the viewport and canvas buffers.
Public Sub Stage2_CompositeAllLayers(ByRef srcImage As pdImage, ByRef dstCanvas As pdCanvas)
        
    'Like the previous stage of the pipeline, we start by performing a number of "do not render the viewport at all" checks.
    
    'First, and most obvious, is to exit now if the public g_AllowViewportRendering parameter has been forcibly disabled.
    If Not g_AllowViewportRendering Then Exit Sub
    
    'Make sure the target canvas is valid
    If dstCanvas Is Nothing Then Exit Sub
    
    'If the pdImage object associated with this form is inactive, ignore this request
    If Not srcImage.IsActive Then Exit Sub
    
    'This function can return timing reports if desired; at present, this is automatically activated in PRE-ALPHA and ALPHA builds,
    ' but disabled for BETA and PRODUCTION builds; see the LoadTheProgram() function for details.
    Dim startTime As Double
    If g_DisplayTimingReports Then startTime = Timer
    
    'Stage 1 of the pipeline (Stage1_InitializeBuffer) prepared srcImage.BackBuffer for us.  The goal of this stage is two-fold:
    ' 1) Fill the viewport area of the canvas with a checkerboard pattern
    ' 2) Render the fully composited image atop the checkerboard pattern
    
    'Note that the imgCompositor object will handle most of this stage for us, as it performs the actual compositing.
    
    'These variables represent the source width - e.g. the size of the viewable picture box, divided by the zoom coefficient.
    ' Because rounding errors may occur with certain image sizes, we apply a special check when zoom = 100.
    If srcImage.currentZoomValue = g_Zoom.getZoom100Index Then
        srcWidth = srcImage.imgViewport.targetWidth
        srcHeight = srcImage.imgViewport.targetHeight
    Else
        srcWidth = srcImage.imgViewport.targetWidth / m_ZoomRatio
        srcHeight = srcImage.imgViewport.targetHeight / m_ZoomRatio
    End If
        
    'These variables are the offset into the source image, as determined by the scroll bar's values.  PD supports partial
    ' compositing of a given region of the image.  This allows for excellent performance when the image is larger than the
    ' available screen real-estate (as we don't waste time compositing invisible regions).
    If dstCanvas.getScrollVisibility(PD_HORIZONTAL) Then srcX = dstCanvas.getScrollValue(PD_HORIZONTAL) Else srcX = 0
    If dstCanvas.getScrollVisibility(PD_VERTICAL) Then srcY = dstCanvas.getScrollValue(PD_VERTICAL) Else srcY = 0
        
    'Before rendering the image, apply a checkerboard pattern to the viewport region of the source image's back buffer.
    ' TODO: cache g_CheckerboardPattern persistently, in GDI+ format, so we don't have to recreate it on every draw.
    With srcImage.imgViewport
        GDI_Plus.GDIPlusFillDIBRect_Pattern srcImage.backBuffer, .targetLeft, .targetTop, .targetWidth, .targetHeight, g_CheckerboardPattern
    End With
    
    'As a failsafe, perform a GDI+ check.  PD probably won't work at all without GDI+, so I could look at dropping this check
    ' in the future... but for now, we leave it, just in case.
    If g_GDIPlusAvailable Then
        
        'We can now use PD's amazing rect-specific compositor to retrieve only the relevant section of the current viewport.
        ' Note that we request our own interpolation mode, and we determine this based on the user's viewport performance preference.
        ' (TODO: consider exposing bilinear interpolation as an option, which is blurrier, but doesn't suffer from the defects of
        '        GDI+'s preprocessing, which screws up subpixel positioning.)
        
        'When we've been asked to maximize performance, use nearest neighbor for all zoom modes
        If g_ViewportPerformance = PD_PERF_FASTEST Then
            srcImage.getCompositedRect srcImage.backBuffer, srcImage.imgViewport.targetLeft, srcImage.imgViewport.targetTop, srcImage.imgViewport.targetWidth, srcImage.imgViewport.targetHeight, srcX, srcY, srcWidth, srcHeight, InterpolationModeNearestNeighbor
            
        'Otherwise, switch dynamically between high-quality and low-quality interpolation depending on the current zoom.
        ' Note that the compositor will perform some additional checks, and if the image is zoomed-in, it will switch to nearest-neighbor
        ' automatically (regardless of what method we request).
        Else
            srcImage.getCompositedRect srcImage.backBuffer, srcImage.imgViewport.targetLeft, srcImage.imgViewport.targetTop, srcImage.imgViewport.targetWidth, srcImage.imgViewport.targetHeight, srcX, srcY, srcWidth, srcHeight, IIf(m_ZoomRatio <= 1, InterpolationModeHighQualityBicubic, InterpolationModeNearestNeighbor)
        End If
                
    'This is an emergency fallback, only.  PD won't work without GDI+, so rendering the viewport is pointless.
    Else
        Message "WARNING!  GDI+ could not be found.  (PhotoDemon requires GDI+ for proper program operation.)"
    End If
    
    'Pass control to the next stage of the pipeline.
    Stage3_CompositeCanvas srcImage, dstCanvas
    
    'If timing reports are enabled, we report them after the rest of the pipeline has finished.
    If g_DisplayTimingReports Then Debug.Print "Viewport render timing: " & Format(CStr((Timer - startTime) * 1000), "0000.00") & " ms"
    
End Sub

'Per its name, Stage1_InitializeBuffer is responsible for preparing a bunch of math related to viewport rendering.  Its duties include:
    '1) Calculating all zoom-related math
    '2) Determining whether scroll bars are required, and if they are, what their max/min values should be
    '3) Canvas offsets, if the image is zoomed out far enough that dead space is present in the viewport.
    
'This function is crucial, because all subsequent pipeline operations operate on the values determined by this function.

'Because this function does no actual rendering - only preparation math - it only needs to be executed under specific conditions,
' namely when:
    '1) an image is first loaded
    '2) the viewport's zoom value is changed
    '3) the main PhotoDemon window is resized
    '4) toolbars are hidden or shown (similar to resizing, this changes available viewport area)
    '5) edits that modify an image's size (resizing, rotating, etc - basically anything that changes the size of the back buffer)

'Because the full rendering pipeline must be executed when this function is called, it is considered highly expensive, even though
' the math it performs is relatively quick.  To help cut down on overuse of this function (e.g. sloppy pipeline requests), an optional
' "reasonForRedraw" parameter is used.  This untranslated string, supplied by the caller, has proven helpful while optimizing.
' Similarly, if you see a bunch of Stage1_InitializeBuffer requests happening back-to-back in the Debug window, you should investigate,
' because such operations are likely hurting performance.

'While this function is primarily concerned with the math required to handle zoom and scroll operations correctly, there are a few
' additional parameters that are occasionally necessary.  "Zoom to Coordinate" behavior, used when the mousewheel is applied while
' over a specific pixel, will pass targetX and targetY parameters to the function.  If present, Stage1_InitializeBuffer will automatically
' set the scroll bar values after its calculations are complete, in a way that preserves the on-screen position of the passed
' coordinate.  (Note that it does this as closely as it can, but some zoom changes make this impossible, such as zooming out to a
' point where scroll bars are no longer visible).

'As an important follow-up note, two sets of target coordinates must be passed for this capability to work: one set of coordinates
' in *canvas space*, and one set in *image space*.  Both are required, because Stage1_InitializeBuffer doesn't keep track of past zoom values.
' This means that once the viewport's zoom level has been changed (as will likely happen prior to calling this function, by mousewheel),
' this function does know what the prior zoom level was - and a single set of coordinates is not enough to maintain image positioning.

'Thus, when making use of "zoom to coordinate" behavior, you must handle zoom changes in the following order:
' 1) cache x/y coordinate values in two coordinate spaces: image and canvas
' 2) disable automatic canvas redraws
' 3) change the zoom value; this allows the zoom engine to reconstruct conditional values, like "fit to window"
' 4) re-enable automatic canvas redraws (this can happen now, or after step 5 - just don't forget to do it!)
' 5) request a manual redraw via Stage1_InitializeBuffer, and be sure supply the previously cached x/y values
Public Sub Stage1_InitializeBuffer(ByRef srcImage As pdImage, ByRef dstCanvas As pdCanvas, Optional ByRef reasonForRedraw As String, Optional ByVal oldXCanvas As Long = 0, Optional ByVal oldYCanvas As Long = 0, Optional ByVal targetXImage As Double = 0, Optional ByVal targetYImage As Double = 0)
    
    
    On Error GoTo ViewportPipeline_Stage1_Error
    
    
    'This initial stage of the pipeline contains a number of "do not render the viewport at all" checks.
    
    'First, and most obvious, is to exit now if the public g_AllowViewportRendering parameter has been forcibly disabled.
    ' (Detailed explanation: this routine is automatically triggered by the main window's resize notifications.  When new images
    '  are loaded, the image tabstrip will likely appear, which in turn changes the available viewport space, just like a resize
    '  event.  To prevent  this behavior from triggering multiple Stage1_InitializeBuffer requests, g_AllowViewportRendering is
    '  utilized.)
    If Not g_AllowViewportRendering Then Exit Sub
    
    'Second, exit if the destination canvas has not been initialized yet; this can happen during program initialization.
    If dstCanvas Is Nothing Then Exit Sub
    
    'Third, exit if no images have been loaded.  The canvas will take care of rendering a blank viewport.
    If g_OpenImageCount = 0 Then
        FormMain.mainCanvas(0).clearCanvas
        Exit Sub
    End If
    
    'Fourth, exit if the source image is invalid.
    If srcImage Is Nothing Then Exit Sub
    
    'Fifth, if the source image is inactive (e.g. it has been unloaded at some point in the past), do not execute a redraw.
    ' (For performance reasons, PD does not shrink its primary pdImages() array unless required due to memory pressure.
    '  Instead, it just deactivates entries by marking the .IsActive property - so that property must be considered
    '  prior to executing events for an image.)
    If Not srcImage.IsActive Then Exit Sub
    
    
    'If we made it all the way here, the viewport pipeline needs to be executed.
    
    
    'We will be referencing the source pdImage object many times.  To improve performance, cache its ID value
    Dim curImage As Long
    curImage = srcImage.imageID
    
    'Because this routine is time-consuming, I carefully track its usage to try and minimize how frequently it's called.
    ' Feel free to comment out this line if you don't find it helpful.
    Debug.Print "Preparing viewport: " & reasonForRedraw & " | (" & curImage & ") "
    
    'This crucial value is the mathematical ratio of the current zoom value: 1 for 100%, 0.5 for 50%, 2 for 200%, etc.
    ' We can't generate this automatically, because specialty zoom values (like "fit to window") must be externally generated
    ' by PD's zoom handler.
    m_ZoomRatio = g_Zoom.getZoomValue(srcImage.currentZoomValue)
    
    'The fundamental problem this first pipeline stage must solve is: how much screen real-estate do we have to work with, and how
    ' must we fit the image into that real-state.  It quickly becomes complicated because some decisions we make will actually
    ' change the available real-estate (e.g. enabling a vertical scrollbar reduces horizontal real-estate, requiring a re-calculation
    ' of any horizontal data up to that point).
    
    'Also problematic is the potential of future feature additions, like rulers, that also interfere with our available screen
    ' real-estate.  To try and preempt the changes required by such features, you'll notice various "offsets" used prior to
    ' calculating image positioning.  These may not do anything at present, so don't worry if they go unused.
    
    'Another important clarification is use of the terms "viewport" and "canvas".
    
    ' Viewport = the area of the screen dedicated to just the image
    ' Canvas = the area of the screen dedicated to the canvas, and any surrounding dead space (relevant when zoomed out)
    
    'Sometimes these two rectangles will be identical.  Sometimes they will not.  If the canvas rect is larger than
    ' the viewport rect, the viewport rect will automatically be moved so that it is centered within the viewport area.
    ' (This behavior will need to be modified in the future, to allow for scrolling past canvas edges.)
    
    'Calculate the width and height of a full-size viewport based on the current zoom value
    m_ImageWidthZoomed = (srcImage.Width * m_ZoomRatio)
    m_ImageHeightZoomed = (srcImage.Height * m_ZoomRatio)
    
    'Calculate the vertical offset of the viewport.  While not relevant at present, it will someday be necessary to allow
    ' for rulers.
    Dim verticalOffset As Long
    verticalOffset = srcImage.imgViewport.getVerticalOffset
    
    'Grab the canvas dimensions; note that these are just thin wrappers to the .ScaleWidth and .ScaleHeight properties
    ' of the control.
    Dim canvasWidth As Long, canvasHeight As Long
    canvasWidth = dstCanvas.getCanvasWidth
    canvasHeight = dstCanvas.getCanvasHeight - verticalOffset
    
    'These variables will reflect whether or not scroll bars are enabled; this is used rather than the .Enabled property so we
    ' can defer rendering the scroll bars until the last possible instant (rather than turning them on-and-off mid-subroutine).
    Dim hScrollEnabled As Boolean, vScrollEnabled As Boolean
    hScrollEnabled = False
    vScrollEnabled = False
    
    'Step 1: compare zoomed image width to canvas width.  If the zoomed image width is larger, we need to enable a horizontal
    ' scroll bar.  Also, because fractional zoom values are allowed, Int() is used to clamp.
    If Int(m_ImageWidthZoomed) > canvasWidth Then hScrollEnabled = True
    
    'Step 2: repeat Step 1, but in the vertical direction.  Note that we must subtract the horizontal scroll bar's height, if
    ' it was enabled by step 1.
    If (Int(m_ImageHeightZoomed) > canvasHeight) Then
        vScrollEnabled = True
    Else
        If hScrollEnabled And (Int(m_ImageHeightZoomed) > (canvasHeight - dstCanvas.getScrollHeight(PD_HORIZONTAL))) Then
            vScrollEnabled = True
        End If
    End If
        
    'Step 3: one last check on horizontal viewport width; if the vertical scrollbar was enabled by step 2, the horizontal
    ' viewport width has changed.
    If vScrollEnabled And (Not hScrollEnabled) Then
        If (Int(m_ImageWidthZoomed) > (canvasWidth - dstCanvas.getScrollWidth(PD_VERTICAL))) Then hScrollEnabled = True
    End If
    
    'We now know which scroll bars need to be enabled, which allows us to finalize the viewport's position and size.
    ' (Remember that the viewport's position must be changed if either dimension is smaller than the canvas area, as we
    '  need to center it inside the canvas.)
    Dim viewportLeft As Double, viewportTop As Double
    Dim viewportWidth As Double, viewportHeight As Double
    
    'These nested If statements basically cover the case of neither or one or both scroll bars being active.
    If hScrollEnabled Then
        viewportLeft = 0
        If Not vScrollEnabled Then
            viewportWidth = canvasWidth
        Else
            viewportWidth = canvasWidth - dstCanvas.getScrollWidth(PD_VERTICAL)
        End If
    Else
        viewportWidth = m_ImageWidthZoomed
        If Not vScrollEnabled Then
            viewportLeft = (canvasWidth - m_ImageWidthZoomed) / 2
        Else
            viewportLeft = ((canvasWidth - dstCanvas.getScrollWidth(PD_VERTICAL)) - m_ImageWidthZoomed) / 2
        End If
    End If
    
    If vScrollEnabled Then
        viewportTop = 0
        If Not hScrollEnabled Then
            viewportHeight = canvasHeight
        Else
            viewportHeight = canvasHeight - dstCanvas.getScrollHeight(PD_HORIZONTAL)
        End If
    Else
        viewportHeight = m_ImageHeightZoomed
        If Not hScrollEnabled Then
            viewportTop = (canvasHeight - m_ImageHeightZoomed) / 2
        Else
            viewportTop = ((canvasHeight - dstCanvas.getScrollHeight(PD_HORIZONTAL)) - m_ImageHeightZoomed) / 2
        End If
    End If
    
    'Now we know a whole bunch of things:
    ' 1) which scrollbars, if any, are enabled
    ' 2) the position of the viewport (image portion of the canvas)
    ' 3) the size of the viewport.
    
    'From these three things, we can now calculate scroll bar maximum values, if necessary.
    
    'First, however, let's cover the case of "no scroll bars are enabled."  When that happens, this function's work is
    ' already complete, so we can advance to the next stage of the pipeline!
    If (Not hScrollEnabled) And (Not vScrollEnabled) Then
    
        'Reset the scroll bar values to zero; this allows future pipeline stages to shortcut some calculations
        dstCanvas.setRedrawSuspension True
        dstCanvas.setScrollValue PD_BOTH, 0
        dstCanvas.setRedrawSuspension False
    
        'If the scroll bars are currently visible, hide 'em.  Note that the canvas itself will determine whether the
        ' primary canvas area needs to be moved as a result of this.
        dstCanvas.setScrollVisibility PD_BOTH, False
        
        'Resize the back buffer and store the relevant painting information into the passed pdImages() object.
        ' (TODO: roll the canvas color over to the central themer.)
        srcImage.backBuffer.createBlank canvasWidth, canvasHeight, 24, g_CanvasBackground
        srcImage.imgViewport.targetLeft = viewportLeft
        srcImage.imgViewport.targetTop = viewportTop
        srcImage.imgViewport.targetWidth = viewportWidth
        srcImage.imgViewport.targetHeight = viewportHeight
        
        'Pass control to the next stage of the pipeline
        Viewport_Engine.Stage2_CompositeAllLayers srcImage, dstCanvas
        
    
    'This Else() bracket covers the case of one or both viewport scroll bars being enabled.  Inside this block, we will calculate
    ' the scrollbar's maximum values, and if zoom-to-position is used, we will also calculate the scrollbar's values.
    Else
    
        Dim newScrollMax As Long
        Dim newXCanvas As Double, newYCanvas As Double, canvasXDiff As Double, canvasYDiff As Double
        
        'We are now going to set a bunch of scroll bar properties, all at once.  These changes may cause the scroll bars
        ' to initiate a pipeline request - to prevent that, we forcibly disable screen refreshes in advance.
        dstCanvas.setRedrawSuspension True
        
        'Horizontal scroll bar is processed first.
        If hScrollEnabled Then
            
            'If zoomed out, set the scroll bar range to the number of not-visible pixels.  This will result in sub-pixel scrolling
            ' if the scrollbar is clicked-and-held.
            If m_ZoomRatio <= 1 Then
                newScrollMax = srcImage.Width - Int(viewportWidth * g_Zoom.getZoomOffsetFactor(srcImage.currentZoomValue) + 0.5)
                
            'If zoomed-out, we must divide by the zoom factor (instead of multiplying by it).  This allows us to scroll by integer
            ' pixel values, which is more convenient, especially at massive zoom levels.
            Else
                newScrollMax = srcImage.Width - Int(viewportWidth / g_Zoom.getZoomOffsetFactor(srcImage.currentZoomValue) + 0.5)
                
            End If
            
            'Set the new maximum value
            dstCanvas.setScrollMax PD_HORIZONTAL, newScrollMax
            
            'If the calling function supplied a targetX value, we will use that to calculate the theoretical scroll bar value that
            ' maintains the position of that pixel on the screen.
            ' (Note: I call the value "theoretical", because it may lie outside the range of the scroll bar.  If this happens, PD's
            '  custom scroll bar class will automatically bring the value in-bounds.)
            If oldXCanvas <> 0 Then
                
                'From the supplied coordinates, we know that image coordinate targetXImage was originally located at position
                ' oldXCanvas.  Our goal is to make targetXImage *remain* at oldXCanvas position, while accounting for
                ' any changes made to zoom (and thus to scroll bar max/min values).
                
                'Start by converting targetXCanvas to the current canvas space.  This will give us a value NewCanvasX, that describes
                ' where that coordinate lies on the *new* canvas.
                dstCanvas.setScrollValue PD_HORIZONTAL, 0
                Drawing.convertImageCoordsToCanvasCoords FormMain.mainCanvas(0), pdImages(g_CurrentImage), targetXImage, targetYImage, newXCanvas, newYCanvas, False
                
                'Use the difference between newCanvasX and oldCanvasX to determine a new scroll bar value.
                canvasXDiff = newXCanvas - oldXCanvas
                
                'Modify the scrollbar by canvasXDiff amount, while accounting for zoom (as different zoom levels cause scroll bar
                ' notches to represent varying amounts of pixels)
                dstCanvas.setScrollValue PD_HORIZONTAL, canvasXDiff / g_Zoom.getZoomValue(srcImage.currentZoomValue)
                
            End If
                            
            'As a convenience to the user, make the scroll bar's LargeChange parameter proportional to the scroll bar's new
            ' maximum value.
            If (dstCanvas.getScrollMax(PD_HORIZONTAL) > 15) And (g_Zoom.getZoomValue(srcImage.currentZoomValue) <= 1) Then
                dstCanvas.setScrollLargeChange PD_HORIZONTAL, dstCanvas.getScrollMax(PD_HORIZONTAL) \ 16
            Else
                dstCanvas.setScrollLargeChange PD_HORIZONTAL, 1
            End If
            
        End If
        
        'Now repeat all of the above steps, but for the vertical scroll bar.
        If vScrollEnabled Then
            
            If m_ZoomRatio <= 1 Then
                newScrollMax = srcImage.Height - Int(viewportHeight * g_Zoom.getZoomOffsetFactor(srcImage.currentZoomValue) + 0.5)
            Else
                newScrollMax = srcImage.Height - Int(viewportHeight / g_Zoom.getZoomOffsetFactor(srcImage.currentZoomValue) + 0.5)
            End If
            
            dstCanvas.setScrollMax PD_VERTICAL, newScrollMax
            
            If oldYCanvas <> 0 Then
                
                dstCanvas.setScrollValue PD_VERTICAL, 0
                
                Drawing.convertImageCoordsToCanvasCoords FormMain.mainCanvas(0), pdImages(g_CurrentImage), targetXImage, targetYImage, newXCanvas, newYCanvas, False
                canvasYDiff = newYCanvas - oldYCanvas
                dstCanvas.setScrollValue PD_VERTICAL, canvasYDiff / g_Zoom.getZoomValue(srcImage.currentZoomValue)
                
            End If
            
            If (dstCanvas.getScrollMax(PD_VERTICAL) > 15) And (g_Zoom.getZoomValue(srcImage.currentZoomValue) <= 1) Then
                dstCanvas.setScrollLargeChange PD_VERTICAL, dstCanvas.getScrollMax(PD_VERTICAL) \ 16
            Else
                dstCanvas.setScrollLargeChange PD_VERTICAL, 1
            End If
            
        End If
        
        
        'At this point, scroll bar max values are now properly set.
        
        
        'It is now time to display the scroll bars, if they aren't displayed already.  As part of this step, we may also need
        ' to *hide* the scroll bars if they were previously visible, but aren't now.
        If hScrollEnabled Then
            dstCanvas.moveScrollBar PD_HORIZONTAL, 0, canvasHeight - dstCanvas.getScrollHeight(PD_HORIZONTAL), viewportWidth, dstCanvas.getScrollHeight(PD_HORIZONTAL)
            dstCanvas.setScrollVisibility PD_HORIZONTAL, True
        Else
            
            'If the scroll bar is being hidden, set its value to 0.  This allows subsequent pipeline stages to skip some steps.
            dstCanvas.setScrollValue PD_HORIZONTAL, 0
            dstCanvas.setScrollVisibility PD_HORIZONTAL, False
            
        End If
        
        'Repeat the above steps for the vertical scroll bar
        If vScrollEnabled Then
            dstCanvas.moveScrollBar PD_VERTICAL, canvasWidth - dstCanvas.getScrollWidth(PD_VERTICAL), srcImage.imgViewport.getTopOffset, dstCanvas.getScrollWidth(PD_VERTICAL), viewportHeight
            dstCanvas.setScrollVisibility PD_VERTICAL, True
        Else
            dstCanvas.setScrollValue PD_VERTICAL, 0
            dstCanvas.setScrollVisibility PD_VERTICAL, False
        End If
        
        
        'With all major UI elements now positioned and updated, we can re-enable automatic viewport pipeline requests
        dstCanvas.setRedrawSuspension False
        
        'This pipeline stage is pretty much complete.  All that's left to do is intializing this pdImage's back buffer to its new size,
        ' and caching all relevant viewport measurements (as subsequent stages need them).
        
        'Prepare the back buffer.  Note that we can shrink it slightly if scroll bars are active.
        Dim finalCanvasWidth As Long, finalCanvasHeight As Long
        If vScrollEnabled Then finalCanvasWidth = canvasWidth - dstCanvas.getScrollWidth(PD_VERTICAL) Else finalCanvasWidth = canvasWidth
        If hScrollEnabled Then finalCanvasHeight = canvasHeight - dstCanvas.getScrollHeight(PD_HORIZONTAL) Else finalCanvasHeight = canvasHeight
        
        'Testing shows no measurable difference between a 32-bit or 24-bit back buffer.  I am going to try 24-bit for now,
        ' but you can easily swap in the other if desired.  (NOTE!  32-bit screws up selection rendering, because it always assumes
        ' a 24-bit target for performance reasons.  Should revisit!)
        If (srcImage.backBuffer.getDIBWidth <> finalCanvasWidth) Or (srcImage.backBuffer.getDIBHeight <> finalCanvasHeight) Then
            srcImage.backBuffer.createBlank finalCanvasWidth, finalCanvasHeight, 24, g_CanvasBackground, 255
        Else
            GDI_Plus.GDIPlusFillDIBRect srcImage.backBuffer, 0, 0, finalCanvasWidth, finalCanvasHeight, g_CanvasBackground, 255, CompositingModeSourceCopy
        End If
        
        'Cache our viewport position and measurements inside the source object.  Future pipeline stages need these values.
        srcImage.imgViewport.targetLeft = viewportLeft
        srcImage.imgViewport.targetTop = viewportTop
        srcImage.imgViewport.targetWidth = viewportWidth
        srcImage.imgViewport.targetHeight = viewportHeight
            
        'Pass control to the next pipeline stage.
        Stage2_CompositeAllLayers srcImage, dstCanvas
        
    End If
    
    
    
    'This stage of the pipeline has completed successfully!
    Exit Sub



ViewportPipeline_Stage1_Error:

    Select Case Err.Number
    
        'Out of memory
        Case 480
            pdMsgBox "There is not enough memory available to continue this operation.  Please free up system memory (RAM) and try again.  If the problem persists, reduce the zoom value and try again.", vbExclamation + vbOKOnly, "Out of memory"
            SetProgBarVal 0
            releaseProgressBar
            Message "Operation halted."
            
        'Anything else.  (Never encountered; failsafe only.)
        Case Else
            Message "Viewport rendering paused due to unexpected error (#%1)", Err
            
    End Select

End Sub

'When all images have been unloaded, the temporary front buffer can also be erased to keep memory usage as low as possible.
' While not actually part of the viewport pipeline, I find it intuitive to store this function here.
Public Sub eraseViewportBuffers()
    If Not frontBuffer Is Nothing Then
        frontBuffer.eraseDIB
        Set frontBuffer = Nothing
    End If
End Sub
