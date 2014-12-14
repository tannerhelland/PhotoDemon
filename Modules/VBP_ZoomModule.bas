Attribute VB_Name = "Viewport_Engine"
'***************************************************************************
'Viewport Handler - builds and draws the image viewport and associated scroll bars
'Copyright ©2001-2014 by Tanner Helland
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

'Viewport_Engine.Stage3_CompositeCanvas is the last step in the viewport chain.  (Viewport_Engine.Stage1_InitializeBuffer -> Viewport_Engine.Stage2_CompositeAllLayers -> Viewport_Engine.Stage3_CompositeCanvas)
' It can only be executed after both Viewport_Engine.Stage1_InitializeBuffer and Viewport_Engine.Stage2_CompositeAllLayers have been run at least once.  It assumes a fully composited backbuffer,
' which is then copied to the front buffer, and any final composites (such as a selection) are drawn atop that.
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

    'Reset the front buffer
    If Not (frontBuffer Is Nothing) Then
        frontBuffer.eraseDIB
        Set frontBuffer = Nothing
    End If
    Set frontBuffer = New pdDIB
    
    'We can use the .Tag property of the target form to locate the matching pdImage in the pdImages array
    Dim curImage As Long
    curImage = srcImage.imageID
    
    'Copy the current back buffer into the front buffer
    frontBuffer.createFromExistingDIB srcImage.backBuffer
    
    'If the user has allowed interface decorations, handle that next
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
    
    'Check to see if a selection is active.
    If srcImage.selectionActive Then
    
        'If it is, composite the selection against the front buffer
        srcImage.mainSelection.renderCustom frontBuffer, srcImage, FormMain.mainCanvas(0), srcImage.imgViewport.targetLeft, srcImage.imgViewport.targetTop, srcImage.imgViewport.targetWidth, srcImage.imgViewport.targetHeight, toolbar_Options.cboSelRender.ListIndex, toolbar_Options.csSelectionHighlight.Color
    
    End If
    
    'In the future, any additional UI compositing can be handled here.
    
    'Because AutoRedraw can cause the form's DC to change without warning, we must re-apply color management settings any time
    ' we redraw the screen.  I do not like this any more than you do, but we risk losing our DC's settings otherwise.
    If Not (g_UseSystemColorProfile And g_IsSystemColorProfileSRGB) Then
        assignDefaultColorProfileToObject dstCanvas.hWnd, dstCanvas.hDC
        turnOnColorManagementForDC dstCanvas.hDC
    End If
    
    'Finally, flip the front buffer to the screen
    'BitBlt formToBuffer.hDC, 0, 26, frontBuffer.getDIBWidth, frontBuffer.getDIBHeight, frontBuffer.getDIBDC, 0, 0, vbSrcCopy
    BitBlt dstCanvas.hDC, 0, srcImage.imgViewport.getTopOffset, frontBuffer.getDIBWidth, frontBuffer.getDIBHeight, frontBuffer.getDIBDC, 0, 0, vbSrcCopy
        
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
    
    'Finally, we can do some tool-specific rendering directly onto the form.
    Select Case g_CurrentTool
    
        'The nav tool provides two render options at present: draw layer borders, and draw layer transform nodes
        Case NAV_MOVE
        
            'If the user has requested visible layer borders, draw them now
            If CBool(toolbar_Options.chkLayerBorder) Then
                
                'Draw layer borders
                Drawing.drawLayerBoundaries pdImages(g_CurrentImage).getActiveLayerIndex
                
            End If
            
            'If the user has requested visible transformation nodes, draw them now
            If CBool(toolbar_Options.chkLayerNodes) Then
                
                'Draw layer nodes
                Drawing.drawLayerNodes pdImages(g_CurrentImage).getActiveLayerIndex
                
            End If
        
        'Selections are always rendered onto the canvas.  If a selection is active AND a selection tool is active, we can also
        ' draw transform nodes around the selection area.
        Case SELECT_RECT, SELECT_CIRC, SELECT_LINE, SELECT_POLYGON, SELECT_WAND
            
            'Next, check to see if a selection is active and transformable.  If it is, draw nodes around the selected area.
            If srcImage.selectionActive Then
                srcImage.mainSelection.renderTransformNodes srcImage, dstCanvas, srcImage.imgViewport.targetLeft, srcImage.imgViewport.targetTop
            End If
        
    End Select
    
    'With all rendering complete, copy the form's image into the .Picture (e.g. render it on-screen) and refresh
    dstCanvas.requestBufferSync
    
End Sub

'Stage2_CompositeAllLayers is used to update the on-screen image when the scroll bars are used.
' Given how frequently it is used, I've tried to make it as small and fast as possible.
Public Sub Stage2_CompositeAllLayers(ByRef srcImage As pdImage, ByRef dstCanvas As pdCanvas)
    
    
    'Like the previous stage of the pipeline, we start by performing a number of "do not render the viewport at all" checks.
    
    'First, and most obvious, is to exit now if the public g_AllowViewportRendering parameter has been forcibly disabled.
    If Not g_AllowViewportRendering Then Exit Sub
    
    'I think we can successfully ignore this check, as the previous stage handles it, but I'm leaving it here "just in case"
    'If g_OpenImageCount = 0 Then
    '    FormMain.mainCanvas(0).clearCanvas
    '    Exit Sub
    'End If
    
    'Make sure the target canvas is valid
    If dstCanvas Is Nothing Then Exit Sub
    
    'If the pdImage object associated with this form is inactive, ignore this request
    If Not srcImage.IsActive Then Exit Sub
    
    'This function can return timing reports if desired; at present, this is automatically activated in PRE-ALPHA and ALPHA builds,
    ' but disabled for BETA and PRODUCTION builds; see the LoadTheProgram() function for details.
    Dim startTime As Double
    If g_DisplayTimingReports Then startTime = Timer
    
    'Stage 1 of the pipeline (Stage1_InitializeBuffer) prepared
    
    'These variables represent the source width - e.g. the size of the viewable picture box, divided by the zoom coefficient.
    ' Because rounding errors may occur with cerain image sizes, apply a special check when zoom = 100.
    If srcImage.currentZoomValue = g_Zoom.getZoom100Index Then
        srcWidth = srcImage.imgViewport.targetWidth
        srcHeight = srcImage.imgViewport.targetHeight
    Else
        srcWidth = srcImage.imgViewport.targetWidth / m_ZoomRatio
        srcHeight = srcImage.imgViewport.targetHeight / m_ZoomRatio
    End If
        
    'These variables are the offset, as determined by the scroll bar values
    If dstCanvas.getScrollVisibility(PD_HORIZONTAL) Then srcX = dstCanvas.getScrollValue(PD_HORIZONTAL) Else srcX = 0
    If dstCanvas.getScrollVisibility(PD_VERTICAL) Then srcY = dstCanvas.getScrollValue(PD_VERTICAL) Else srcY = 0
        
    'Before rendering the image, apply a checkerboard pattern to the viewport region of the source image's back buffer.
    ' TODO: cache g_CheckerboardPattern persistently, in GDI+ format, so we don't have to recreate it on every draw.
    With srcImage.imgViewport
        GDI_Plus.GDIPlusFillDIBRect_Pattern srcImage.backBuffer, .targetLeft, .targetTop, .targetWidth - 1, .targetHeight - 1, g_CheckerboardPattern
        Debug.Print "Fill GDI+ calc: ", .targetLeft, .targetTop, .targetWidth, .targetHeight
    End With
    
    'As a failsafe, perform a GDI+ check.  PD probably won't work at all without GDI+, so I could look at dropping this check
    ' in the future... but for now, we leave it, just in case.
    If g_GDIPlusAvailable Then
        
        'Use our new rect-specific compositor to retrieve only the relevant section of the current viewport.  Interpolation mode depends
        ' on the current zoom value, and the user's viewport performance preference.
        
        'When we've been asked to maximize performance, use nearest neighbor for all zoom modes
        If g_ViewportPerformance = PD_PERF_FASTEST Then
            srcImage.getCompositedRect srcImage.backBuffer, srcImage.imgViewport.targetLeft, srcImage.imgViewport.targetTop, srcImage.imgViewport.targetWidth, srcImage.imgViewport.targetHeight, srcX, srcY, srcWidth, srcHeight, InterpolationModeNearestNeighbor
            
        'Otherwise, switch dynamically between high-quality and low-quality interpolation depending on the current zoom
        Else
            srcImage.getCompositedRect srcImage.backBuffer, srcImage.imgViewport.targetLeft, srcImage.imgViewport.targetTop, srcImage.imgViewport.targetWidth, srcImage.imgViewport.targetHeight, srcX, srcY, srcWidth, srcHeight, IIf(m_ZoomRatio <= 1, InterpolationModeHighQualityBicubic, InterpolationModeNearestNeighbor)
        End If
        
    'This is an emergency fallback, only.  PD won't work without GDI+, so rendering the viewport is pointless.
    Else
        Message "WARNING!  GDI+ could not be found.  (PhotoDemon requires GDI+ for proper program operation.)"
    End If
    
    'Pass control to the viewport renderer, which will handle the final compositing
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
        
        'Testing shows no measurable difference between a 32-bit or 24-bit back buffer.  I am going to try 32-bit for now,
        ' but you can easily swap 32 for 24 if desired.
        srcImage.backBuffer.createBlank finalCanvasWidth, finalCanvasHeight, 32, g_CanvasBackground, 255
        
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
