Attribute VB_Name = "Viewport_Handler"
'***************************************************************************
'Viewport Handler - builds and draws the image viewport and associated scroll bars
'Copyright ©2001-2014 by Tanner Helland
'Created: 4/15/01
'Last updated: 15/September/13
'Last update: Optimize viewport scrolling if GDI+ is available.
'
'Module for handling the image viewport.  There are key routines:
' - PrepareViewport: for recalculating all viewport variables and controls (done only when the zoom value is changed or a new picture is loaded)
' - ScrollViewport: when the viewport is scrolled (minimal redrawing is done, since the zoom value hasn't changed)
' - RenderViewport: perform any final compositing, such as the Selection Tool effect, then draw the viewport on-screen
'
'PhotoDemon is intelligent about calling the lowest routine in the "render chain", which is how it is able to render the viewport
' so quickly regardless of zoom or scroll values.
'
'Finally, note that StretchBlt is used for the actual rendering, and its "halftone" mode is explicitly specified for shrinking the image.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Width and height values of the image AFTER zoom has been applied.  (For example, if the image is 100x100
' and the zoom value is 200%, zWidth and zHeight will be 200.)
Private zWidth As Double, zHeight As Double

'These variables represent the source width - e.g. the size of the viewable picture box, divided by the zoom coefficient
Private srcWidth As Double, srcHeight As Double

'The ZoomVal value is the actual coefficient for the current zoom value.  (For example, 0.50 for "50% zoom")
Private zoomVal As Double

'These variables are the offset, as determined by the scroll bar values
Private srcX As Long, srcY As Long

'frontBuffer holds the final composited image, including any overlays (like selections)
Private frontBuffer As pdDIB

'cornerFix holds a small gray box that is copied over the corner between the horizontal and vertical scrollbars, if they exist
Private cornerFix As pdDIB

'RenderViewport is the last step in the viewport chain.  (PrepareViewport -> ScrollViewport -> RenderViewport)
' It can only be executed after both PrepareViewport and ScrollViewport have been run at least once.  It assumes a fully composited backbuffer,
' which is then copied to the front buffer, and any final composites (such as a selection) are drawn atop that.
Public Sub RenderViewport(ByRef srcImage As pdImage, ByRef dstCanvas As pdCanvas)

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
    
    'Check to see if a selection is active.
    If srcImage.selectionActive Then
    
        'If it is, composite the selection against the front buffer
        srcImage.mainSelection.renderCustom frontBuffer, srcImage, FormMain.mainCanvas(0), srcImage.imgViewport.targetLeft, srcImage.imgViewport.targetTop, srcImage.imgViewport.targetWidth, srcImage.imgViewport.targetHeight, toolbar_Tools.cmbSelRender(0).ListIndex
    
    End If
        
    'If the user has requested a drop shadow drawn onto the canvas, handle that next
    If g_CanvasDropShadow Then
    
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
    
    'In the future, additional compositing can be handled here.
    
    'Because AutoRedraw can cause the form's DC to change without warning, we must re-apply color management settings any time
    ' we redraw the screen.  I do not like this any more than you do, but we risk losing our DC's settings otherwise.
    assignDefaultColorProfileToObject dstCanvas.hWnd, dstCanvas.hDC
    turnOnColorManagementForDC dstCanvas.hDC
    
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
    
    'Check to see if a selection is active and transformable.  If it is, draw nodes around the selected area.
    If srcImage.selectionActive And srcImage.mainSelection.isTransformable Then
    
        'If it is, composite the selection against the temporary buffer
        srcImage.mainSelection.renderTransformNodes srcImage, dstCanvas, srcImage.imgViewport.targetLeft, srcImage.imgViewport.targetTop
    
    End If
    
    
    'With all rendering complete, copy the form's image into the .Picture (e.g. render it on-screen) and refresh
    dstCanvas.requestBufferSync
    
End Sub

'ScrollViewport is used to update the on-screen image when the scroll bars are used.
' Given how frequently it is used, I've tried to make it as small and fast as possible.
Public Sub ScrollViewport(ByRef srcImage As pdImage, ByRef dstCanvas As pdCanvas)
    
    'If no images have been loaded, clear the canvas and exit
    If g_OpenImageCount = 0 Then
        FormMain.mainCanvas(0).clearCanvas
        Exit Sub
    End If
    
    'Make sure the target form is valid
    If dstCanvas Is Nothing Then Exit Sub
    
    'If the image associated with this form is inactive, ignore this request
    If Not srcImage.IsActive Then Exit Sub
    
    'We can use the .Tag property of the target form to locate the matching pdImage in the pdImages array
    Dim curImage As Long
    curImage = srcImage.imageID
    
    'The ZoomVal value is the actual coefficient for the current zoom value.  (For example, 0.50 for "50% zoom")
    zoomVal = g_Zoom.getZoomValue(srcImage.currentZoomValue)

    'These variables represent the source width - e.g. the size of the viewable picture box, divided by the zoom coefficient
    srcWidth = srcImage.imgViewport.targetWidth / zoomVal
    srcHeight = srcImage.imgViewport.targetHeight / zoomVal
        
    'These variables are the offset, as determined by the scroll bar values
    If dstCanvas.getScrollVisibility(PD_HORIZONTAL) Then srcX = dstCanvas.getScrollValue(PD_HORIZONTAL) Else srcX = 0
    If dstCanvas.getScrollVisibility(PD_VERTICAL) Then srcY = dstCanvas.getScrollValue(PD_VERTICAL) Else srcY = 0
        
    'Paint the image from the back buffer to the front buffer.  We handle this as two cases: one for zooming in, another for zooming out.
    ' This is simpler from a coding standpoint, as each case involves a number of specialized calculations.
    
    'Retrieve a composited copy of the image.  (TODO: optimize this process... somehow.)
    Dim compositedImage As pdDIB
    Set compositedImage = New pdDIB
    srcImage.getCompositedImage compositedImage
    
    If zoomVal < 1 Then
        
        'ZOOMED OUT
        
        'Check for alpha channel.  If it's found, perform pre-multiplication against a checkered background before rendering.
        If compositedImage.getDIBColorDepth = 32 Then
            
            'Before rendering the image, apply a checkerboard pattern to the target image's back buffer
            Drawing.fillDIBWithAlphaCheckerboard srcImage.backBuffer, srcImage.imgViewport.targetLeft, srcImage.imgViewport.targetTop, srcImage.imgViewport.targetWidth, srcImage.imgViewport.targetHeight
            
            'Update 14 Feb '14: If GDI+ is available, use it to render 32bpp images when zoomed out.  (This is preferable to StretchBlt,
            ' as StretchBlt erases all alpha channel data if HALFTONE mode is used, limiting it to nearest-neighbor only!)
            If g_GDIPlusAvailable Then
                GDIPlusResizeDIB srcImage.backBuffer, srcImage.imgViewport.targetLeft, srcImage.imgViewport.targetTop, srcImage.imgViewport.targetWidth, srcImage.imgViewport.targetHeight, compositedImage, srcX, srcY, srcWidth, srcHeight, InterpolationModeHighQualityBicubic
            Else
            
                'Create a blank DIB in the parent pdImages object.  (For performance reasons, we create this image at the size of the viewport.)
                srcImage.alphaFixDIB.createBlank srcWidth, srcHeight, 32
                BitBlt srcImage.alphaFixDIB.getDIBDC, 0, 0, srcWidth, srcHeight, compositedImage.getDIBDC, srcX, srcY, vbSrcCopy
                
                'Paint that chopped-out DIB to the target image's back buffer
                srcImage.alphaFixDIB.alphaBlendToDC srcImage.backBuffer.getDIBDC, 255, srcImage.imgViewport.targetLeft, srcImage.imgViewport.targetTop, srcImage.imgViewport.targetWidth, srcImage.imgViewport.targetHeight
                
            End If
            
        Else
            SetStretchBltMode srcImage.backBuffer.getDIBDC, STRETCHBLT_HALFTONE
            StretchBlt srcImage.backBuffer.getDIBDC, srcImage.imgViewport.targetLeft, srcImage.imgViewport.targetTop, srcImage.imgViewport.targetWidth, srcImage.imgViewport.targetHeight, compositedImage.getDIBDC, srcX, srcY, srcWidth, srcHeight, vbSrcCopy
        End If
        
    Else
    
        'ZOOMED IN (OR 100%)
        
        'When zoomed in, the blitting call must be modified as follows: restrict it to multiples of the current zoom factor.
        ' (Without this fix, funny stretching occurs; to see it yourself, place the zoom at 300%, and drag an image's window larger or smaller.)
        ' NOTE: I have removed that stretching fix, because it causes invalid rendering later down the chain.  As it's not
        '       a particularly pressing concern, I will revisit at some point in the future (ETA to be determined).
        Dim bltWidth As Long, bltHeight As Long
        bltWidth = srcImage.imgViewport.targetWidth '+ (Int(g_Zoom.getZoomOffsetFactor(srcImage.currentZoomValue)) - (srcImage.imgViewport.targetWidth Mod Int(g_Zoom.getZoomOffsetFactor(srcImage.currentZoomValue))))
        srcWidth = bltWidth / zoomVal
        bltHeight = srcImage.imgViewport.targetHeight '+ (Int(g_Zoom.getZoomOffsetFactor(srcImage.currentZoomValue)) - (srcImage.imgViewport.targetHeight Mod Int(g_Zoom.getZoomOffsetFactor(srcImage.currentZoomValue))))
        srcHeight = bltHeight / zoomVal
        
        'Check for alpha channel.  If it's found, perform pre-multiplication against a checkered background before rendering.
        If compositedImage.getDIBColorDepth = 32 Then
            
            'Fill the target area with the alpha checkerboard
            Drawing.fillDIBWithAlphaCheckerboard srcImage.backBuffer, srcImage.imgViewport.targetLeft, srcImage.imgViewport.targetTop, srcImage.imgViewport.targetWidth, srcImage.imgViewport.targetHeight
            
            'Create a temporary streched copy of the image
            srcImage.alphaFixDIB.createBlank bltWidth, bltHeight, 32
            SetStretchBltMode srcImage.alphaFixDIB.getDIBDC, STRETCHBLT_COLORONCOLOR
            StretchBlt srcImage.alphaFixDIB.getDIBDC, 0, 0, bltWidth, bltHeight, compositedImage.getDIBDC, srcX, srcY, srcWidth, srcHeight, vbSrcCopy
            
            'Alpha blend the DIB onto the checkerboard background
            srcImage.alphaFixDIB.alphaBlendToDC srcImage.backBuffer.getDIBDC, 255, srcImage.imgViewport.targetLeft, srcImage.imgViewport.targetTop, srcImage.imgViewport.targetWidth, srcImage.imgViewport.targetHeight
            
        Else
            SetStretchBltMode srcImage.backBuffer.getDIBDC, STRETCHBLT_COLORONCOLOR
            StretchBlt srcImage.backBuffer.getDIBDC, srcImage.imgViewport.targetLeft, srcImage.imgViewport.targetTop, bltWidth, bltHeight, compositedImage.getDIBDC, srcX, srcY, srcWidth, srcHeight, vbSrcCopy
        End If
        
    End If
    
    'Pass control to the viewport renderer, which will handle the final compositing
    RenderViewport srcImage, dstCanvas

End Sub

'Per its name, PrepareViewport is responsible for calculating the maximum values and positions of the viewport scroll bars
' based on an image form's size and position.  It needs to be executed when:
    '1) an image is first loaded
    '2) an image's zoom value is changed
    '3) an image's container form is resized
    '4) other special cases (resizing an image, rotating an image - basically anything that changes the size of the back buffer)

'Note that specific zoom values are calculated in other routines; they are only USED here.

'This routine requires a target form as a parameter.  This form will almost always be pdImages(g_CurrentImage).containingForm, but in
' certain rare cases (cascading windows, for example), it may be necessary to recalculate the viewport and scroll bars
' in non-active windows - in those cases, the calling routine must specify which viewport it wants rebuilt.

'Because redrawing a viewport from scratch is an expensive operation, this function also takes a "reasonForRedraw" parameter, which
' is an untranslated string supplied by the caller.  I use this to track when viewport redraws are requested, and to try and keep
' such requests as infrequent as possible.
Public Sub PrepareViewport(ByRef srcImage As pdImage, ByRef dstCanvas As pdCanvas, Optional ByRef reasonForRedraw As String)

    'Don't attempt to resize the scroll bars if g_AllowViewportRendering is disabled. This is used to provide a smoother user experience,
    ' especially when images are being loaded. (This routine is triggered on Form_Resize, which is in turn triggered when a
    ' new picture is loaded.  To prevent PrepareViewport from being fired multiple times, g_AllowViewportRendering is utilized.)
    If Not g_AllowViewportRendering Then Exit Sub
    
    'Make sure the target canvas is valid
    If dstCanvas Is Nothing Then Exit Sub
    
    'If no images have been loaded, clear the canvas and exit
    If g_OpenImageCount = 0 Then
        FormMain.mainCanvas(0).clearCanvas
        Exit Sub
    End If
    
    'We can use the .Tag property of the target form to locate the matching pdImage in the pdImages array
    Dim curImage As Long
    curImage = srcImage.imageID
    
    'If the image associated with this form is inactive, ignore this request
    If Not srcImage.IsActive Then Exit Sub
    
    'Because this routine is time-consuming, I track it carefully to try and minimize how frequently it's called.  Feel free to comment out this line.
    Debug.Print "Preparing viewport: " & reasonForRedraw & " | (" & curImage & ") | " '& formToBuffer.Caption
    
    On Error GoTo ZoomErrorHandler
    
    'Get the mathematical zoom multiplier (based on the current combo box setting - for example, 0.50 for "50% zoom")
    Dim zoomVal As Double
    zoomVal = g_Zoom.getZoomValue(srcImage.currentZoomValue)
    
    'Calculate the width and height of a full-size viewport based on the current zoom value
    zWidth = (srcImage.Width * zoomVal)
    zHeight = (srcImage.Height * zoomVal)
    
    'Calculate the vertical offset of the viewport.  This changes according to the height of the top-aligned status bar,
    ' and in the future, it will also change if rulers are visible.
    Dim verticalOffset As Long
    verticalOffset = srcImage.imgViewport.getVerticalOffset
    
    'Grab the form dimensions; these are necessary for rendering the scroll bars
    Dim canvasWidth As Long, canvasHeight As Long
    canvasWidth = dstCanvas.getCanvasWidth
    canvasHeight = dstCanvas.getCanvasHeight - verticalOffset
    
    'These variables will reflect whether or not scroll bars are enabled; this is used rather than the .Enabled property so we
    ' can defer rendering the scroll bars until the last possible instant (rather than turning them on-and-off mid-subroutine)
    Dim hScrollEnabled As Boolean, vScrollEnabled As Boolean
    hScrollEnabled = False
    vScrollEnabled = False
    
    'Step 1: compare viewport width to zoomed image width
    If Int(zWidth) > canvasWidth Then hScrollEnabled = True
    
    'Step 2: compare viewport height to zoomed image height.  If the horizontal scrollbar has been enabled, factor that into our calculations
    If (Int(zHeight) > canvasHeight) Or (hScrollEnabled And (Int(zHeight) > (canvasHeight - dstCanvas.getScrollHeight(PD_HORIZONTAL)))) Then vScrollEnabled = True
    
    'Step 3: one last check on horizontal viewport width; if the vertical scrollbar was enabled, the horizontal viewport width has changed.
    If vScrollEnabled And (Not hScrollEnabled) And (Int(zWidth) > (canvasWidth - dstCanvas.getScrollWidth(PD_VERTICAL))) Then hScrollEnabled = True
    
    'We now know which scroll bars need to be enabled.  Before calculating scroll bar stuff, however, let's figure out where our viewport will
    ' be located - on the edge if scroll bars are enabled, or centered in the viewable area if scroll bars are NOT enabled.
    
    'Additionally, calculate viewport size - full form size if scroll bars are enabled, full image size (including zoom) if they are not
    Dim viewportLeft As Long, viewportTop As Long
    Dim viewportWidth As Long, viewportHeight As Long
    
    If hScrollEnabled Then
        viewportLeft = 0
        If Not vScrollEnabled Then
            viewportWidth = canvasWidth
        Else
            viewportWidth = canvasWidth - dstCanvas.getScrollWidth(PD_VERTICAL)
        End If
    Else
        viewportWidth = zWidth
        If Not vScrollEnabled Then
            viewportLeft = (canvasWidth - zWidth) / 2
        Else
            viewportLeft = ((canvasWidth - dstCanvas.getScrollWidth(PD_VERTICAL)) - zWidth) / 2
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
        viewportHeight = zHeight
        If Not hScrollEnabled Then
            viewportTop = (canvasHeight - zHeight) / 2
        Else
            viewportTop = ((canvasHeight - dstCanvas.getScrollHeight(PD_HORIZONTAL)) - zHeight) / 2
        End If
    End If
    
    'Now we know 1) which scrollbars are enabled, 2) the position of our viewport, 3) the size of our viewport.  Knowing this, we can now calculate
    ' the scroll bar values.
    
    'First - if no scroll bars are enabled, draw the viewport and exit.
    If (Not hScrollEnabled) And (Not vScrollEnabled) Then
    
        'Reset the scroll bar values so ScrollViewport doesn't assume we want scrolling
        dstCanvas.setRedrawSuspension True
        dstCanvas.setScrollValue PD_BOTH, 0
        dstCanvas.setRedrawSuspension False
    
        'Hide the scroll bars if necessary
        If dstCanvas.getScrollVisibility(PD_HORIZONTAL) Then dstCanvas.setScrollVisibility PD_HORIZONTAL, False
        If dstCanvas.getScrollVisibility(PD_VERTICAL) Then dstCanvas.setScrollVisibility PD_VERTICAL, False
            
        'Resize the buffer and store the relevant painting information into this pdImages object
        srcImage.backBuffer.createBlank canvasWidth, canvasHeight, 24, g_CanvasBackground
        srcImage.imgViewport.targetLeft = viewportLeft
        srcImage.imgViewport.targetTop = viewportTop
        srcImage.imgViewport.targetWidth = viewportWidth
        srcImage.imgViewport.targetHeight = viewportHeight
        
        'Pass control to the viewport renderer
        ScrollViewport srcImage, dstCanvas
        
        Exit Sub
        
    End If
    
    'If we've reached this point, one or both scroll bars are enabled.  The time has come to calculate their values.
    'Horizontal scroll bar comes first.
    Dim newScrollMax As Long
    
    If hScrollEnabled Then
    
        'If zoomed-in, set the scroll bar range to the number of not visible pixels.
        If zoomVal <= 1 Then
            newScrollMax = srcImage.Width - Int(viewportWidth * g_Zoom.getZoomOffsetFactor(srcImage.currentZoomValue) + 0.5)
        'If zoomed-out, use a modified formula (as there is no reason to scroll at sub-pixel levels.)
        Else
            newScrollMax = srcImage.Width - Int(viewportWidth / g_Zoom.getZoomOffsetFactor(srcImage.currentZoomValue) + 0.5)
        End If
        
        'If the current scroll value exceeds the new scroll maximum, reduce the scroll bar value to compensate
        If dstCanvas.getScrollValue(PD_HORIZONTAL) > newScrollMax Then
            dstCanvas.setRedrawSuspension True
            dstCanvas.setScrollValue PD_HORIZONTAL, newScrollMax
            dstCanvas.setRedrawSuspension False
        End If
        
        dstCanvas.setScrollMax PD_HORIZONTAL, newScrollMax
        
        'As a convenience to the user, make the scroll bar's LargeChange parameter proportional to the scroll bar's new maximum value
        If dstCanvas.getScrollMax(PD_HORIZONTAL) > 15 Then
            dstCanvas.setScrollLargeChange PD_HORIZONTAL, dstCanvas.getScrollMax(PD_HORIZONTAL) \ 16
        Else
            dstCanvas.setScrollLargeChange PD_HORIZONTAL, 1
        End If
        
    End If
    
    'Same formula, but with width and height swapped for vertical scrolling
    If vScrollEnabled Then
    
        'If zoomed-in, set the scroll bar range to the number of not visible pixels.
        If zoomVal <= 1 Then
            newScrollMax = srcImage.Height - Int(viewportHeight * g_Zoom.getZoomOffsetFactor(srcImage.currentZoomValue) + 0.5)
        'If zoomed-out, use a modified formula (as there is no reason to scroll at sub-pixel levels.)
        Else
            newScrollMax = srcImage.Height - Int(viewportHeight / g_Zoom.getZoomOffsetFactor(srcImage.currentZoomValue) + 0.5)
        End If
        
        'If the current scroll value exceeds the new scroll maximum, reduce the scroll bar value to compensate
        If dstCanvas.getScrollValue(PD_VERTICAL) > newScrollMax Then
            dstCanvas.setRedrawSuspension True
            dstCanvas.setScrollValue PD_VERTICAL, newScrollMax
            dstCanvas.setRedrawSuspension False
        End If
        
        dstCanvas.setScrollMax PD_VERTICAL, newScrollMax
        
        'As a convenience to the user, make the scroll bar's LargeChange parameter proportional to the scroll bar's new maximum value
        If dstCanvas.getScrollMax(PD_VERTICAL) > 15 Then
            dstCanvas.setScrollLargeChange PD_VERTICAL, dstCanvas.getScrollMax(PD_VERTICAL) \ 16
        Else
            dstCanvas.setScrollLargeChange PD_VERTICAL, 1
        End If
        
    End If
    
    'Added to our list of "things we know" is the scroll bar maximum values (and they have already been set).
    ' As such, the time has come to render everything to the screen.
    
    'Horizontal scroll bar gets rendered first...
    If hScrollEnabled Then
        dstCanvas.moveScrollBar PD_HORIZONTAL, 0, canvasHeight - dstCanvas.getScrollHeight(PD_HORIZONTAL), viewportWidth, dstCanvas.getScrollHeight(PD_HORIZONTAL)
        If (Not dstCanvas.getScrollVisibility(PD_HORIZONTAL)) Then dstCanvas.setScrollVisibility PD_HORIZONTAL, True
    Else
        
        'Note that we disable automatic canvas redraws prior to changing the scroll bar value; otherwise, the change will
        ' force a redraw, and we don't want that yet.
        dstCanvas.setRedrawSuspension True
        dstCanvas.setScrollValue PD_HORIZONTAL, 0
        If dstCanvas.getScrollVisibility(PD_HORIZONTAL) Then dstCanvas.setScrollVisibility PD_HORIZONTAL, False
        dstCanvas.setRedrawSuspension False
        
    End If
    
    'Then vertical scroll bar...
    If vScrollEnabled Then
        dstCanvas.moveScrollBar PD_VERTICAL, canvasWidth - dstCanvas.getScrollWidth(PD_VERTICAL), srcImage.imgViewport.getTopOffset, dstCanvas.getScrollWidth(PD_VERTICAL), viewportHeight
        If (Not dstCanvas.getScrollVisibility(PD_VERTICAL)) Then dstCanvas.setScrollVisibility PD_VERTICAL, True
    Else
    
        dstCanvas.setRedrawSuspension True
        dstCanvas.setScrollValue PD_VERTICAL, 0
        If dstCanvas.getScrollVisibility(PD_VERTICAL) Then dstCanvas.setScrollVisibility PD_VERTICAL, False
        dstCanvas.setRedrawSuspension False
        
    End If
    
    'We don't actually render the image here; instead, we prepare the buffer (backBuffer) and store the relevant
    ' drawing variables to this pdImages object.  ScrollViewport (above) will handle the actual drawing.
    Dim newVWidth As Long, newVHeight As Long
    If hScrollEnabled Then newVWidth = viewportWidth Else newVWidth = canvasWidth
    If vScrollEnabled Then newVHeight = viewportHeight Else newVHeight = canvasHeight
    
    'Prepare the relevant back buffer
    If (Not srcImage.backBuffer Is Nothing) Then srcImage.backBuffer.eraseDIB
    srcImage.backBuffer.createBlank newVWidth, newVHeight, 24, g_CanvasBackground
    
    srcImage.imgViewport.targetLeft = viewportLeft
    srcImage.imgViewport.targetTop = viewportTop
    srcImage.imgViewport.targetWidth = viewportWidth
    srcImage.imgViewport.targetHeight = viewportHeight
        
    'Pass control to the viewport renderer (found at the top of this module)
    ScrollViewport srcImage, dstCanvas

    Exit Sub

ZoomErrorHandler:

    If Err = 480 Then
        pdMsgBox "There is not enough memory available to continue this operation.  Please free up system memory (RAM) and try again.  If the problem persists, reduce the zoom value and try again.", vbExclamation + vbOKOnly, "Out of memory"
        SetProgBarVal 0
        releaseProgressBar
        Message "Operation halted."
    ElseIf Err = 13 Then
        Message "Invalid zoom value."
        Exit Sub
    Else
        Message "Viewport rendering paused due to unexpected error (#%1)", Err
        Exit Sub
    End If

End Sub

'When all images have been unloaded, the temporary front buffer can also be erased to keep memory usage as low as possible.
Public Sub eraseViewportBuffers()
    If Not frontBuffer Is Nothing Then
        frontBuffer.eraseDIB
        Set frontBuffer = Nothing
    End If
End Sub
