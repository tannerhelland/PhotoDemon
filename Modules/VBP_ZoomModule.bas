Attribute VB_Name = "Viewport_Handler"
'***************************************************************************
'Viewport Handler - builds and draws the image viewport and associated scroll bars
'Copyright ©2000-2013 by Tanner Helland
'Created: 4/15/01
'Last updated: 12/November/12
'Last update: Maintain scroll bar value whenever possible (e.g. when Undo/Redo data is loaded, do not reset the scroll bars unless we absolutely have to)
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
'***************************************************************************

Option Explicit

'This is the ListIndex of the FormMain zoom combo box that corresponds to 100%
Public Const ZoomIndex100 As Long = 11

'Width and height values of the image AFTER zoom has been applied.  (For example, if the image is 100x100
' and the zoom value is 200%, zWidth and zHeight will be 200.)
Dim zWidth As Double, zHeight As Double

'These variables represent the source width - e.g. the size of the viewable picture box, divided by the zoom coefficient
Dim srcWidth As Double, srcHeight As Double

'The ZoomVal value is the actual coefficient for the current zoom value.  (For example, 0.50 for "50% zoom")
Dim ZoomVal As Double

'These variables are the offset, as determined by the scroll bar values
Dim srcX As Long, srcY As Long

'frontBuffer holds the final composited image, including any overlays (like selections)
Dim frontBuffer As pdLayer

'cornerFix holds a small gray box that is copied over the corner between the horizontal and vertical scrollbars, if they exist
Dim cornerFix As pdLayer

'renderViewport is the last step in the viewport chain.  (PrepareViewport -> ScrollViewport -> renderViewport)
' It can only be executed after both PrepareViewport and ScrollViewport have been run at least once.  It assumes a fully composited backbuffer,
' which is then copied to the front buffer, and any final composites (such as a selection) are drawn atop that.
Public Sub RenderViewport(ByRef formToBuffer As Form)

    'Make sure the form is valid
    If formToBuffer Is Nothing Then Exit Sub
    
    'If the image associated with this form is inactive, ignore this request
    If pdImages(formToBuffer.Tag).IsActive = False Then Exit Sub

    'Reset the front buffer
    If Not (frontBuffer Is Nothing) Then
        frontBuffer.eraseLayer
        Set frontBuffer = Nothing
    End If
    Set frontBuffer = New pdLayer
    
    'Copy the current back buffer into the front buffer
    frontBuffer.createFromExistingLayer pdImages(formToBuffer.Tag).backBuffer

    'Check to see if a selection is active.
    If pdImages(formToBuffer.Tag).selectionActive Then
    
        'If it is, composite the selection against the temporary buffer
        pdImages(formToBuffer.Tag).mainSelection.renderCustom frontBuffer, formToBuffer, pdImages(formToBuffer.Tag).targetLeft, pdImages(formToBuffer.Tag).targetTop, pdImages(formToBuffer.Tag).targetWidth, pdImages(formToBuffer.Tag).targetHeight, g_selectionRenderPreference
    
    End If
        
    'If the user has requested a drop shadow drawn onto the canvas, handle that next
    If g_CanvasDropShadow Then
    
        'We'll handle this in two steps; first, the horizontal stretches
        If formToBuffer.VScroll.Visible = False Then
                    
            'Make sure the image isn't snugly fit inside the viewport; if it is, this is a waste of time
            If pdImages(formToBuffer.Tag).targetTop <> 0 Then
                'Top edge
                StretchBlt frontBuffer.getLayerDC, pdImages(formToBuffer.Tag).targetLeft, pdImages(formToBuffer.Tag).targetTop - PD_CANVASSHADOWSIZE, pdImages(formToBuffer.Tag).targetWidth, PD_CANVASSHADOWSIZE, g_CanvasShadow.getShadowDC(0), 0, 0, 1, PD_CANVASSHADOWSIZE, vbSrcCopy
                'Bottom edge
                StretchBlt frontBuffer.getLayerDC, pdImages(formToBuffer.Tag).targetLeft, pdImages(formToBuffer.Tag).targetTop + pdImages(formToBuffer.Tag).targetHeight, pdImages(formToBuffer.Tag).targetWidth, PD_CANVASSHADOWSIZE, g_CanvasShadow.getShadowDC(1), 0, 0, 1, PD_CANVASSHADOWSIZE, vbSrcCopy
            End If
        
        End If
        
        'Second, the vertical stretches
        If Not formToBuffer.HScroll.Visible Then
                    
            'Make sure the image isn't snugly fit inside the viewport; if it is, this is a waste of time
            If pdImages(formToBuffer.Tag).targetLeft <> 0 Then
                'Left edge
                StretchBlt frontBuffer.getLayerDC, pdImages(formToBuffer.Tag).targetLeft - PD_CANVASSHADOWSIZE, pdImages(formToBuffer.Tag).targetTop, PD_CANVASSHADOWSIZE, pdImages(formToBuffer.Tag).targetHeight, g_CanvasShadow.getShadowDC(2), 0, 0, PD_CANVASSHADOWSIZE, 1, vbSrcCopy
                'Right edge
                StretchBlt frontBuffer.getLayerDC, pdImages(formToBuffer.Tag).targetLeft + pdImages(formToBuffer.Tag).targetWidth, pdImages(formToBuffer.Tag).targetTop, PD_CANVASSHADOWSIZE, pdImages(formToBuffer.Tag).targetHeight, g_CanvasShadow.getShadowDC(3), 0, 0, PD_CANVASSHADOWSIZE, 1, vbSrcCopy
            End If
        
        End If
        
        'Finally, the corners, which are only drawn if both scroll bars are invisible
        If (Not formToBuffer.VScroll.Visible) And (Not formToBuffer.HScroll.Visible) Then
        
            'NW corner
            StretchBlt frontBuffer.getLayerDC, pdImages(formToBuffer.Tag).targetLeft - PD_CANVASSHADOWSIZE, pdImages(formToBuffer.Tag).targetTop - PD_CANVASSHADOWSIZE, PD_CANVASSHADOWSIZE, PD_CANVASSHADOWSIZE, g_CanvasShadow.getShadowDC(4), 0, 0, PD_CANVASSHADOWSIZE, PD_CANVASSHADOWSIZE, vbSrcCopy
            'NE corner
            StretchBlt frontBuffer.getLayerDC, pdImages(formToBuffer.Tag).targetLeft + pdImages(formToBuffer.Tag).targetWidth, pdImages(formToBuffer.Tag).targetTop - PD_CANVASSHADOWSIZE, PD_CANVASSHADOWSIZE, PD_CANVASSHADOWSIZE, g_CanvasShadow.getShadowDC(5), 0, 0, PD_CANVASSHADOWSIZE, PD_CANVASSHADOWSIZE, vbSrcCopy
            'SW corner
            StretchBlt frontBuffer.getLayerDC, pdImages(formToBuffer.Tag).targetLeft - PD_CANVASSHADOWSIZE, pdImages(formToBuffer.Tag).targetTop + pdImages(formToBuffer.Tag).targetHeight, PD_CANVASSHADOWSIZE, PD_CANVASSHADOWSIZE, g_CanvasShadow.getShadowDC(6), 0, 0, PD_CANVASSHADOWSIZE, PD_CANVASSHADOWSIZE, vbSrcCopy
            'SE corner
            StretchBlt frontBuffer.getLayerDC, pdImages(formToBuffer.Tag).targetLeft + pdImages(formToBuffer.Tag).targetWidth, pdImages(formToBuffer.Tag).targetTop + pdImages(formToBuffer.Tag).targetHeight, PD_CANVASSHADOWSIZE, PD_CANVASSHADOWSIZE, g_CanvasShadow.getShadowDC(7), 0, 0, PD_CANVASSHADOWSIZE, PD_CANVASSHADOWSIZE, vbSrcCopy
        
        End If
    
    End If
    
    'In the future, additional compositing can be handled here.
    
    'Finally, flip the front buffer to the screen
    BitBlt formToBuffer.hDC, 0, 0, frontBuffer.getLayerWidth, frontBuffer.getLayerHeight, frontBuffer.getLayerDC, 0, 0, vbSrcCopy
    
    'If both scrollbars are active, copy a gray square over the small space between them
    If formToBuffer.HScroll.Visible And formToBuffer.VScroll.Visible Then
        
        'Only initialize the corner fix image once
        If cornerFix Is Nothing Then
            Set cornerFix = New pdLayer
            cornerFix.createBlank formToBuffer.VScroll.Width, formToBuffer.HScroll.Height, 24, vbButtonFace
        End If
        
        'Draw the square over any exposed parts of the image in the bottom-right of the image, between the scroll bars
        BitBlt formToBuffer.hDC, formToBuffer.VScroll.Left, formToBuffer.HScroll.Top, cornerFix.getLayerWidth, cornerFix.getLayerHeight, cornerFix.getLayerDC, 0, 0, vbSrcCopy
        
    End If
    
    formToBuffer.Picture = formToBuffer.Image
    formToBuffer.Refresh
    
    'If we don't fire DoEvents here, the image will only scroll after the mouse button is released.
    'DoEvents

End Sub

'ScrollViewport is used to update the on-screen image when the scroll bars are used.
' Given how frequently it is used, I've tried to make it as small and fast as possible.
Public Sub ScrollViewport(ByRef formToBuffer As Form)
    
    'Make sure the form is valid
    If formToBuffer Is Nothing Then Exit Sub
    
    'If the image associated with this form is inactive, ignore this request
    If pdImages(formToBuffer.Tag).IsActive = False Then Exit Sub
    
    'The ZoomVal value is the actual coefficient for the current zoom value.  (For example, 0.50 for "50% zoom")
    ZoomVal = g_Zoom.ZoomArray(pdImages(formToBuffer.Tag).CurrentZoomValue)

    'These variables represent the source width - e.g. the size of the viewable picture box, divided by the zoom coefficient
    srcWidth = pdImages(formToBuffer.Tag).targetWidth / ZoomVal
    srcHeight = pdImages(formToBuffer.Tag).targetHeight / ZoomVal
        
    'These variables are the offset, as determined by the scroll bar values
    If formToBuffer.HScroll.Visible Then srcX = formToBuffer.HScroll.Value Else srcX = 0
    If formToBuffer.VScroll.Visible Then srcY = formToBuffer.VScroll.Value Else srcY = 0
        
    'Paint the image from the back buffer to the front buffer
    If ZoomVal < 1 Then
        
        'Check for alpha channel.  If it's found, perform pre-multiplication against a checkered background before rendering.
        If pdImages(formToBuffer.Tag).mainLayer.getLayerColorDepth = 32 Then
            
            'Create a copy of the current layer in the parent pdImages object
            pdImages(formToBuffer.Tag).alphaFixLayer.createBlank pdImages(formToBuffer.Tag).targetWidth, pdImages(formToBuffer.Tag).targetHeight, 32

            'Now comes a nasty hack; HALFTONE stretching does not preserve the alpha channel, but COLORONCOLOR does.  So make two copies -
            ' one with color-on-color, from which we'll steal alpha values, and a high-quality halftone one for pixel values.
            Dim hackLayer As pdLayer
            Set hackLayer = New pdLayer
            hackLayer.createBlank pdImages(formToBuffer.Tag).targetWidth, pdImages(formToBuffer.Tag).targetHeight, 32
            
            SetStretchBltMode hackLayer.getLayerDC, STRETCHBLT_COLORONCOLOR
            StretchBlt hackLayer.getLayerDC, 0, 0, pdImages(formToBuffer.Tag).targetWidth, pdImages(formToBuffer.Tag).targetHeight, pdImages(formToBuffer.Tag).mainLayer.getLayerDC, srcX, srcY, srcWidth, srcHeight, vbSrcCopy
            
            SetStretchBltMode pdImages(formToBuffer.Tag).alphaFixLayer.getLayerDC, STRETCHBLT_HALFTONE
            StretchBlt pdImages(formToBuffer.Tag).alphaFixLayer.getLayerDC, 0, 0, pdImages(formToBuffer.Tag).targetWidth, pdImages(formToBuffer.Tag).targetHeight, pdImages(formToBuffer.Tag).mainLayer.getLayerDC, srcX, srcY, srcWidth, srcHeight, vbSrcCopy
            pdImages(formToBuffer.Tag).alphaFixLayer.compositeBackgroundColorSpecial hackLayer
            BitBlt pdImages(formToBuffer.Tag).backBuffer.getLayerDC, pdImages(formToBuffer.Tag).targetLeft, pdImages(formToBuffer.Tag).targetTop, pdImages(formToBuffer.Tag).targetWidth, pdImages(formToBuffer.Tag).targetHeight, pdImages(formToBuffer.Tag).alphaFixLayer.getLayerDC, 0, 0, vbSrcCopy
            
            'Remove our temporary layer from memory to prevent leaks
            hackLayer.eraseLayer
            Set hackLayer = Nothing
            
        Else
            SetStretchBltMode pdImages(formToBuffer.Tag).backBuffer.getLayerDC, STRETCHBLT_HALFTONE
            StretchBlt pdImages(formToBuffer.Tag).backBuffer.getLayerDC, pdImages(formToBuffer.Tag).targetLeft, pdImages(formToBuffer.Tag).targetTop, pdImages(formToBuffer.Tag).targetWidth, pdImages(formToBuffer.Tag).targetHeight, pdImages(formToBuffer.Tag).mainLayer.getLayerDC(), srcX, srcY, srcWidth, srcHeight, vbSrcCopy
        End If
        
    Else
        'When zoomed in, the blitting call must be modified as follows: restrict it to multiples of the current zoom factor.
        ' (Without this fix, funny stretching occurs; to see it yourself, place the zoom at 300%, and drag an image's window larger or smaller.)
        Dim bltWidth As Long, bltHeight As Long
        bltWidth = pdImages(formToBuffer.Tag).targetWidth + (Int(g_Zoom.ZoomFactor(pdImages(formToBuffer.Tag).CurrentZoomValue)) - (pdImages(formToBuffer.Tag).targetWidth Mod Int(g_Zoom.ZoomFactor(pdImages(formToBuffer.Tag).CurrentZoomValue))))
        srcWidth = bltWidth / ZoomVal
        bltHeight = pdImages(formToBuffer.Tag).targetHeight + (Int(g_Zoom.ZoomFactor(pdImages(formToBuffer.Tag).CurrentZoomValue)) - (pdImages(formToBuffer.Tag).targetHeight Mod Int(g_Zoom.ZoomFactor(pdImages(formToBuffer.Tag).CurrentZoomValue))))
        srcHeight = bltHeight / ZoomVal
        
        'Check for alpha channel.  If it's found, perform pre-multiplication against a checkered background before rendering.
        If pdImages(formToBuffer.Tag).mainLayer.getLayerColorDepth = 32 Then
            pdImages(formToBuffer.Tag).alphaFixLayer.createBlank bltWidth, bltHeight, 32
            SetStretchBltMode pdImages(formToBuffer.Tag).alphaFixLayer.getLayerDC, STRETCHBLT_COLORONCOLOR
            StretchBlt pdImages(formToBuffer.Tag).alphaFixLayer.getLayerDC, 0, 0, bltWidth, bltHeight, pdImages(formToBuffer.Tag).mainLayer.getLayerDC(), srcX, srcY, srcWidth, srcHeight, vbSrcCopy
            pdImages(formToBuffer.Tag).alphaFixLayer.compositeBackgroundColor
            BitBlt pdImages(formToBuffer.Tag).backBuffer.getLayerDC, pdImages(formToBuffer.Tag).targetLeft, pdImages(formToBuffer.Tag).targetTop, pdImages(formToBuffer.Tag).targetWidth, pdImages(formToBuffer.Tag).targetHeight, pdImages(formToBuffer.Tag).alphaFixLayer.getLayerDC, 0, 0, vbSrcCopy
        Else
            SetStretchBltMode pdImages(formToBuffer.Tag).backBuffer.getLayerDC, STRETCHBLT_COLORONCOLOR
            StretchBlt pdImages(formToBuffer.Tag).backBuffer.getLayerDC, pdImages(formToBuffer.Tag).targetLeft, pdImages(formToBuffer.Tag).targetTop, bltWidth, bltHeight, pdImages(formToBuffer.Tag).mainLayer.getLayerDC, srcX, srcY, srcWidth, srcHeight, vbSrcCopy
        End If
        
    End If
    
    'Pass control to the viewport renderer, which will handle the final compositing
    RenderViewport formToBuffer

End Sub

'PrepareViewport is responsible for calculating the position and size of the main viewport picture box, as well as the maximum values
' and positions of the viewport scroll bars.  It needs to be executed when:
    '1) an image is first loaded
    '2) an image's zoom value is changed
    '3) other special cases (resizing an image, rotating an image - basically anything that changes the size of the back buffer)

'Note that specific zoom values are calculated in other routines; they are only USED here.

'This routine requires a target form as a parameter.  This form will almost always be FormMain.ActiveForm, but in
' certain rare cases (cascading windows, for example), it may be necessary to recalculate the viewport and scroll bars
' in non-active windows - in those cases, the calling routine must tell us which viewport it wants rebuilt.
Public Sub PrepareViewport(ByRef formToBuffer As Form, Optional ByRef reasonForRedraw As String)

    'Don't attempt to resize the scroll bars if g_FixScrolling is disabled. This is used to provide a smoother user experience,
    ' especially when images are being loaded. (This routine is triggered on Form_Resize, which is in turn triggered when a
    ' new picture is loaded.  To prevent PrepareViewport from being fired multiple times, g_FixScrolling is utilized.)
    If g_FixScrolling = False Then Exit Sub
    
    'Make sure the form is valid
    If formToBuffer Is Nothing Then Exit Sub
    
    'If the image associated with this form is inactive, ignore this request
    If pdImages(formToBuffer.Tag).IsActive = False Then Exit Sub
    
    'Because this routine is time-consuming, I track it carefully to try and minimize how frequently it's called.  Feel free to comment out this line.
    Debug.Print "Preparing viewport: " & reasonForRedraw & " | (" & formToBuffer.Tag & ") | " & formToBuffer.Caption
    
    On Error GoTo ZoomErrorHandler
    
    'Get the mathematical zoom multiplier (based on the current combo box setting - for example, 0.50 for "50% zoom")
    Dim ZoomVal As Double
    ZoomVal = g_Zoom.ZoomArray(pdImages(formToBuffer.Tag).CurrentZoomValue)
    
    'Calculate the width and height of the full-size viewport based on the current zoom value
    zWidth = (pdImages(formToBuffer.Tag).Width * ZoomVal)
    zHeight = (pdImages(formToBuffer.Tag).Height * ZoomVal)
    
    'Grab the form dimensions; these are necessary for rendering the scroll bars
    Dim FormWidth As Long, FormHeight As Long
    FormWidth = formToBuffer.ScaleWidth
    FormHeight = formToBuffer.ScaleHeight
    
    'These variables will reflect whether or not scroll bars are enabled; this is used rather than the .Enabled property so we
    ' can defer rendering the scroll bars until the last possible instant (rather than turning them on-and-off mid-subroutine)
    Dim hScrollEnabled As Boolean, vScrollEnabled As Boolean
    hScrollEnabled = False
    vScrollEnabled = False
    
    'Step 1: compare viewport width to zoomed image width
    If Int(zWidth) > FormWidth Then hScrollEnabled = True
    
    'Step 2: compare viewport height to zoomed image height.  If the horizontal scrollbar has been enabled, factor that into our calculations
    If (Int(zHeight) > FormHeight) Or ((hScrollEnabled = True) And (Int(zHeight) > (FormHeight - formToBuffer.HScroll.Height))) Then vScrollEnabled = True
    
    'Step 3: one last check on horizontal viewport width; if the vertical scrollbar was enabled, the horizontal viewport width has changed.
    If (vScrollEnabled = True) And (hScrollEnabled = False) And (Int(zWidth) > (FormWidth - formToBuffer.VScroll.Width)) Then hScrollEnabled = True
    
    'We now know which scroll bars need to be enabled.  Before calculating scroll bar stuff, however, let's figure out where our viewport will
    ' be located - on the edge if scroll bars are enabled, or centered in the viewable area if scroll bars are NOT enabled.
    
    'Additionally, calculate viewport size - full form size if scroll bars enabled, full zoomed size if they are not
    Dim viewportLeft As Long, viewportTop As Long
    Dim viewportWidth As Long, viewportHeight As Long
    
    If hScrollEnabled Then
        viewportLeft = 0
        If Not vScrollEnabled Then
            viewportWidth = FormWidth
        Else
            viewportWidth = FormWidth - formToBuffer.VScroll.Width
        End If
    Else
        viewportWidth = zWidth
        If Not vScrollEnabled Then
            viewportLeft = (FormWidth - zWidth) / 2
        Else
            viewportLeft = ((FormWidth - formToBuffer.VScroll.Width) - zWidth) / 2
        End If
    End If
    
    If vScrollEnabled Then
        viewportTop = 0
        If Not hScrollEnabled Then
            viewportHeight = FormHeight
        Else
            viewportHeight = FormHeight - formToBuffer.HScroll.Height
        End If
    Else
        viewportHeight = zHeight
        If Not hScrollEnabled Then
            viewportTop = (FormHeight - zHeight) / 2
        Else
            viewportTop = ((FormHeight - formToBuffer.HScroll.Height) - zHeight) / 2
        End If
    End If
    
    'Now we know 1) which scrollbars are enabled, 2) the position of our viewport, 3) the size of our viewport.  Knowing this, we can now calculate
    ' the scroll bar values.
    
    'First - if no scroll bars are enabled, draw the viewport and exit.
    If (Not hScrollEnabled) And (Not vScrollEnabled) Then
    
        'Reset the scroll bar values so ScrollViewport doesn't assume we want scrolling
        formToBuffer.HScroll.Value = 0
        formToBuffer.VScroll.Value = 0
    
        'Hide the scroll bars if necessary
        If formToBuffer.HScroll.Visible Then formToBuffer.HScroll.Visible = False
        If formToBuffer.VScroll.Visible Then formToBuffer.VScroll.Visible = False
            
        'Resize the buffer and store the relevant painting information into this pdImages object
        pdImages(formToBuffer.Tag).backBuffer.createBlank FormWidth, FormHeight, 24, g_CanvasBackground
        pdImages(formToBuffer.Tag).targetLeft = viewportLeft
        pdImages(formToBuffer.Tag).targetTop = viewportTop
        pdImages(formToBuffer.Tag).targetWidth = viewportWidth
        pdImages(formToBuffer.Tag).targetHeight = viewportHeight
        
        'Pass control to the viewport renderer
        ScrollViewport formToBuffer
        
        Exit Sub
        
    End If
    
    'If we've reached this point, one or both scroll bars are enabled.  The time has come to calculate their values.
    'Horizontal scroll bar comes first.
    Dim newScrollMax As Long
    
    If hScrollEnabled Then
    
        'If zoomed-in, set the scroll bar range to the number of not visible pixels.
        If ZoomVal <= 1 Then
            newScrollMax = pdImages(formToBuffer.Tag).Width - Int(viewportWidth * g_Zoom.ZoomFactor(pdImages(formToBuffer.Tag).CurrentZoomValue) + 0.5)
        'If zoomed-out, use a modified formula (as there is no reason to scroll at sub-pixel levels.)
        Else
            newScrollMax = pdImages(formToBuffer.Tag).Width - Int(viewportWidth / g_Zoom.ZoomFactor(pdImages(formToBuffer.Tag).CurrentZoomValue) + 0.5)
        End If
        
        If formToBuffer.HScroll.Value > newScrollMax Then formToBuffer.HScroll.Value = newScrollMax
        formToBuffer.HScroll.Max = newScrollMax
        
        'As a convenience to the user, make the scroll bar's LargeChange parameter proportional to the scroll bar's new maximum value
        If formToBuffer.HScroll.Max > 7 Then formToBuffer.HScroll.LargeChange = formToBuffer.HScroll.Max \ 8
        
    End If
    
    'Same formula, but with width and height swapped for vertical scrolling
    If vScrollEnabled = True Then
    
        'If zoomed-in, set the scroll bar range to the number of not visible pixels.
        If ZoomVal <= 1 Then
            newScrollMax = pdImages(formToBuffer.Tag).Height - Int(viewportHeight * g_Zoom.ZoomFactor(pdImages(formToBuffer.Tag).CurrentZoomValue) + 0.5)
        'If zoomed-out, use a modified formula (as there is no reason to scroll at sub-pixel levels.)
        Else
            newScrollMax = pdImages(formToBuffer.Tag).Height - Int(viewportHeight / g_Zoom.ZoomFactor(pdImages(formToBuffer.Tag).CurrentZoomValue) + 0.5)
        End If
        
        If formToBuffer.VScroll.Value > newScrollMax Then formToBuffer.VScroll.Value = newScrollMax
        formToBuffer.VScroll.Max = newScrollMax
        
        'As a convenience to the user, make the scroll bar's LargeChange parameter proportional to the scroll bar's new maximum value
        If formToBuffer.VScroll.Max > 7 Then formToBuffer.VScroll.LargeChange = formToBuffer.VScroll.Max \ 8
        
    End If
    
    'Added to our list of "things we know" is the scroll bar maximum values (and they have already been set).
    ' As such, the time has come to render everything to the screen.
    
    'Horizontal scroll bar gets rendered first...
    If hScrollEnabled Then
        formToBuffer.HScroll.Move 0, FormHeight - formToBuffer.HScroll.Height, viewportWidth, formToBuffer.HScroll.Height
        If (Not formToBuffer.HScroll.Visible) Then formToBuffer.HScroll.Visible = True
    Else
        formToBuffer.HScroll.Value = 0
        If formToBuffer.HScroll.Visible Then formToBuffer.HScroll.Visible = False
    End If
    
    'Then vertical scroll bar...
    If vScrollEnabled Then
        formToBuffer.VScroll.Move FormWidth - formToBuffer.VScroll.Width, 0, formToBuffer.VScroll.Width, viewportHeight
        If (Not formToBuffer.VScroll.Visible) Then formToBuffer.VScroll.Visible = True
    Else
        formToBuffer.VScroll.Value = 0
        If formToBuffer.VScroll.Visible Then formToBuffer.VScroll.Visible = False
    End If
    
    'We don't actually render the image here; instead, we prepare the buffer (backBuffer) and store the relevant
    ' drawing variables to this pdImages object.  ScrollViewport (above) will handle the actual drawing.
    Dim newVWidth As Long, newVHeight As Long
    If hScrollEnabled Then newVWidth = viewportWidth Else newVWidth = FormWidth
    If vScrollEnabled Then newVHeight = viewportHeight Else newVHeight = FormHeight
    
    'Prepare the relevant back buffer
    If (Not pdImages(formToBuffer.Tag).backBuffer Is Nothing) Then pdImages(formToBuffer.Tag).backBuffer.eraseLayer
    pdImages(formToBuffer.Tag).backBuffer.createBlank newVWidth, newVHeight, 24, g_CanvasBackground
    
    pdImages(formToBuffer.Tag).targetLeft = viewportLeft
    pdImages(formToBuffer.Tag).targetTop = viewportTop
    pdImages(formToBuffer.Tag).targetWidth = viewportWidth
    pdImages(formToBuffer.Tag).targetHeight = viewportHeight
        
    'Pass control to the viewport renderer (found at the top of this module)
    ScrollViewport formToBuffer

    Exit Sub

ZoomErrorHandler:

    If Err = 480 Then
        pdMsgBox "There is not enough memory available to continue this operation.  Please free up system memory (RAM) and try again.  If the problem persists, reduce the zoom value and try again.", vbExclamation + vbOKOnly, "Not Enough Memory"
        SetProgBarVal 0
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
        frontBuffer.eraseLayer
        Set frontBuffer = Nothing
    End If
End Sub

'When the program is first loaded, we need to populate a number of viewport-related values.
Public Sub initializeViewportEngine()

    'This list of zoom values is (effectively) arbitrary.  I've based this list off similar lists (Paint.NET, GIMP)
    ' while including a few extra values for convenience's sake
    
    'Total number of available zoom values
    g_Zoom.ZoomCount = 25
    
    ReDim g_Zoom.ZoomArray(0 To g_Zoom.ZoomCount) As Double
    ReDim g_Zoom.ZoomFactor(0 To g_Zoom.ZoomCount) As Double
    
    'Manually create a list of user-friendly zoom values
    FormMain.CmbZoom.AddItem "3200%", 0
        g_Zoom.ZoomArray(0) = 32
        g_Zoom.ZoomFactor(0) = 32
        
    FormMain.CmbZoom.AddItem "2400%", 1
        g_Zoom.ZoomArray(1) = 24
        g_Zoom.ZoomFactor(1) = 24
        
    FormMain.CmbZoom.AddItem "1600%", 2
        g_Zoom.ZoomArray(2) = 16
        g_Zoom.ZoomFactor(2) = 16
        
    FormMain.CmbZoom.AddItem "1200%", 3
        g_Zoom.ZoomArray(3) = 12
        g_Zoom.ZoomFactor(3) = 12
        
    FormMain.CmbZoom.AddItem "800%", 4
        g_Zoom.ZoomArray(4) = 8
        g_Zoom.ZoomFactor(4) = 8
        
    FormMain.CmbZoom.AddItem "700%", 5
        g_Zoom.ZoomArray(5) = 7
        g_Zoom.ZoomFactor(5) = 7
        
    FormMain.CmbZoom.AddItem "600%", 6
        g_Zoom.ZoomArray(6) = 6
        g_Zoom.ZoomFactor(6) = 6
        
    FormMain.CmbZoom.AddItem "500%", 7
        g_Zoom.ZoomArray(7) = 5
        g_Zoom.ZoomFactor(7) = 5
        
    FormMain.CmbZoom.AddItem "400%", 8
        g_Zoom.ZoomArray(8) = 4
        g_Zoom.ZoomFactor(8) = 4
        
    FormMain.CmbZoom.AddItem "300%", 9
        g_Zoom.ZoomArray(9) = 3
        g_Zoom.ZoomFactor(9) = 3
        
    FormMain.CmbZoom.AddItem "200%", 10
        g_Zoom.ZoomArray(10) = 2
        g_Zoom.ZoomFactor(10) = 2
        
    FormMain.CmbZoom.AddItem "100%", 11
        g_Zoom.ZoomArray(11) = 1
        g_Zoom.ZoomFactor(11) = 1
        
    FormMain.CmbZoom.AddItem "75%", 12
        g_Zoom.ZoomArray(12) = 3 / 4
        g_Zoom.ZoomFactor(12) = 4 / 3
        
    FormMain.CmbZoom.AddItem "67%", 13
        g_Zoom.ZoomArray(13) = 2 / 3
        g_Zoom.ZoomFactor(13) = 3 / 2
        
    FormMain.CmbZoom.AddItem "50%", 14
        g_Zoom.ZoomArray(14) = 0.5
        g_Zoom.ZoomFactor(14) = 2
        
    FormMain.CmbZoom.AddItem "33%", 15
        g_Zoom.ZoomArray(15) = 1 / 3
        g_Zoom.ZoomFactor(15) = 3
        
    FormMain.CmbZoom.AddItem "25%", 16
        g_Zoom.ZoomArray(16) = 0.25
        g_Zoom.ZoomFactor(16) = 4
        
    FormMain.CmbZoom.AddItem "20%", 17
        g_Zoom.ZoomArray(17) = 0.2
        g_Zoom.ZoomFactor(17) = 5
        
    FormMain.CmbZoom.AddItem "16%", 18
        g_Zoom.ZoomArray(18) = 0.16
        g_Zoom.ZoomFactor(18) = 100 / 16
        
    FormMain.CmbZoom.AddItem "12%", 19
        g_Zoom.ZoomArray(19) = 0.12
        g_Zoom.ZoomFactor(19) = 100 / 12
        
    FormMain.CmbZoom.AddItem "8%", 20
        g_Zoom.ZoomArray(20) = 0.08
        g_Zoom.ZoomFactor(20) = 100 / 8
        
    FormMain.CmbZoom.AddItem "6%", 21
        g_Zoom.ZoomArray(21) = 0.06
        g_Zoom.ZoomFactor(21) = 100 / 6
        
    FormMain.CmbZoom.AddItem "4%", 22
        g_Zoom.ZoomArray(22) = 0.04
        g_Zoom.ZoomFactor(22) = 25
        
    FormMain.CmbZoom.AddItem "3%", 23
        g_Zoom.ZoomArray(23) = 0.03
        g_Zoom.ZoomFactor(23) = 100 / 0.03
        
    FormMain.CmbZoom.AddItem "2%", 24
        g_Zoom.ZoomArray(24) = 0.02
        g_Zoom.ZoomFactor(24) = 50
        
    FormMain.CmbZoom.AddItem "1%", 25
        g_Zoom.ZoomArray(25) = 0.01
        g_Zoom.ZoomFactor(25) = 100
    
    'Set the main form's zoom combo box to display "100%"
    FormMain.CmbZoom.ListIndex = ZoomIndex100

End Sub
