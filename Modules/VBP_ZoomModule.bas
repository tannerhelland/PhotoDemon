Attribute VB_Name = "Zoom_Handler"
'***************************************************************************
'Zoom Handler - builds and draws the image viewport and associated scroll bars
'Copyright ©2000-2012 by Tanner Helland
'Created: 4/15/01
'Last updated: 10/September/12
'Last update: Fixed a scrollbar reset problem when switching between an extremely zoomed-in view and 100% zoom.
'Still needs: option to draw borders around the image
'
'Module for handling the "zoom" feature on the main form.  There are two routines - 'PrepareViewport' for rebuilding all related objects
' (done only when the zoom value is changed or a new picture is loaded) and 'ScrollViewport' (when the view is scrolled but the zoom
' variables don't change).  StretchBlt is used for the actual rendering, and its "halftone" mode is explicitly specified for shrinking the image.
'
'***************************************************************************

Option Explicit

'This is the ListIndex of the FormMain zoom combo box that corresponds to 100%
Public Const zoomIndex100 As Long = 11

'Width and height values of the image AFTER zoom has been applied.  (For example, if the image is 100x100
' and the zoom value is 200%, zWidth and zHeight will be 200.)
Dim zWidth As Single, zHeight As Single

'32bpp images require pre-multiplication against a background (otherwise it will be black).  To make sure that the original alpha
' channel is correctly preserved, we must reapply this pre-multiplication every time we draw the image to screen.  This object will
' hold the object used to perform the pre-multiplication.
Dim alphaFixLayer As pdLayer

'These variables represent the source width - e.g. the size of the viewable picture box, divided by the zoom coefficient
Dim SrcWidth As Double, SrcHeight As Double

'The zoom value is the actual coefficient for the current zoom value.  (For example, 0.50 for "50% zoom")
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
    Set frontBuffer = New pdLayer
    
    'Copy the current back buffer into the front buffer
    frontBuffer.createFromExistingLayer pdImages(formToBuffer.Tag).backBuffer

    'Check to see if a selection is active.
    If pdImages(formToBuffer.Tag).selectionActive Then
    
        'If it is, composite the selection against the temporary buffer
        pdImages(formToBuffer.Tag).mainSelection.renderCustom frontBuffer, formToBuffer, pdImages(formToBuffer.Tag).targetLeft, pdImages(formToBuffer.Tag).targetTop, pdImages(formToBuffer.Tag).targetWidth, pdImages(formToBuffer.Tag).targetHeight, selectionRenderPreference
    
    End If
        
    'If the user has requested a drop shadow drawn onto the canvas, handle that next
    If CanvasDropShadow Then
    
        'We'll handle this in two steps; first, the horizontal stretches
        If formToBuffer.VScroll.Visible = False Then
                    
            'Make sure the image isn't snugly fit inside the viewport; if it is, this is a waste of time
            If pdImages(formToBuffer.Tag).targetTop <> 0 Then
                'Top edge
                StretchBlt frontBuffer.getLayerDC, pdImages(formToBuffer.Tag).targetLeft, pdImages(formToBuffer.Tag).targetTop - PD_CANVASSHADOWSIZE, pdImages(formToBuffer.Tag).targetWidth, PD_CANVASSHADOWSIZE, canvasShadow.getShadowDC(0), 0, 0, 1, PD_CANVASSHADOWSIZE, vbSrcCopy
                'Bottom edge
                StretchBlt frontBuffer.getLayerDC, pdImages(formToBuffer.Tag).targetLeft, pdImages(formToBuffer.Tag).targetTop + pdImages(formToBuffer.Tag).targetHeight, pdImages(formToBuffer.Tag).targetWidth, PD_CANVASSHADOWSIZE, canvasShadow.getShadowDC(1), 0, 0, 1, PD_CANVASSHADOWSIZE, vbSrcCopy
            End If
        
        End If
        
        'Second, the vertical stretches
        If formToBuffer.HScroll.Visible = False Then
                    
            'Make sure the image isn't snugly fit inside the viewport; if it is, this is a waste of time
            If pdImages(formToBuffer.Tag).targetLeft <> 0 Then
                'Left edge
                StretchBlt frontBuffer.getLayerDC, pdImages(formToBuffer.Tag).targetLeft - PD_CANVASSHADOWSIZE, pdImages(formToBuffer.Tag).targetTop, PD_CANVASSHADOWSIZE, pdImages(formToBuffer.Tag).targetHeight, canvasShadow.getShadowDC(2), 0, 0, PD_CANVASSHADOWSIZE, 1, vbSrcCopy
                'Right edge
                StretchBlt frontBuffer.getLayerDC, pdImages(formToBuffer.Tag).targetLeft + pdImages(formToBuffer.Tag).targetWidth, pdImages(formToBuffer.Tag).targetTop, PD_CANVASSHADOWSIZE, pdImages(formToBuffer.Tag).targetHeight, canvasShadow.getShadowDC(3), 0, 0, PD_CANVASSHADOWSIZE, 1, vbSrcCopy
            End If
        
        End If
        
        'Finally, the corners, which are only drawn if both scroll bars are invisible
        If (formToBuffer.VScroll.Visible = False) And (formToBuffer.HScroll.Visible = False) Then
        
            'NW corner
            StretchBlt frontBuffer.getLayerDC, pdImages(formToBuffer.Tag).targetLeft - PD_CANVASSHADOWSIZE, pdImages(formToBuffer.Tag).targetTop - PD_CANVASSHADOWSIZE, PD_CANVASSHADOWSIZE, PD_CANVASSHADOWSIZE, canvasShadow.getShadowDC(4), 0, 0, PD_CANVASSHADOWSIZE, PD_CANVASSHADOWSIZE, vbSrcCopy
            'NE corner
            StretchBlt frontBuffer.getLayerDC, pdImages(formToBuffer.Tag).targetLeft + pdImages(formToBuffer.Tag).targetWidth, pdImages(formToBuffer.Tag).targetTop - PD_CANVASSHADOWSIZE, PD_CANVASSHADOWSIZE, PD_CANVASSHADOWSIZE, canvasShadow.getShadowDC(5), 0, 0, PD_CANVASSHADOWSIZE, PD_CANVASSHADOWSIZE, vbSrcCopy
            'SW corner
            StretchBlt frontBuffer.getLayerDC, pdImages(formToBuffer.Tag).targetLeft - PD_CANVASSHADOWSIZE, pdImages(formToBuffer.Tag).targetTop + pdImages(formToBuffer.Tag).targetHeight, PD_CANVASSHADOWSIZE, PD_CANVASSHADOWSIZE, canvasShadow.getShadowDC(6), 0, 0, PD_CANVASSHADOWSIZE, PD_CANVASSHADOWSIZE, vbSrcCopy
            'SE corner
            StretchBlt frontBuffer.getLayerDC, pdImages(formToBuffer.Tag).targetLeft + pdImages(formToBuffer.Tag).targetWidth, pdImages(formToBuffer.Tag).targetTop + pdImages(formToBuffer.Tag).targetHeight, PD_CANVASSHADOWSIZE, PD_CANVASSHADOWSIZE, canvasShadow.getShadowDC(7), 0, 0, PD_CANVASSHADOWSIZE, PD_CANVASSHADOWSIZE, vbSrcCopy
        
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
    
    'The zoom value is the actual coefficient for the current zoom value.  (For example, 0.50 for "50% zoom")
    ZoomVal = Zoom.ZoomArray(pdImages(formToBuffer.Tag).CurrentZoomValue)

    'These variables represent the source width - e.g. the size of the viewable picture box, divided by the zoom coefficient
    SrcWidth = pdImages(formToBuffer.Tag).targetWidth / ZoomVal
    SrcHeight = pdImages(formToBuffer.Tag).targetHeight / ZoomVal
    
    'These variables are the offset, as determined by the scroll bar values
    If formToBuffer.HScroll.Enabled Then srcX = formToBuffer.HScroll.Value Else srcX = 0
    If formToBuffer.VScroll.Enabled Then srcY = formToBuffer.VScroll.Value Else srcY = 0
        
    'Paint the image from the back buffer to the front buffer
    If ZoomVal < 1 Then
        
        'Check for alpha channel.  If it's found, perform pre-multiplication against a white background before rendering.
        If pdImages(formToBuffer.Tag).mainLayer.getLayerColorDepth = 32 Then
            'pdImages(formToBuffer.Tag).alphaFixLayer.createBlank srcWidth, srcHeight, 32
            'BitBlt pdImages(formToBuffer.Tag).alphaFixLayer.getLayerDC, 0, 0, srcWidth, srcHeight, pdImages(formToBuffer.Tag).mainLayer.getLayerDC, srcX, srcY, vbSrcCopy
            'pdImages(formToBuffer.Tag).alphaFixLayer.compositeBackgroundColor
            'StretchBlt pdImages(formToBuffer.Tag).backBuffer.getLayerDC, pdImages(formToBuffer.Tag).targetLeft, pdImages(formToBuffer.Tag).targetTop, pdImages(formToBuffer.Tag).targetWidth, pdImages(formToBuffer.Tag).targetHeight, pdImages(formToBuffer.Tag).alphaFixLayer.getLayerDC(), 0, 0, srcWidth, srcHeight, vbSrcCopy
            
            pdImages(formToBuffer.Tag).alphaFixLayer.createBlank pdImages(formToBuffer.Tag).targetWidth, pdImages(formToBuffer.Tag).targetHeight, 32

            'Now comes a nasty hack; halftone stretching does not preserve the alpha channel, but coloroncolor does.  So make two copies - one with
            ' color-on-color, from which we'll steal alpha values, and a high-quality halftone one for pixel values.
            Dim hackLayer As pdLayer
            Set hackLayer = New pdLayer
            hackLayer.createBlank pdImages(formToBuffer.Tag).targetWidth, pdImages(formToBuffer.Tag).targetHeight, 32
            
            SetStretchBltMode hackLayer.getLayerDC, STRETCHBLT_COLORONCOLOR
            StretchBlt hackLayer.getLayerDC, 0, 0, pdImages(formToBuffer.Tag).targetWidth, pdImages(formToBuffer.Tag).targetHeight, pdImages(formToBuffer.Tag).mainLayer.getLayerDC, srcX, srcY, SrcWidth, SrcHeight, vbSrcCopy
            
            SetStretchBltMode pdImages(formToBuffer.Tag).alphaFixLayer.getLayerDC, STRETCHBLT_HALFTONE
            StretchBlt pdImages(formToBuffer.Tag).alphaFixLayer.getLayerDC, 0, 0, pdImages(formToBuffer.Tag).targetWidth, pdImages(formToBuffer.Tag).targetHeight, pdImages(formToBuffer.Tag).mainLayer.getLayerDC, srcX, srcY, SrcWidth, SrcHeight, vbSrcCopy
            pdImages(formToBuffer.Tag).alphaFixLayer.compositeBackgroundColorSpecial hackLayer
            BitBlt pdImages(formToBuffer.Tag).backBuffer.getLayerDC, pdImages(formToBuffer.Tag).targetLeft, pdImages(formToBuffer.Tag).targetTop, pdImages(formToBuffer.Tag).targetWidth, pdImages(formToBuffer.Tag).targetHeight, pdImages(formToBuffer.Tag).alphaFixLayer.getLayerDC, 0, 0, vbSrcCopy
            hackLayer.eraseLayer
        Else
            SetStretchBltMode pdImages(formToBuffer.Tag).backBuffer.getLayerDC, STRETCHBLT_HALFTONE
            StretchBlt pdImages(formToBuffer.Tag).backBuffer.getLayerDC, pdImages(formToBuffer.Tag).targetLeft, pdImages(formToBuffer.Tag).targetTop, pdImages(formToBuffer.Tag).targetWidth, pdImages(formToBuffer.Tag).targetHeight, pdImages(formToBuffer.Tag).mainLayer.getLayerDC(), srcX, srcY, SrcWidth, SrcHeight, vbSrcCopy
        End If
        
    Else
        'When zoomed in, the blitting call must be modified as follows: restrict it to multiples of the current zoom factor.
        ' (Without this fix, funny stretching occurs; to see it yourself, place the zoom at 300%, and drag an image's window larger or smaller.)
        Dim bltWidth As Long, bltHeight As Long
        bltWidth = pdImages(formToBuffer.Tag).targetWidth + (Int(Zoom.ZoomFactor(pdImages(formToBuffer.Tag).CurrentZoomValue)) - (pdImages(formToBuffer.Tag).targetWidth Mod Int(Zoom.ZoomFactor(pdImages(formToBuffer.Tag).CurrentZoomValue))))
        SrcWidth = bltWidth / ZoomVal
        bltHeight = pdImages(formToBuffer.Tag).targetHeight + (Int(Zoom.ZoomFactor(pdImages(formToBuffer.Tag).CurrentZoomValue)) - (pdImages(formToBuffer.Tag).targetHeight Mod Int(Zoom.ZoomFactor(pdImages(formToBuffer.Tag).CurrentZoomValue))))
        SrcHeight = bltHeight / ZoomVal
        
        'Check for alpha channel.  If it's found, perform pre-multiplication against a white background before rendering.
        If pdImages(formToBuffer.Tag).mainLayer.getLayerColorDepth = 32 Then
            'pdImages(formToBuffer.Tag).alphaFixLayer.createBlank SrcWidth, SrcHeight, 32
            'BitBlt pdImages(formToBuffer.Tag).alphaFixLayer.getLayerDC, 0, 0, SrcWidth, SrcHeight, pdImages(formToBuffer.Tag).mainLayer.getLayerDC, srcX, srcY, vbSrcCopy
            'pdImages(formToBuffer.Tag).alphaFixLayer.compositeBackgroundColor
            'StretchBlt pdImages(formToBuffer.Tag).backBuffer.getLayerDC, pdImages(formToBuffer.Tag).targetLeft, pdImages(formToBuffer.Tag).targetTop, bltWidth, bltHeight, pdImages(formToBuffer.Tag).alphaFixLayer.getLayerDC(), 0, 0, SrcWidth, SrcHeight, vbSrcCopy
            pdImages(formToBuffer.Tag).alphaFixLayer.createBlank bltWidth, bltHeight, 32
            SetStretchBltMode pdImages(formToBuffer.Tag).alphaFixLayer.getLayerDC, STRETCHBLT_COLORONCOLOR
            StretchBlt pdImages(formToBuffer.Tag).alphaFixLayer.getLayerDC, 0, 0, bltWidth, bltHeight, pdImages(formToBuffer.Tag).mainLayer.getLayerDC(), srcX, srcY, SrcWidth, SrcHeight, vbSrcCopy
            pdImages(formToBuffer.Tag).alphaFixLayer.compositeBackgroundColor
            BitBlt pdImages(formToBuffer.Tag).backBuffer.getLayerDC, pdImages(formToBuffer.Tag).targetLeft, pdImages(formToBuffer.Tag).targetTop, pdImages(formToBuffer.Tag).targetWidth, pdImages(formToBuffer.Tag).targetHeight, pdImages(formToBuffer.Tag).alphaFixLayer.getLayerDC, 0, 0, vbSrcCopy
        Else
            SetStretchBltMode pdImages(formToBuffer.Tag).backBuffer.getLayerDC, STRETCHBLT_COLORONCOLOR
            StretchBlt pdImages(formToBuffer.Tag).backBuffer.getLayerDC, pdImages(formToBuffer.Tag).targetLeft, pdImages(formToBuffer.Tag).targetTop, bltWidth, bltHeight, pdImages(formToBuffer.Tag).mainLayer.getLayerDC, srcX, srcY, SrcWidth, SrcHeight, vbSrcCopy
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

    'Don't attempt to resize the scroll bars if FixScrolling is disabled. Yhis is used to provide a smoother user experience,
    ' especially when images are being loaded. (This routine is triggered on Form_Resize, which is in turn triggered when a
    ' new picture is loaded.  To prevent PrepareViewport from being fired multiple times, FixScrolling is utilized.)
    If FixScrolling = False Then Exit Sub
    
    'Make sure the form is valid
    If formToBuffer Is Nothing Then Exit Sub
    
    'If the image associated with this form is inactive, ignore this request
    If pdImages(formToBuffer.Tag).IsActive = False Then Exit Sub
    
    'Because this routine is time-consuming, I track it carefully to try and minimize how frequently it's called.  Feel free to comment out this line.
    Debug.Print "Preparing viewport: " & reasonForRedraw & " | (" & formToBuffer.Tag & ") | " & formToBuffer.Caption
    
    On Error GoTo ZoomErrorHandler
    
    'Get the mathematical zoom multiplier (based on the current combo box setting - for example, 0.50 for "50% zoom")
    Dim ZoomVal As Double
    ZoomVal = Zoom.ZoomArray(pdImages(formToBuffer.Tag).CurrentZoomValue)
    
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
    ' be located - on the edge if scroll bars are enabled, centered in the viewable area if scroll bars are not enabled.
    
    'Similarly, calculate viewport size - full form size if scroll bars enabled, full zoom size if they are not
    Dim viewportLeft As Long, viewportTop As Long
    Dim viewportWidth As Long, viewportHeight As Long
    
    If hScrollEnabled = True Then
        viewportLeft = 0
        If vScrollEnabled = False Then
            viewportWidth = FormWidth
        Else
            viewportWidth = FormWidth - formToBuffer.VScroll.Width
        End If
    Else
        viewportWidth = zWidth
        If vScrollEnabled = False Then
            viewportLeft = (FormWidth - zWidth) / 2
        Else
            viewportLeft = ((FormWidth - formToBuffer.VScroll.Width) - zWidth) / 2
        End If
    End If
    
    If vScrollEnabled = True Then
        viewportTop = 0
        If hScrollEnabled = False Then
            viewportHeight = FormHeight
        Else
            viewportHeight = FormHeight - formToBuffer.HScroll.Height
        End If
    Else
        viewportHeight = zHeight
        If hScrollEnabled = False Then
            viewportTop = (FormHeight - zHeight) / 2
        Else
            viewportTop = ((FormHeight - formToBuffer.HScroll.Height) - zHeight) / 2
        End If
    End If
    
    'Now we know 1) which scrollbars are enabled, 2) the position of our viewport, 3) the size of our viewport.  Knowing this, we can now calculate
    ' the scroll bar values.
    
    'First - if no scroll bars are enabled, draw the viewport and exit.
    If hScrollEnabled = False And vScrollEnabled = False Then
    
        'Reset the scroll bar values so ScrollViewport doesn't assume we want scrolling
        formToBuffer.HScroll.Value = 0
        formToBuffer.VScroll.Value = 0
    
        'Hide the scroll bars if necessary
        If formToBuffer.HScroll.Visible = True Then formToBuffer.HScroll.Visible = False
        If formToBuffer.VScroll.Visible = True Then formToBuffer.VScroll.Visible = False
            
        'Resize the buffer and store the relevant painting information into this pdImages object
        pdImages(formToBuffer.Tag).backBuffer.createBlank FormWidth, FormHeight, 24, CanvasBackground
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
    If hScrollEnabled = True Then
    
        'If zoomed-in, set the scroll bar range to the number of not visible pixels.
        If ZoomVal <= 1 Then
            formToBuffer.HScroll.Max = pdImages(formToBuffer.Tag).Width - Int(viewportWidth * Zoom.ZoomFactor(pdImages(formToBuffer.Tag).CurrentZoomValue) + 0.5)
        'If zoomed-out, use a modified formula (as there is no reason to scroll at sub-pixel levels.)
        Else
            formToBuffer.HScroll.Max = pdImages(formToBuffer.Tag).Width - Int(viewportWidth / Zoom.ZoomFactor(pdImages(formToBuffer.Tag).CurrentZoomValue) + 0.5)
        End If
        
        'As a convenience to the user, make the scroll bar's LargeChange parameter proportional to the scroll bar's new maximum value
        If formToBuffer.HScroll.Max > 7 Then formToBuffer.HScroll.LargeChange = formToBuffer.HScroll.Max \ 8
        
    End If
    
    'Same formula, but with width and height swapped for vertical scrolling
    If vScrollEnabled = True Then
    
        'If zoomed-in, set the scroll bar range to the number of not visible pixels.
        If ZoomVal <= 1 Then
            formToBuffer.VScroll.Max = pdImages(formToBuffer.Tag).Height - Int(viewportHeight * Zoom.ZoomFactor(pdImages(formToBuffer.Tag).CurrentZoomValue) + 0.5)
        'If zoomed-out, use a modified formula (as there is no reason to scroll at sub-pixel levels.)
        Else
            formToBuffer.VScroll.Max = pdImages(formToBuffer.Tag).Height - Int(viewportHeight / Zoom.ZoomFactor(pdImages(formToBuffer.Tag).CurrentZoomValue) + 0.5)
        End If
        
        'As a convenience to the user, make the scroll bar's LargeChange parameter proportional to the scroll bar's new maximum value
        If formToBuffer.VScroll.Max > 7 Then formToBuffer.VScroll.LargeChange = formToBuffer.VScroll.Max \ 8
        
    End If
    
    'Added to our list of "things we know" is the scroll bar maximum values (and they have already been set).
    ' As such, the time has come to render everything to the screen.
    
    'Horizontal scroll bar gets rendered first...
    If hScrollEnabled = True Then
        formToBuffer.HScroll.Move 0, FormHeight - formToBuffer.HScroll.Height, viewportWidth, formToBuffer.HScroll.Height
        If formToBuffer.HScroll.Visible = False Then formToBuffer.HScroll.Visible = True
    Else
        formToBuffer.HScroll.Value = 0
        If formToBuffer.HScroll.Visible = True Then formToBuffer.HScroll.Visible = False
    End If
    
    'Then vertical scroll bar...
    If vScrollEnabled = True Then
        formToBuffer.VScroll.Move FormWidth - formToBuffer.VScroll.Width, 0, formToBuffer.VScroll.Width, viewportHeight
        If formToBuffer.VScroll.Visible = False Then formToBuffer.VScroll.Visible = True
    Else
        formToBuffer.VScroll.Value = 0
        If formToBuffer.VScroll.Visible = True Then formToBuffer.VScroll.Visible = False
    End If
    
    'We don't actually render the image here; instead, we prepare the buffer (backBuffer) and store the relevant
    ' drawing variables to this pdImages object.  ScrollViewport (above) will handle the actual drawing.
    Dim newVWidth As Long, newVHeight As Long
    If hScrollEnabled = True Then newVWidth = viewportWidth Else newVWidth = FormWidth
    If vScrollEnabled = True Then newVHeight = viewportHeight Else newVHeight = FormHeight
    pdImages(formToBuffer.Tag).backBuffer.createBlank newVWidth, newVHeight, 24, CanvasBackground
    
    pdImages(formToBuffer.Tag).targetLeft = viewportLeft
    pdImages(formToBuffer.Tag).targetTop = viewportTop
    pdImages(formToBuffer.Tag).targetWidth = viewportWidth
    pdImages(formToBuffer.Tag).targetHeight = viewportHeight
        
    'Pass control to the viewport renderer (found at the top of this module)
    ScrollViewport formToBuffer

    Exit Sub

ZoomErrorHandler:

    If Err = 480 Then
        MsgBox "There is not enough memory available to continue this operation.  Please free up system memory (RAM) and try again.  If the problem persists, reduce the zoom value and try again.", vbCritical + vbOKOnly, "Not Enough Memory"
        SetProgBarVal 0
        Message "Operation halted."
    ElseIf Err = 13 Then
        Message "Invalid zoom value."
        Exit Sub
    Else
        Message "Zoom paused due to unexpected error (#" & Err & ")."
        Exit Sub
    End If

End Sub
