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

'ScrollViewport is used to update the on-screen image when the scroll bars are used.
' Given how frequently it is used, I've tried to make it as small and fast as possible.
Public Sub ScrollViewport(ByRef formToBuffer As Form)
    
    '32bpp images require pre-multiplication against a white background (otherwise it will be black).  To make sure that the original alpha
    ' channel is correctly preserved, we must reapply this pre-multiplication every time we draw the image to screen.  This object will
    ' hold the object used to perform the pre-multiplication.
    If pdImages(formToBuffer.Tag).mainLayer.getLayerColorDepth = 32 Then
        Dim alphaFixLayer As pdLayer
        Set alphaFixLayer = New pdLayer
    End If
    
    'The zoom value is the actual coefficient for the current zoom value.  (For example, 0.50 for "50% zoom")
    Dim ZoomVal As Double
    ZoomVal = Zoom.ZoomArray(pdImages(formToBuffer.Tag).CurrentZoomValue)

    'These variables represent the source width - e.g. the size of the viewable picture box, divided by the zoom coefficient
    'Dim SrcWidth As Double, SrcHeight As Double
    Dim SrcWidth As Double, SrcHeight As Double
    SrcWidth = pdImages(formToBuffer.Tag).targetWidth / ZoomVal
    SrcHeight = pdImages(formToBuffer.Tag).targetHeight / ZoomVal
    
    'These variables are the offset, as determined by the scroll bar values
    Dim srcX As Long, srcY As Long
    srcX = formToBuffer.HScroll.Value
    srcY = formToBuffer.VScroll.Value

    'Prepare the background (checkerboard or color, per the user's setting in Edit -> Preferences)
    DrawSpecificCanvas formToBuffer

    'When zoomed out, specify halftone mode (for limited resampling).  Otherwise, nearest-neighbor sampling is fine.
    If ZoomVal >= 1 Then
        SetStretchBltMode pdImages(formToBuffer.Tag).backBuffer.getLayerDC, STRETCHBLT_COLORONCOLOR
    Else
        SetStretchBltMode pdImages(formToBuffer.Tag).backBuffer.getLayerDC, STRETCHBLT_HALFTONE
    End If
    
    'Paint the image from the back buffer to the front buffer
    If ZoomVal <= 1 Then
        
        'Check for alpha channel.  If it's found, perform pre-multiplication against a white background before rendering.
        If pdImages(formToBuffer.Tag).mainLayer.getLayerColorDepth = 32 Then
            alphaFixLayer.createBlank SrcWidth, SrcHeight, 32
            BitBlt alphaFixLayer.getLayerDC, 0, 0, SrcWidth, SrcHeight, pdImages(formToBuffer.Tag).mainLayer.getLayerDC, srcX, srcY, vbSrcCopy
            alphaFixLayer.compositeBackgroundColor
            StretchBlt pdImages(formToBuffer.Tag).backBuffer.getLayerDC, pdImages(formToBuffer.Tag).targetLeft, pdImages(formToBuffer.Tag).targetTop, pdImages(formToBuffer.Tag).targetWidth, pdImages(formToBuffer.Tag).targetHeight, alphaFixLayer.getLayerDC(), 0, 0, SrcWidth, SrcHeight, vbSrcCopy
        Else
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
            alphaFixLayer.createBlank SrcWidth, SrcHeight, 32
            BitBlt alphaFixLayer.getLayerDC, 0, 0, SrcWidth, SrcHeight, pdImages(formToBuffer.Tag).mainLayer.getLayerDC, srcX, srcY, vbSrcCopy
            alphaFixLayer.compositeBackgroundColor
            StretchBlt pdImages(formToBuffer.Tag).backBuffer.getLayerDC, pdImages(formToBuffer.Tag).targetLeft, pdImages(formToBuffer.Tag).targetTop, bltWidth, bltHeight, alphaFixLayer.getLayerDC(), 0, 0, SrcWidth, SrcHeight, vbSrcCopy
        Else
            StretchBlt pdImages(formToBuffer.Tag).backBuffer.getLayerDC, pdImages(formToBuffer.Tag).targetLeft, pdImages(formToBuffer.Tag).targetTop, bltWidth, bltHeight, pdImages(formToBuffer.Tag).mainLayer.getLayerDC, srcX, srcY, SrcWidth, SrcHeight, vbSrcCopy
        End If
        
    End If
    
    'Next, check to see if a selection is active.
    If pdImages(formToBuffer.Tag).selectionActive Then
    
        'If it is, check to see if it's locked in
        'If pdImages(formToBuffer.Tag).mainSelection.isLockedIn Then
        '    formToBuffer.FrontBuffer.Picture = formToBuffer.FrontBuffer.Image
        '    pdImages(formToBuffer.Tag).mainSelection.renderFinal formToBuffer, pdImages(formToBuffer.Tag).targetLeft, pdImages(formToBuffer.Tag).targetTop, pdImages(formToBuffer.Tag).targetWidth, pdImages(formToBuffer.Tag).targetHeight
        'Else
        '    pdImages(formToBuffer.Tag).mainSelection.renderIntermediate formToBuffer, pdImages(formToBuffer.Tag).targetLeft, pdImages(formToBuffer.Tag).targetTop
        'End If
    
    End If
        
    'Flip the front buffer to the screen
    formToBuffer.Picture = LoadPicture("")
    BitBlt formToBuffer.hDC, 0, 0, pdImages(formToBuffer.Tag).backBuffer.getLayerWidth, pdImages(formToBuffer.Tag).backBuffer.getLayerHeight, pdImages(formToBuffer.Tag).backBuffer.getLayerDC, 0, 0, vbSrcCopy
    formToBuffer.Picture = formToBuffer.Image
    'formToBuffer.Refresh
    
    'If we don't fire DoEvents here, the image will only scroll after the mouse button is released.
    DoEvents
    
    'Delete the temporary rendering image used for premultiplication
    If pdImages(formToBuffer.Tag).mainLayer.getLayerColorDepth = 32 Then
        alphaFixLayer.eraseLayer
        Set alphaFixLayer = Nothing
    End If

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
    zWidth = (pdImages(CurrentImage).Width * ZoomVal)
    zHeight = (pdImages(CurrentImage).Height * ZoomVal)
    
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
            formToBuffer.HScroll.Max = pdImages(CurrentImage).Width - Int(viewportWidth * Zoom.ZoomFactor(pdImages(formToBuffer.Tag).CurrentZoomValue) + 0.5)
        'If zoomed-out, use a modified formula (as there is no reason to scroll at sub-pixel levels.)
        Else
            formToBuffer.HScroll.Max = pdImages(CurrentImage).Width - Int(viewportWidth / Zoom.ZoomFactor(pdImages(formToBuffer.Tag).CurrentZoomValue) + 0.5)
        End If
        
        'As a convenience to the user, make the scroll bar's LargeChange parameter proportional to the scroll bar's new maximum value
        If formToBuffer.HScroll.Max > 7 Then formToBuffer.HScroll.LargeChange = formToBuffer.HScroll.Max \ 8
        
    End If
    
    'Same formula, but with width and height swapped for vertical scrolling
    If vScrollEnabled = True Then
    
        'If zoomed-in, set the scroll bar range to the number of not visible pixels.
        If ZoomVal <= 1 Then
            formToBuffer.VScroll.Max = pdImages(CurrentImage).Height - Int(viewportHeight * Zoom.ZoomFactor(pdImages(formToBuffer.Tag).CurrentZoomValue) + 0.5)
        'If zoomed-out, use a modified formula (as there is no reason to scroll at sub-pixel levels.)
        Else
            formToBuffer.VScroll.Max = pdImages(CurrentImage).Height - Int(viewportHeight / Zoom.ZoomFactor(pdImages(formToBuffer.Tag).CurrentZoomValue) + 0.5)
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
    
    'We don't actually render the image here; instead, we prepare the buffer (vBackBuffer) and store the relevant
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
