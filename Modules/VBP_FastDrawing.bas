Attribute VB_Name = "FastDrawing"
'***************************************************************************
'Fast API Graphics Routines Interface
'Copyright ©2001-2013 by Tanner Helland
'Created: 12/June/01
'Last updated: 15/September/13
'Last update: during a preview, clients should not have to track and calculate the difference between preview dimensions
'              and actual image dimensions (important for things like Blur, where the preview radius must be reduced in
'              order to provide an accurate preview).  PrepImageData now does this for them, and they can simply access
'              the .previewModifier value as necessary.
'
'This interface provides API support for the main image interaction routines. It assigns memory data
' into a useable array, and later transfers that array back into memory.  Very fast, very compact, can't
' live without it. These functions are arguably the most integral part of PhotoDemon.
'
'If you want to know more about how DIB sections work - and why they're so fast compared to VB's internal
' .PSet and .Point methods - please visit http://www.tannerhelland.com/42/vb-graphics-programming-3/
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

Private Declare Function VarPtrArray Lib "msvbvm60" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Long, lpSrc As Long, ByVal byteLength As Long)

'Any time a tool dialog is in use, the image to be operated on will be stored IN THIS LAYER.
'- In preview mode, workingLayer will contain a small, preview-size version of the image.
'- In non-preview mode, workingLayer will contain a copy of the full image.  We do not allow various effects and tools to operate
'   on the original image data, in case they cancel a function mid-way (we must be able to restore the original image).
'- If a selection is active, workingLayer will contain the selected part of the image, converted to 32bpp mode as necessary
'   (e.g. if feathering or antialiasing is enabled on the selection).
Public workingLayer As pdLayer

'When the workingLayer is first created, we store a backup of it in this layer.  This backup is used to rebuild the full image
' while accounting for any selected areas; we merge the selected areas onto this original copy, then copy the composited result
' back onto the image.  This is easier than attempting to merge the area onto the main layer while doing all our extra
' selection processing.
Private workingLayerBackup As pdLayer

'prepImageData is the function all PhotoDemon tools call when they need a copy of the image to operate on.  That function fills a
' variable of this type (FilterInfo) with everything a filter or effect could possibly want to know about the current layer.
' Note that filters are free to ignore this data, but it is ALWAYS populated and made available.
Public Type FilterInfo
    
    'Coordinates of the top-left location the filter is supposed to operate on
    Left As Long
    Top As Long
    
    'Note that Right and Bottom could be inferred from Left + Width and Top + Height, but we
    ' explicitly state them to minimize effort on the calling routine's part
    Right As Long
    Bottom As Long
    
    'Dimensions of the area the filter is supposed to operate on.  (Relevant if a selection is active.)
    Width As Long
    Height As Long
    
    'The lowest coordinate the filter is allowed to check.  This is almost always the top-left of the image (0, 0).
    minX As Long
    MinY As Long
    
    'The highest coordinate the filter is allowed to check.  This is almost always (width, height).
    maxX As Long
    MaxY As Long
    
    'The colorDepth of the current layer, specified as BITS per pixel; this will always be 24 or 32
    colorDepth As Long
    
    'BytesPerPixel is simply colorDepth / 8.  It is provided for convenience.
    BytesPerPixel As Long
    
    'Filters shouldn't have to worry about where the layer is physically located, but when it comes
    ' time to set the layer back in place, knowing the layer's location on the primary image may be
    ' useful (as when previewing, for example)
    LayerX As Long
    LayerY As Long
    
    'When in preview mode, the on-screen image will typically be represented at a smaller-than-actual size.
    ' If an effect or filter operates on a radius (e.g. "blur radius 20"), that radius value has to be shrunk
    ' when working on the preview - otherwise, the preview effect will look much stronger than it actually is!
    ' This value can be multiplied by a radius or other value but ONLY WHEN PREVIEW MODE IS ACTIVE.
    previewModifier As Double
    
End Type

'Calling functions can use this variable to access all FilterInfo for the current workingLayer copy.
Public curLayerValues As FilterInfo

'We may need a temporary copy of the selection mask for rendering purposes; if so, it will be stored here
Private tmpSelectionMask As pdLayer

'This function can be used to populate a valid SAFEARRAY2D structure against any layer
Public Sub prepSafeArray(ByRef srcSA As SAFEARRAY2D, ByRef srcLayer As pdLayer)
    
    'Populate a relevant SafeArray variable for the supplied layer
    With srcSA
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lBound = 0
        .Bounds(0).cElements = srcLayer.getLayerHeight
        .Bounds(1).lBound = 0
        .Bounds(1).cElements = srcLayer.getLayerArrayWidth
        .pvData = srcLayer.getLayerDIBits
    End With
    
End Sub


'prepImageData's job is to copy the relevant layer (or part of a layer, if a selection is active) into a temporary object,
' which individual filters and effects can then operate on.  prepImageData() also populates a relevant SafeArray object and
' a host of other variables, which filters and effects can copy locally to ensure the fastest possible runtime speed.
'
'In one of the better triumphs of PD's design, this function is used for both previews and actual filter applications.
' The isPreview parameter is used to notify the function of the intended purpose of a given call.  If isPreview is TRUE,
' the image will automatically be scaled to the size of the preview area, which allows the tool dialog to render much faster.
' Note that for thsi to work, an fxPreview control must be passed to the function.
'
'Finally, the calling routine can optionally specify a different progress bar maximum value.  By default, this is the current
' layer's width, but some routines run vertically and the progress bar needs to be changed accordingly.
Public Sub prepImageData(ByRef tmpSA As SAFEARRAY2D, Optional isPreview As Boolean = False, Optional previewTarget As fxPreviewCtl, Optional newProgBarMax As Long = -1)

    'Reset the public "cancel current action" tracker
    cancelCurrentAction = False

    'Prepare our temporary layer
    Set workingLayer = New pdLayer
        
    'If this is a preview, we need to calculate new width and height for the image that will appear in the preview window.
    Dim dstWidth As Double, dstHeight As Double
    Dim srcWidth As Double, srcHeight As Double
    Dim newWidth As Long, newHeight As Long
    
    'If this is not a preview, simply copy the current layer without modification
    If Not isPreview Then
    
        'Check for an active selection; if one is present, use that instead of the full layer.  Note that no special processing is
        ' applied to the selected area - a full rectangle is passed to the source function, with no accounting for non-rectangular
        ' boundaries or feathering.  All that work is handled *after* the processing is complete.
        If pdImages(g_CurrentImage).selectionActive Then
            workingLayer.createBlank pdImages(g_CurrentImage).mainSelection.boundWidth, pdImages(g_CurrentImage).mainSelection.boundHeight, pdImages(g_CurrentImage).getActiveLayer().getLayerColorDepth
            BitBlt workingLayer.getLayerDC, 0, 0, pdImages(g_CurrentImage).mainSelection.boundWidth, pdImages(g_CurrentImage).mainSelection.boundHeight, pdImages(g_CurrentImage).getActiveLayer().getLayerDC, pdImages(g_CurrentImage).mainSelection.boundLeft, pdImages(g_CurrentImage).mainSelection.boundTop, vbSrcCopy
        Else
            workingLayer.createFromExistingLayer pdImages(g_CurrentImage).getActiveLayer()
        End If
        
        'Premultiplied alpha is removed prior to processing; this allows various tools to return proper results.
        If workingLayer.getLayerColorDepth = 32 Then workingLayer.fixPremultipliedAlpha False
    
    'This IS a preview, meaning more work is involved.  We must prepare a unique copy of the image that matches the requested
    ' dimensions of the preview area (which are not assumed to be universal!).
    Else
    
        'Start by calculating the source area for the preview.  Generally this is the entire image, unless a selection is active;
        ' in that case, we only want to preview the selected area.  (I may change this behavior in the future.)
        If pdImages(g_CurrentImage).selectionActive Then
            srcWidth = pdImages(g_CurrentImage).mainSelection.boundWidth
            srcHeight = pdImages(g_CurrentImage).mainSelection.boundHeight
        Else
            srcWidth = pdImages(g_CurrentImage).getActiveLayer().getLayerWidth
            srcHeight = pdImages(g_CurrentImage).getActiveLayer().getLayerHeight
        End If
        
        'Destination width/height are generally the dimensions of the preview box, taking into account aspect ratio.  The only
        ' exception to this is when the image is actually smaller than the preview area - in that case use the whole image.
        dstWidth = previewTarget.getPreviewPic.ScaleWidth
        dstHeight = previewTarget.getPreviewPic.ScaleHeight
        
        If (srcWidth > dstWidth) Or (srcHeight > dstHeight) Then
            convertAspectRatio srcWidth, srcHeight, dstWidth, dstHeight, newWidth, newHeight
        Else
            newWidth = srcWidth
            newHeight = srcHeight
        End If
        
        'Next, we will create the temporary object (called "workingLayer") at the calculated preview dimensions.
        
        'Just like with a full image, if a selection is active, we only want to process the selected area.
        If pdImages(g_CurrentImage).selectionActive Then
        
            'Start by chopping out the full rectangular bounding area of the selection, and placing it inside a temporary object.
            ' This is done at the same color depth as the source image.
            Dim tmpLayer As pdLayer
            Set tmpLayer = New pdLayer
            tmpLayer.createBlank pdImages(g_CurrentImage).mainSelection.boundWidth, pdImages(g_CurrentImage).mainSelection.boundHeight, pdImages(g_CurrentImage).getActiveLayer().getLayerColorDepth
            BitBlt tmpLayer.getLayerDC, 0, 0, tmpLayer.getLayerWidth, tmpLayer.getLayerHeight, pdImages(g_CurrentImage).getActiveLayer.getLayerDC, pdImages(g_CurrentImage).mainSelection.boundLeft, pdImages(g_CurrentImage).mainSelection.boundTop, vbSrcCopy
            
            'If the source area is 32bpp, we want to remove premultiplication before doing any resizing
            'If tmpLayer.getLayerColorDepth = 32 Then tmpLayer.fixPremultipliedAlpha
            
            'We now want to shrink the selected area to the size of the preview box
            workingLayer.createFromExistingLayer tmpLayer, newWidth, newHeight
            
            'Release our temporary layer
            tmpLayer.eraseLayer
            Set tmpLayer = Nothing
        
        'If a selection is not currently active, this step is incredibly simple!
        Else
            workingLayer.createFromExistingLayer pdImages(g_CurrentImage).getActiveLayer(), newWidth, newHeight
        End If
        
        'Give the preview object a copy of this original, unmodified image data so it can show it to the user if requested
        If Not previewTarget.hasOriginalImage Then previewTarget.setOriginalImage workingLayer
        
        If workingLayer.getLayerColorDepth = 32 Then workingLayer.fixPremultipliedAlpha False
        
    End If
    
    'If a selection is active, make a backup of the selected area.  (We do this regardless of whether the current
    ' action is a preview or not.
    If pdImages(g_CurrentImage).selectionActive Then
        Set workingLayerBackup = New pdLayer
        workingLayerBackup.createFromExistingLayer workingLayer
    End If
    
    'With our temporary layer successfully created, populate the relevant SafeArray variable
    prepSafeArray tmpSA, workingLayer

    'Finally, populate the ubiquitous curLayerValues variable with everything a filter might want to know
    With curLayerValues
        .Left = 0
        .Top = 0
        .Right = workingLayer.getLayerWidth - 1
        .Bottom = workingLayer.getLayerHeight - 1
        .Width = workingLayer.getLayerWidth
        .Height = workingLayer.getLayerHeight
        .minX = 0
        .MinY = 0
        .maxX = workingLayer.getLayerWidth - 1
        .MaxY = workingLayer.getLayerHeight - 1
        .colorDepth = workingLayer.getLayerColorDepth
        .BytesPerPixel = (workingLayer.getLayerColorDepth \ 8)
        .LayerX = 0
        .LayerY = 0
        .previewModifier = workingLayer.getLayerWidth / pdImages(g_CurrentImage).getActiveLayer().getLayerWidth
    End With

    'Set up the progress bar (only if this is NOT a preview, mind you - during previews, the progress bar is not touched)
    If Not isPreview Then
        If newProgBarMax = -1 Then
            SetProgBarMax (curLayerValues.Left + curLayerValues.Width)
        Else
            SetProgBarMax newProgBarMax
        End If
    End If
    
    'If desired, the statement below can be used to verify that the function created a working layer at the proper dimensions
    'Debug.Print "prepImageData worked: " & workingLayer.getLayerHeight & ", " & workingLayer.getLayerWidth & " (" & workingLayer.getLayerArrayWidth & ")" & ", " & workingLayer.getLayerDIBits

End Sub


'The counterpart to prepImageData, finalizeImageData copies the working layer back into the source image, then renders
' everything to the screen.  Like prepImageData(), a preview target can also be named.  In this case, finalizeImageData
' will rely on the preview-related values calculated by prepImageData(), as it's presumed that preImageData will ALWAYS
' be called before this routine.
'
'Unlike prepImageData, this function has to do quite a bit of processing when selections are active.  The selection
' mask must be scanned for each pixel, and the results blended with the original image as appropriate.  For 32bpp images
' this is especially ugly.  (This is the price we pay for full selection feathering support!)
Public Sub finalizeImageData(Optional isPreview As Boolean = False, Optional previewTarget As fxPreviewCtl)

    'If the user canceled the current action, disregard the working layer and exit immediately.  The central processor
    ' will take care of additional clean-up.
    If (Not isPreview) And cancelCurrentAction Then
        
        workingLayer.eraseLayer
        Set workingLayer = Nothing
        
        Exit Sub
        
    End If
    
    'Prepare a few image arrays (and array headers) in advance.
    Dim wlImageData() As Byte
    Dim wlSA As SAFEARRAY2D
    
    Dim selImageData() As Byte
    Dim selSA As SAFEARRAY2D
    
    Dim x As Long, y As Long
    
    'Regardless of whether or not this is a preview, we process selections identically - by merging the newly modified
    ' workingLayer with its original version (as stored in workingLayerBackup), while accounting for any selection intricacies.
    If pdImages(g_CurrentImage).selectionActive Then
    
        'Before continuing further, create a copy of the selection mask at the relevant image size; note that "relevant size"
        ' is obviously calculated differently for previews.
        Dim selMaskCopy As pdLayer
        Set selMaskCopy = New pdLayer
        selMaskCopy.createBlank pdImages(g_CurrentImage).mainSelection.boundWidth, pdImages(g_CurrentImage).mainSelection.boundHeight
        BitBlt selMaskCopy.getLayerDC, 0, 0, selMaskCopy.getLayerWidth, selMaskCopy.getLayerHeight, pdImages(g_CurrentImage).mainSelection.selMask.getLayerDC, pdImages(g_CurrentImage).mainSelection.boundLeft, pdImages(g_CurrentImage).mainSelection.boundTop, vbSrcCopy
        
        'If this is a preview, resize the selection mask to match the preview size
        If isPreview Then
            Dim tmpLayer As pdLayer
            Set tmpLayer = New pdLayer
            tmpLayer.createFromExistingLayer selMaskCopy
            
            GDIPlusResizeLayer selMaskCopy, 0, 0, workingLayer.getLayerWidth, workingLayer.getLayerHeight, tmpLayer, 0, 0, tmpLayer.getLayerWidth, tmpLayer.getLayerHeight, InterpolationModeHighQualityBilinear
            
            tmpLayer.eraseLayer
            Set tmpLayer = Nothing
        End If
        
        'We now have a layer that represents the selection mask at the same offset and size as the workingLayer.  This allows
        ' us to process the selected area identically, regardless of whether this is a preview or a true full-layer operation.
        
        'A few rare functions actually change the color depth of the image.  Check for that now, and make sure the workingLayer
        ' and workingLayerBackup layers are the same bit-depth.
        If workingLayer.getLayerColorDepth <> workingLayerBackup.getLayerColorDepth Then
            If workingLayer.getLayerColorDepth = 24 Then
                workingLayerBackup.convertTo24bpp
            Else
                workingLayerBackup.convertTo32bpp
            End If
        End If
        
        'Next, point three arrays at three images: the original image, the newly modified image, and the selection mask copy
        ' we just created.
        prepSafeArray wlSA, workingLayer
        CopyMemory ByVal VarPtrArray(wlImageData()), VarPtr(wlSA), 4
        
        prepSafeArray selSA, selMaskCopy
        CopyMemory ByVal VarPtrArray(selImageData()), VarPtr(selSA), 4
        
        Dim dstImageData() As Byte
        Dim dstSA As SAFEARRAY2D
        prepSafeArray dstSA, workingLayerBackup
        CopyMemory ByVal VarPtrArray(dstImageData()), VarPtr(dstSA), 4
        
        Dim i As Long
        Dim thisAlpha As Long
        Dim blendAlpha As Double
        
        Dim dstQuickVal As Long
        dstQuickVal = pdImages(g_CurrentImage).getActiveLayer().getLayerColorDepth \ 8
            
        Dim workingLayerCD As Long
        workingLayerCD = workingLayer.getLayerColorDepth \ 8
        
        For x = 0 To workingLayer.getLayerWidth - 1
        For y = 0 To workingLayer.getLayerHeight - 1
            
            'Retrieve the selection mask value at this position.  Its value determines how this pixel is handled.
            thisAlpha = selImageData(x * 3, y)
            
            Select Case thisAlpha
                    
                'This pixel is not part of the selection, so completely ignore it
                Case 0
                
                'This pixel completely replaces the destination one, so simply copy it over
                Case 255
                    For i = 0 To dstQuickVal - 1
                        dstImageData(x * dstQuickVal + i, y) = wlImageData(x * workingLayerCD + i, y)
                    Next i
                        
                    'This pixel is antialiased or feathered, so it needs to be blended with the destination at the level specified
                    ' by the selection mask.
                    Case Else
                        blendAlpha = thisAlpha / 255
                        For i = 0 To dstQuickVal - 1
                            dstImageData(x * dstQuickVal + i, y) = BlendColors(dstImageData(x * dstQuickVal + i, y), wlImageData(x * workingLayerCD + i, y), blendAlpha)
                        Next i
                    
                End Select
                
            Next y
            Next x
            
            'With our work complete, point both ImageData() arrays away from their DIBs and deallocate them
            CopyMemory ByVal VarPtrArray(wlImageData), 0&, 4
            CopyMemory ByVal VarPtrArray(dstImageData), 0&, 4
            CopyMemory ByVal VarPtrArray(selImageData), 0&, 4
            
            Erase wlImageData
            Erase dstImageData
            Erase selImageData
    
    End If
        
    'Processing past this point is contingent on whether or not the current action is a preview.
        
    'If this is not a preview, simply copy the processed data back into the active layer
    If Not isPreview Then
        
        Message "Rendering image to screen..."
        
        'If a selection is active, copy the processed area into its proper place.
        If pdImages(g_CurrentImage).selectionActive Then
        
            If workingLayerBackup.getLayerColorDepth = 32 Then workingLayerBackup.fixPremultipliedAlpha True
            BitBlt pdImages(g_CurrentImage).getActiveLayer().getLayerDC, pdImages(g_CurrentImage).mainSelection.boundLeft, pdImages(g_CurrentImage).mainSelection.boundTop, pdImages(g_CurrentImage).mainSelection.boundWidth, pdImages(g_CurrentImage).mainSelection.boundHeight, workingLayerBackup.getLayerDC, 0, 0, vbSrcCopy
        
        'If a selection is not active, replace the entire layer with the contents of the working layer
        Else
            If workingLayer.getLayerColorDepth = 32 Then workingLayer.fixPremultipliedAlpha True
            pdImages(g_CurrentImage).getActiveLayer().createFromExistingLayer workingLayer
        End If
                
        'workingLayer and its backup have served their purposes, so erase them from memory
        Set workingLayer = Nothing
        Set workingLayerBackup = Nothing
                
        'If we're setting data to the screen, we can reasonably assume that the progress bar should be reset
        releaseProgressBar
        
        'Pass control to the viewport renderer, which will perform the actual rendering
        ScrollViewport pdImages(g_CurrentImage).containingForm
        
        Message "Finished."
    
    'If this is a preview, we need to repaint a preview box
    Else
        
        'If a selection is active, use the contents of workingLayerBackup instead of workingLayer to render the preview
        If pdImages(g_CurrentImage).selectionActive Then
            If workingLayerBackup.getLayerColorDepth = 32 Then workingLayerBackup.fixPremultipliedAlpha True
            previewTarget.setFXImage workingLayerBackup
        
        Else
            If workingLayer.getLayerColorDepth = 32 Then workingLayer.fixPremultipliedAlpha True
            previewTarget.setFXImage workingLayer
        
        End If
        
        'workingLayer and its backup have served their purposes, so erase them from memory
        Set workingLayer = Nothing
        Set workingLayerBackup = Nothing
        
    End If
    
End Sub

