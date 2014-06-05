Attribute VB_Name = "FastDrawing"
'***************************************************************************
'Fast API Graphics Routines Interface
'Copyright ©2001-2014 by Tanner Helland
'Created: 12/June/01
'Last updated: 05/June/14
'Last update: add support for individual filters and adjustments to override alpha premultiplication handling
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
'- In preview mode, workingDIB will contain a small, preview-size version of the image.
'- In non-preview mode, workingDIB will contain a copy of the full image.  We do not allow various effects and tools to operate
'   on the original image data, in case they cancel a function mid-way (we must be able to restore the original image).
'- If a selection is active, workingDIB will contain the selected part of the image, converted to 32bpp mode as necessary
'   (e.g. if feathering or antialiasing is enabled on the selection).
Public workingDIB As pdDIB

'When the workingDIB is first created, we store a backup of it in this DIB.  This backup is used to rebuild the full image
' while accounting for any selected areas; we merge the selected areas onto this original copy, then copy the composited result
' back onto the image.  This is easier than attempting to merge the area onto the main DIB while doing all our extra
' selection processing.
Private workingDIBBackup As pdDIB

'prepImageData is the function all PhotoDemon tools call when they need a copy of the image to operate on.  That function fills a
' variable of this type (FilterInfo) with everything a filter or effect could possibly want to know about the current DIB.
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
    
    'The colorDepth of the current DIB, specified as BITS per pixel; this will always be 24 or 32
    colorDepth As Long
    
    'BytesPerPixel is simply colorDepth / 8.  It is provided for convenience.
    BytesPerPixel As Long
    
    'Filters shouldn't have to worry about where the DIB is physically located, but when it comes
    ' time to set the image databack in place, knowing the layer's location on the primary image
    ' may be useful (as when previewing, for example)
    dibX As Long
    dibY As Long
    
    'When in preview mode, the on-screen image will typically be represented at a smaller-than-actual size.
    ' If an effect or filter operates on a radius (e.g. "blur radius 20"), that radius value has to be shrunk
    ' when working on the preview - otherwise, the preview effect will look much stronger than it actually is!
    ' This value can be multiplied by a radius or other value but ONLY WHEN PREVIEW MODE IS ACTIVE.
    previewModifier As Double
    
End Type

'Calling functions can use this variable to access all FilterInfo for the current workingDIB copy.
Public curDIBValues As FilterInfo

'We may need a temporary copy of the selection mask for rendering purposes; if so, it will be stored here
Private tmpSelectionMask As pdDIB

'This function can be used to populate a valid SAFEARRAY2D structure against any DIB
Public Sub prepSafeArray(ByRef srcSA As SAFEARRAY2D, ByRef srcDIB As pdDIB)
    
    'Populate a relevant SafeArray variable for the supplied DIB
    With srcSA
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lBound = 0
        .Bounds(0).cElements = srcDIB.getDIBHeight
        .Bounds(1).lBound = 0
        .Bounds(1).cElements = srcDIB.getDIBArrayWidth
        .pvData = srcDIB.getActualDIBBits
    End With
    
End Sub

'For some odd functions (e.g. export JPEG dialog), it's helpful to have the full power of prepImageData, but against
' a target other than the current image's main layer.  This function is roughly equivalent to prepImageData, below, but
' stripped down and specifically designed for PREVIEWS ONLY.  A source image must be explicitly supplied.
Public Sub previewNonStandardImage(ByRef tmpSA As SAFEARRAY2D, ByRef srcDIB As pdDIB, ByRef previewTarget As fxPreviewCtl)
    
    'Prepare our temporary DIB
    Set workingDIB = New pdDIB
        
    'We know this is a preview, so new width and height values need to be calculated against the size of the preview window.
    Dim dstWidth As Double, dstHeight As Double
    Dim srcWidth As Double, srcHeight As Double
    Dim newWidth As Long, newHeight As Long
    
    'Start by calculating the source area for the preview.  This changes based on several criteria:
    ' 1) Is the preview area set to "fit full image" or "100% zoom"?
    ' 2) Is a selection is active?  If so, we only want to preview the selected area.  (I may change this behavior in the future,
    '     so the user can actually see the fully composited result of any changes.)
    
    'The full image is being previewed.  Retrieve the entire thing, so we can shrink it down to size.
    If previewTarget.viewportFitFullImage Then
    
        srcWidth = srcDIB.getDIBWidth
        srcHeight = srcDIB.getDIBHeight
        
    'Only a section of the image is being preview (at 100% zoom).  Retrieve just that section.
    Else
    
        srcWidth = previewTarget.getPreviewWidth
        srcHeight = previewTarget.getPreviewHeight
        
        Dim curAspectRatio As Double
        
        If srcDIB.getDIBWidth < srcWidth Then
            srcWidth = srcDIB.getDIBWidth
        ElseIf srcDIB.getDIBHeight < srcHeight Then
            srcHeight = srcDIB.getDIBHeight
        End If
        
    End If
    
    'Destination width/height are generally the dimensions of the preview box, taking into account aspect ratio.  The only
    ' exception to this is when the image is actually smaller than the preview area - in that case use the whole image.
    dstWidth = previewTarget.getPreviewWidth
    dstHeight = previewTarget.getPreviewHeight
            
    If (srcWidth > dstWidth) Or (srcHeight > dstHeight) Then
        convertAspectRatio srcWidth, srcHeight, dstWidth, dstHeight, newWidth, newHeight
    Else
        newWidth = srcWidth
        newHeight = srcHeight
    End If
    
    'The area may be offset from the (0, 0) position if the user has elected to drag the preview area
    Dim hOffset As Long, vOffset As Long
    
    'Next, we will create the temporary object (called "workingDIB") at the calculated preview dimensions.  All editing
    ' actions are applied to this DIB.
    
    'The user is using "fit full image on-screen" mode for this preview.  Retrieve a tiny version of the image
    If previewTarget.viewportFitFullImage Then
        workingDIB.createFromExistingDIB srcDIB, newWidth, newHeight
        
    'The user is operating at 100% zoom.  Retrieve a subsection of the image, but do not scale it.
    Else
    
        'Calculate offsets, if any, for the image
        hOffset = previewTarget.offsetX
        vOffset = previewTarget.offsetY
        
        workingDIB.createBlank newWidth, newHeight, srcDIB.getDIBColorDepth
        BitBlt workingDIB.getDIBDC, 0, 0, dstWidth, dstHeight, srcDIB.getDIBDC, hOffset, vOffset, vbSrcCopy
        
    End If
    
    'Give the preview object a copy of this original, unmodified image data so it can show it to the user if requested
    If Not previewTarget.hasOriginalImage Then previewTarget.setOriginalImage workingDIB
    
    'For 32bpp layers, fix premultiplication now, as all effects assume UN-premultiplied alpha
    If workingDIB.getDIBColorDepth = 32 Then workingDIB.fixPremultipliedAlpha False
    
    'With our temporary DIB successfully created, populate the relevant SafeArray variable
    prepSafeArray tmpSA, workingDIB

    'Finally, populate the ubiquitous curDIBValues variable with everything a filter might want to know
    With curDIBValues
        .Left = 0
        .Top = 0
        .Right = workingDIB.getDIBWidth - 1
        .Bottom = workingDIB.getDIBHeight - 1
        .Width = workingDIB.getDIBWidth
        .Height = workingDIB.getDIBHeight
        .minX = 0
        .MinY = 0
        .maxX = workingDIB.getDIBWidth - 1
        .MaxY = workingDIB.getDIBHeight - 1
        .colorDepth = workingDIB.getDIBColorDepth
        .BytesPerPixel = (workingDIB.getDIBColorDepth \ 8)
        .dibX = 0
        .dibY = 0
        If previewTarget.viewportFitFullImage Then
            .previewModifier = workingDIB.getDIBWidth / srcDIB.getDIBWidth
        Else
            .previewModifier = 1#
        End If
    End With
    
    'If desired, the statement below can be used to verify that the function created a working DIB at the proper dimensions
    'Debug.Print "previewNonStandardImage worked: " & workingDIB.getDIBHeight & ", " & workingDIB.getDIBWidth & " (" & workingDIB.getDIBArrayWidth & ")" & ", " & workingDIB.getActualDIBBits

End Sub

'The counterpart to previewNonStandardImage, above
Public Sub finalizeNonstandardPreview(ByRef previewTarget As fxPreviewCtl)
    
    'Because is a preview, we only need to repaint a preview box
    
    'Fix premultiplied alpha if necessary
    If workingDIB.getDIBColorDepth = 32 Then workingDIB.fixPremultipliedAlpha True
    
    'Pass the modified image on to the specified preview control
    previewTarget.setFXImage workingDIB
    
    'workingDIB and its backup have served their purposes, so erase them from memory
    Set workingDIB = Nothing
    
End Sub


'prepImageData's job is to copy the relevant DIB (or part of a DIB, if a selection is active) into a temporary object,
' which individual filters and effects can then operate on.  prepImageData() also populates a relevant SafeArray object and
' a host of other variables, which filters and effects can copy locally to ensure the fastest possible runtime speed.
'
'In one of the better triumphs of PD's design, this function is used for both previews and actual filter applications.
' The isPreview parameter is used to notify the function of the intended purpose of a given call.  If isPreview is TRUE,
' the image will automatically be scaled to the size of the preview area, which allows the tool dialog to render much faster.
' Note that for this to work, an fxPreview control must be passed to the function.
'
'Finally, the calling routine can optionally specify a different maximum progress bar value.  By default, this is the current
' DIB's width, but some routines run vertically and the progress bar maximum needs to be changed accordingly.
Public Sub prepImageData(ByRef tmpSA As SAFEARRAY2D, Optional isPreview As Boolean = False, Optional previewTarget As fxPreviewCtl, Optional newProgBarMax As Long = -1, Optional ByVal doNotTouchProgressBar As Boolean = False, Optional ByVal doNotUnPremultiplyAlpha As Boolean = False)

    'Reset the public "cancel current action" tracker
    cancelCurrentAction = False

    'Prepare our temporary DIB
    Set workingDIB = New pdDIB
    
    'The new Layers design sometimes requires us to apply actions outside of a layer's actual boundary.
    ' (For example: a selected area that extends outside the boundary of the current image.)  When this
    ' happens, we have to do some extra handling to render a correct image; basically, we must null-pad
    ' the current layer DIB to the size of the image, then extract the relevant bits after the fact.
    Dim tmpLayer As pdLayer
    If pdImages(g_CurrentImage).selectionActive Then Set tmpLayer = New pdLayer
    
    'If this is a preview, we need to calculate new width and height for the image that will appear in the preview window.
    Dim dstWidth As Double, dstHeight As Double
    Dim srcWidth As Double, srcHeight As Double
    Dim newWidth As Long, newHeight As Long
    
    'If this is not a preview, simply copy the current DIB without modification
    If Not isPreview Then
    
        'Check for an active selection; if one is present, use that instead of the full DIB.  Note that no special processing is
        ' applied to the selected area - a full rectangle is passed to the source function, with no accounting for non-rectangular
        ' boundaries or feathering.  All that work is handled *after* the processing is complete.
        If pdImages(g_CurrentImage).selectionActive Then
            
            'Before proceeding further, null-pad the layer in question.  This will allow any possible selection to work,
            ' regardless of the layer's actual area.
            tmpLayer.CopyExistingLayer pdImages(g_CurrentImage).getActiveLayer
            tmpLayer.convertToNullPaddedLayer pdImages(g_CurrentImage).Width, pdImages(g_CurrentImage).Height
            
            'Now we can proceed to crop out the relevant parts of the layer from the selection boundary.
            workingDIB.createBlank pdImages(g_CurrentImage).mainSelection.boundWidth, pdImages(g_CurrentImage).mainSelection.boundHeight, pdImages(g_CurrentImage).getActiveDIB().getDIBColorDepth
            BitBlt workingDIB.getDIBDC, 0, 0, pdImages(g_CurrentImage).mainSelection.boundWidth, pdImages(g_CurrentImage).mainSelection.boundHeight, tmpLayer.layerDIB.getDIBDC, pdImages(g_CurrentImage).mainSelection.boundLeft, pdImages(g_CurrentImage).mainSelection.boundTop, vbSrcCopy
            
        Else
            workingDIB.createFromExistingDIB pdImages(g_CurrentImage).getActiveDIB()
        End If
        
        'Premultiplied alpha is removed prior to processing; this allows various tools to return proper results.
        ' Note that individual tools can override this behavior - this is helpful in certain cases, e.g. area filters like
        ' blur, where *not* premultiplying alpha causes the black RGB values from transparent areas to be "picked up"
        ' by the area handling.
        If (workingDIB.getDIBColorDepth = 32) And (Not doNotUnPremultiplyAlpha) Then workingDIB.fixPremultipliedAlpha False
    
    'This IS a preview, meaning more work is involved.  We must prepare a unique copy of the active layer that matches
    ' the requested dimensions of the preview area (which are not assumed to be universal), while accounting for the
    ' selection area!  Aaahhh!
    Else
    
        'Start by calculating the source area for the preview.  This changes based on several criteria:
        ' 1) Is the preview area set to "fit full image" or "100% zoom"?
        ' 2) Is a selection is active?  If so, we only want to preview the selected area.  (I may change this behavior in the future,
        '     so the user can actually see the fully composited result of any changes.)
        
        'The full image is being previewed.  Retrieve the entire thing.
        If previewTarget.viewportFitFullImage Then
        
            If pdImages(g_CurrentImage).selectionActive Then
                srcWidth = pdImages(g_CurrentImage).mainSelection.boundWidth
                srcHeight = pdImages(g_CurrentImage).mainSelection.boundHeight
            Else
                srcWidth = pdImages(g_CurrentImage).getActiveDIB().getDIBWidth
                srcHeight = pdImages(g_CurrentImage).getActiveDIB().getDIBHeight
            End If
        
        'Only a section of the image is being preview (at 100% zoom).  Retrieve just that section.
        Else
        
            srcWidth = previewTarget.getPreviewWidth
            srcHeight = previewTarget.getPreviewHeight
            
            Dim curAspectRatio As Double
            
            'If the preview area is larger than the image itself, just retrieve the full image.
            If pdImages(g_CurrentImage).selectionActive Then
            
                If (pdImages(g_CurrentImage).mainSelection.boundWidth < srcWidth) Then
                    srcWidth = pdImages(g_CurrentImage).mainSelection.boundWidth
                ElseIf (pdImages(g_CurrentImage).mainSelection.boundHeight < srcHeight) Then
                    srcHeight = pdImages(g_CurrentImage).mainSelection.boundHeight
                End If
                
            Else
            
                If pdImages(g_CurrentImage).getActiveDIB().getDIBWidth < srcWidth Then
                    srcWidth = pdImages(g_CurrentImage).getActiveDIB().getDIBWidth
                ElseIf pdImages(g_CurrentImage).getActiveDIB().getDIBHeight < srcHeight Then
                    srcHeight = pdImages(g_CurrentImage).getActiveDIB().getDIBHeight
                End If
                
            End If
            
        End If
        
        'Destination width/height are generally the dimensions of the preview box, taking into account aspect ratio.  The only
        ' exception to this is when the image is actually smaller than the preview area - in that case use the whole image.
        dstWidth = previewTarget.getPreviewWidth
        dstHeight = previewTarget.getPreviewHeight
                
        If (srcWidth > dstWidth) Or (srcHeight > dstHeight) Then
            convertAspectRatio srcWidth, srcHeight, dstWidth, dstHeight, newWidth, newHeight
        Else
            newWidth = srcWidth
            newHeight = srcHeight
        End If
        
        'The area may be offset from the (0, 0) position if the user has elected to drag the preview area
        Dim hOffset As Long, vOffset As Long
        
        'Next, we will create the temporary object (called "workingDIB") at the calculated preview dimensions.  All editing
        ' actions are applied to this DIB; if the user does not cancel the action, that DIB will be copied over the
        ' primary image.  If they cancel, we'll simply discard the temporary DIB.
        
        'Just like with a full image, if a selection is active, we only want to process the selected area.
        If pdImages(g_CurrentImage).selectionActive Then
        
            'Start by chopping out the full rectangular bounding area of the selection, and placing it inside a temporary object.
            ' This is done at the same color depth as the source image.  (Note that we do not do any preprocessing of the selection
            ' area at this juncture.  The full bounding rect of the selection is processed as-is, and it as at the *draw* step
            ' that we do any further processing.)
            Dim tmpDIB As pdDIB
            Set tmpDIB = New pdDIB
            tmpDIB.createBlank pdImages(g_CurrentImage).mainSelection.boundWidth, pdImages(g_CurrentImage).mainSelection.boundHeight, pdImages(g_CurrentImage).getActiveDIB().getDIBColorDepth
            
            'Before proceeding further, make a copy of the active layer, and null-pad it to the size of the parent image.
            ' This will allow any possible selection to work, regardless of a layer's actual area.
            tmpLayer.CopyExistingLayer pdImages(g_CurrentImage).getActiveLayer
            tmpLayer.convertToNullPaddedLayer pdImages(g_CurrentImage).Width, pdImages(g_CurrentImage).Height
            
            'NOW we can copy over the active layer's data, within the bounding box of the active selection
            BitBlt tmpDIB.getDIBDC, 0, 0, tmpDIB.getDIBWidth, tmpDIB.getDIBHeight, tmpLayer.layerDIB.getDIBDC, pdImages(g_CurrentImage).mainSelection.boundLeft, pdImages(g_CurrentImage).mainSelection.boundTop, vbSrcCopy
            
            'The user is using "fit full image on-screen" mode for this preview.  Retrieve a tiny version of the selection
            If previewTarget.viewportFitFullImage Then
                workingDIB.createFromExistingDIB tmpDIB, newWidth, newHeight
            
            'The user is operating at 100% zoom.  Retrieve a subsection of the selected area, but do not scale it.
            Else
            
                'Calculate offsets, if any, for the selected area
                hOffset = previewTarget.offsetX
                vOffset = previewTarget.offsetY
                
                workingDIB.createBlank newWidth, newHeight, pdImages(g_CurrentImage).getActiveDIB().getDIBColorDepth
                BitBlt workingDIB.getDIBDC, 0, 0, dstWidth, dstHeight, tmpDIB.getDIBDC, hOffset, vOffset, vbSrcCopy
            
            End If
            
            
            'Release our temporary DIB
            tmpDIB.eraseDIB
            Set tmpDIB = Nothing
        
        'If a selection is not currently active, this step is incredibly simple!
        Else
            
            'The user is using "fit full image on-screen" mode for this preview.  Retrieve a tiny version of the image
            If previewTarget.viewportFitFullImage Then
                workingDIB.createFromExistingDIB pdImages(g_CurrentImage).getActiveDIB(), newWidth, newHeight
                
            'The user is operating at 100% zoom.  Retrieve a subsection of the image, but do not scale it.
            Else
            
                'Calculate offsets, if any, for the image
                hOffset = previewTarget.offsetX
                vOffset = previewTarget.offsetY
                
                workingDIB.createBlank newWidth, newHeight, pdImages(g_CurrentImage).getActiveDIB().getDIBColorDepth
                BitBlt workingDIB.getDIBDC, 0, 0, dstWidth, dstHeight, pdImages(g_CurrentImage).getActiveDIB().getDIBDC, hOffset, vOffset, vbSrcCopy
                
            End If
            
        End If
        
        'Give the preview object a copy of this original, unmodified image data so it can show it to the user if requested
        If Not previewTarget.hasOriginalImage Then previewTarget.setOriginalImage workingDIB
        
        If (workingDIB.getDIBColorDepth = 32) And (Not doNotUnPremultiplyAlpha) Then workingDIB.fixPremultipliedAlpha False
        
    End If
    
    'If a selection is active, make a backup of the selected area.  (We do this regardless of whether the current
    ' action is a preview or not.
    If pdImages(g_CurrentImage).selectionActive Then
        Set workingDIBBackup = New pdDIB
        workingDIBBackup.createFromExistingDIB workingDIB
    End If
    
    'With our temporary DIB successfully created, populate the relevant SafeArray variable
    prepSafeArray tmpSA, workingDIB

    'Finally, populate the ubiquitous curDIBValues variable with everything a filter might want to know
    With curDIBValues
        .Left = 0
        .Top = 0
        .Right = workingDIB.getDIBWidth - 1
        .Bottom = workingDIB.getDIBHeight - 1
        .Width = workingDIB.getDIBWidth
        .Height = workingDIB.getDIBHeight
        .minX = 0
        .MinY = 0
        .maxX = workingDIB.getDIBWidth - 1
        .MaxY = workingDIB.getDIBHeight - 1
        .colorDepth = workingDIB.getDIBColorDepth
        .BytesPerPixel = (workingDIB.getDIBColorDepth \ 8)
        .dibX = 0
        .dibY = 0
        If isPreview Then
            If previewTarget.viewportFitFullImage Then
                .previewModifier = workingDIB.getDIBWidth / pdImages(g_CurrentImage).getActiveDIB().getDIBWidth
            Else
                .previewModifier = 1#
            End If
        Else
            .previewModifier = 1#
        End If
    End With

    'Set up the progress bar (only if this is NOT a preview, mind you - during previews, the progress bar is not touched)
    If (Not isPreview) And (Not doNotTouchProgressBar) Then
        If newProgBarMax = -1 Then
            SetProgBarMax (curDIBValues.Left + curDIBValues.Width)
        Else
            SetProgBarMax newProgBarMax
        End If
    End If
    
    'If desired, the statement below can be used to verify that the function created a working DIB at the proper dimensions
    'Debug.Print "prepImageData worked: " & workingDIB.getDIBHeight & ", " & workingDIB.getDIBWidth & " (" & workingDIB.getDIBArrayWidth & ")" & ", " & workingDIB.getActualDIBBits

End Sub


'The counterpart to prepImageData, finalizeImageData copies the working DIB back into the source image, then renders
' everything to the screen.  Like prepImageData(), a preview target can also be named.  In this case, finalizeImageData
' will rely on the preview-related values calculated by prepImageData(), as it's presumed that preImageData will ALWAYS
' be called before this routine.
'
'Unlike prepImageData, this function has to do quite a bit of processing when selections are active.  The selection
' mask must be scanned for each pixel, and the results blended with the original image as appropriate.  For 32bpp images
' this is especially ugly.  (This is the price we pay for full selection feathering support!)
Public Sub finalizeImageData(Optional isPreview As Boolean = False, Optional previewTarget As fxPreviewCtl, Optional ByVal alphaAlreadyPremultiplied As Boolean = False)

    'If the user canceled the current action, disregard the working DIB and exit immediately.  The central processor
    ' will take care of additional clean-up.
    If (Not isPreview) And cancelCurrentAction Then
        
        workingDIB.eraseDIB
        Set workingDIB = Nothing
        
        Exit Sub
        
    End If
    
    'Prepare a few image arrays (and array headers) in advance.
    Dim wlImageData() As Byte
    Dim wlSA As SAFEARRAY2D
    
    Dim selImageData() As Byte
    Dim selSA As SAFEARRAY2D
    
    Dim x As Long, y As Long
    
    'Regardless of whether or not this is a preview, we process selections identically - by merging the newly modified
    ' workingDIB with its original version (as stored in workingDIBBackup), while accounting for any selection intricacies.
    If pdImages(g_CurrentImage).selectionActive Then
    
        'Before continuing further, create a copy of the selection mask at the relevant image size; note that "relevant size"
        ' is obviously calculated differently for previews.
        Dim selMaskCopy As pdDIB
        Set selMaskCopy = New pdDIB
        selMaskCopy.createBlank pdImages(g_CurrentImage).mainSelection.boundWidth, pdImages(g_CurrentImage).mainSelection.boundHeight
        BitBlt selMaskCopy.getDIBDC, 0, 0, selMaskCopy.getDIBWidth, selMaskCopy.getDIBHeight, pdImages(g_CurrentImage).mainSelection.selMask.getDIBDC, pdImages(g_CurrentImage).mainSelection.boundLeft, pdImages(g_CurrentImage).mainSelection.boundTop, vbSrcCopy
        
        'If this is a preview, resize the selection mask to match the preview size
        If isPreview Then
            Dim tmpDIB As pdDIB
            Set tmpDIB = New pdDIB
            tmpDIB.createFromExistingDIB selMaskCopy
            
            'The preview is a shrunk version of the full image.  Shrink the selection mask to match.
            If previewTarget.viewportFitFullImage Then
                GDIPlusResizeDIB selMaskCopy, 0, 0, workingDIB.getDIBWidth, workingDIB.getDIBHeight, tmpDIB, 0, 0, tmpDIB.getDIBWidth, tmpDIB.getDIBHeight, InterpolationModeHighQualityBicubic
            
            'The preview is a 100% copy of the image.  Copy only the relevant part of the selection mask into the
            ' selection processing DIB.
            Else
                
                Dim hOffset As Long, vOffset As Long
                hOffset = previewTarget.offsetX
                vOffset = previewTarget.offsetY
                
                selMaskCopy.createBlank workingDIB.getDIBWidth, workingDIB.getDIBHeight
                BitBlt selMaskCopy.getDIBDC, 0, 0, selMaskCopy.getDIBWidth, selMaskCopy.getDIBHeight, pdImages(g_CurrentImage).mainSelection.selMask.getDIBDC, pdImages(g_CurrentImage).mainSelection.boundLeft + hOffset, pdImages(g_CurrentImage).mainSelection.boundTop + vOffset, vbSrcCopy
                
            End If
            
            tmpDIB.eraseDIB
            Set tmpDIB = Nothing
        End If
        
        'We now have a DIB that represents the selection mask at the same offset and size as the workingDIB.  This allows
        ' us to process the selected area identically, regardless of whether this is a preview or a true full-DIB operation.
        
        'A few rare functions actually change the color depth of the image.  Check for that now, and make sure the workingDIB
        ' and workingDIBBackup DIBs are the same bit-depth.
        If workingDIB.getDIBColorDepth <> workingDIBBackup.getDIBColorDepth Then
            If workingDIB.getDIBColorDepth = 24 Then
                workingDIBBackup.convertTo24bpp
            Else
                workingDIBBackup.convertTo32bpp
            End If
        End If
        
        'Before applying the selected area back onto the image, we need to null-pad the original layer.  (This is not done
        ' by prepImageData, because the user may elect to cancel a running action - and if they do that, we want to leave
        ' the original image untouched!  Thus, only the workingLayer has been null-padded.)
        pdImages(g_CurrentImage).getActiveLayer.convertToNullPaddedLayer pdImages(g_CurrentImage).Width, pdImages(g_CurrentImage).Height
        
        'Next, point three arrays at three images: the original image, the newly modified image, and the selection mask copy
        ' we just created.
        prepSafeArray wlSA, workingDIB
        CopyMemory ByVal VarPtrArray(wlImageData()), VarPtr(wlSA), 4
        
        prepSafeArray selSA, selMaskCopy
        CopyMemory ByVal VarPtrArray(selImageData()), VarPtr(selSA), 4
        
        Dim dstImageData() As Byte
        Dim dstSA As SAFEARRAY2D
        prepSafeArray dstSA, workingDIBBackup
        CopyMemory ByVal VarPtrArray(dstImageData()), VarPtr(dstSA), 4
        
        Dim i As Long
        Dim thisAlpha As Long
        Dim blendAlpha As Double
        
        Dim dstQuickVal As Long, srcQuickX As Long
        dstQuickVal = pdImages(g_CurrentImage).getActiveDIB().getDIBColorDepth \ 8
            
        Dim workingDIBCD As Long
        workingDIBCD = workingDIB.getDIBColorDepth \ 8
        
        For x = 0 To workingDIB.getDIBWidth - 1
            srcQuickX = x * 3
        For y = 0 To workingDIB.getDIBHeight - 1
            
            'Retrieve the selection mask value at this position.  Its value determines how this pixel is handled.
            thisAlpha = selImageData(srcQuickX, y)
            
            Select Case thisAlpha
                    
                'This pixel is not part of the selection, so completely ignore it
                Case 0
                
                'This pixel completely replaces the destination one, so simply copy it over
                Case 255
                    For i = 0 To dstQuickVal - 1
                        dstImageData(x * dstQuickVal + i, y) = wlImageData(x * workingDIBCD + i, y)
                    Next i
                        
                    'This pixel is antialiased or feathered, so it needs to be blended with the destination at the level specified
                    ' by the selection mask.
                    Case Else
                        blendAlpha = thisAlpha / 255
                        For i = 0 To dstQuickVal - 1
                            dstImageData(x * dstQuickVal + i, y) = BlendColors(dstImageData(x * dstQuickVal + i, y), wlImageData(x * workingDIBCD + i, y), blendAlpha)
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
        
    'If this is not a preview, simply copy the processed data back into the active DIB
    If Not isPreview Then
        
        Message "Rendering image to screen..."
        
        'If a selection is active, copy the processed area into its proper place.
        If pdImages(g_CurrentImage).selectionActive Then
        
            If (workingDIBBackup.getDIBColorDepth = 32) And (Not alphaAlreadyPremultiplied) Then workingDIBBackup.fixPremultipliedAlpha True
            BitBlt pdImages(g_CurrentImage).getActiveDIB().getDIBDC, pdImages(g_CurrentImage).mainSelection.boundLeft, pdImages(g_CurrentImage).mainSelection.boundTop, pdImages(g_CurrentImage).mainSelection.boundWidth, pdImages(g_CurrentImage).mainSelection.boundHeight, workingDIBBackup.getDIBDC, 0, 0, vbSrcCopy
            
            'Un-pad any null pixels we may have added as part of the selection interaction
            pdImages(g_CurrentImage).getActiveLayer.cropNullPaddedLayer
        
        'If a selection is not active, replace the entire DIB with the contents of the working DIB
        Else
            If (workingDIB.getDIBColorDepth = 32) And (Not alphaAlreadyPremultiplied) Then workingDIB.fixPremultipliedAlpha True
            pdImages(g_CurrentImage).getActiveDIB().createFromExistingDIB workingDIB
        End If
                
        'workingDIB and its backup have served their purposes, so erase them from memory
        Set workingDIB = Nothing
        Set workingDIBBackup = Nothing
                
        'If we're setting data to the screen, we can reasonably assume that the progress bar should be reset
        releaseProgressBar
        
        'Pass control to the viewport renderer, which will perform the actual rendering
        ScrollViewport pdImages(g_CurrentImage), FormMain.mainCanvas(0)
        
        Message "Finished."
    
    'If this is a preview, we need to repaint a preview box
    Else
        
        'If a selection is active, use the contents of workingDIBBackup instead of workingDIB to render the preview
        If pdImages(g_CurrentImage).selectionActive Then
            If (workingDIBBackup.getDIBColorDepth = 32) And (Not alphaAlreadyPremultiplied) Then workingDIBBackup.fixPremultipliedAlpha True
            previewTarget.setFXImage workingDIBBackup
        
        Else
            If (workingDIB.getDIBColorDepth = 32) And (Not alphaAlreadyPremultiplied) Then workingDIB.fixPremultipliedAlpha True
            previewTarget.setFXImage workingDIB
        
        End If
        
        'workingDIB and its backup have served their purposes, so erase them from memory
        Set workingDIB = Nothing
        Set workingDIBBackup = Nothing
        
    End If
    
End Sub

