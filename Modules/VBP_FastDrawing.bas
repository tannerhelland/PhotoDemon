Attribute VB_Name = "FastDrawing"
'***************************************************************************
'Fast API Graphics Routines Interface
'Copyright 2001-2017 by Tanner Helland
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
    minY As Long
    
    'The highest coordinate the filter is allowed to check.  This is almost always (width, height).
    maxX As Long
    maxY As Long
    
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

'In March 2015, I implemented unique preview identifiers.  This gives us a way to detect if a preview action is operating on the same image
' as the previous preview action.  Since a single tool dialog may generate thousands of previews (if the user is moving lots of sliders around),
' we can cache the preview image once, then simply copy it.  This is much faster than constantly regenerating the preview, especially if the
' source image is large or a complex selection is active.
Private m_PreviousPreviewID As Double

Private m_PreviousPreviewCopy As pdDIB

'When a preview control is unloaded, it can optionally call this to forcibly reset the preview engine's tracking ID.
' This will force a full refresh on the next preview (generally advised, in case the user switches between images).
Public Sub ResetPreviewIDs()
    m_PreviousPreviewID = 0#
End Sub

'This function can be used to populate a valid SAFEARRAY2D structure against any DIB
Public Sub PrepSafeArray(ByRef srcSA As SAFEARRAY2D, ByRef srcDIB As pdDIB)
    
    'Populate a relevant SafeArray variable for the supplied DIB
    With srcSA
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lBound = 0
        .Bounds(0).cElements = srcDIB.GetDIBHeight
        .Bounds(1).lBound = 0
        .Bounds(1).cElements = srcDIB.GetDIBStride
        .pvData = srcDIB.GetDIBPointer
    End With
    
End Sub

'This function can be used to populate a valid SAFEARRAY2D structure against any DIB, but instead of using bytes, each pixel
' is represented by a full LONG.
' DO NOT USE THIS ON 24-BPP DIBS, OBVIOUSLY!
Public Sub PrepSafeArray_Long(ByRef srcSA As SAFEARRAY2D, ByRef srcDIB As pdDIB)
    
    'Populate a relevant SafeArray variable for the supplied DIB
    With srcSA
        .cbElements = 4
        .cDims = 2
        .Bounds(0).lBound = 0
        .Bounds(0).cElements = srcDIB.GetDIBHeight
        .Bounds(1).lBound = 0
        .Bounds(1).cElements = srcDIB.GetDIBWidth
        .pvData = srcDIB.GetDIBPointer
    End With
    
End Sub

'For some odd functions (e.g. export JPEG dialog), it's helpful to have the full power of prepImageData, but against
' a target other than the current image's main layer.  This function is roughly equivalent to prepImageData, below, but
' stripped down and specifically designed for PREVIEWS ONLY.  A source image must be explicitly supplied.
Public Sub PreviewNonStandardImage(ByRef tmpSA As SAFEARRAY2D, ByRef srcDIB As pdDIB, ByRef previewTarget As pdFxPreviewCtl, Optional ByVal leaveAlphaPremultiplied As Boolean = False)
    
    'Before doing anything else, see if we can simply re-use our previous preview image
    If (m_PreviousPreviewID = previewTarget.GetUniqueID) And (m_PreviousPreviewID <> 0) And (Not workingDIB Is Nothing) And (Not m_PreviousPreviewCopy Is Nothing) Then
    
        'We know workingDIB and m_PreviousPreviewCopy are NOT nothing, thanks to the check above, so no DIB instantation is required.
        'Simply copy m_PreviousPreviewCopy into workingDIB
        workingDIB.CreateFromExistingDIB m_PreviousPreviewCopy
        
    'Something has changed, so we must regenerate our preview image from scratch.  (This is time-consuming and complicated, so we try
    ' to avoid it whenever possible.)
    Else
    
        If (workingDIB Is Nothing) Then Set workingDIB = New pdDIB
        
        'We know this is a preview, so new width and height values need to be calculated against the size of the preview window.
        Dim dstWidth As Double, dstHeight As Double
        Dim srcWidth As Double, srcHeight As Double
        Dim newWidth As Long, newHeight As Long
        
        'The full image is being previewed.  Retrieve the entire thing.
        If previewTarget.ViewportFitFullImage Then
            srcWidth = srcDIB.GetDIBWidth
            srcHeight = srcDIB.GetDIBHeight
            
        'Only a section of the image is being preview (at 100% zoom).  Retrieve just that section.
        Else
        
            srcWidth = previewTarget.GetPreviewWidth
            srcHeight = previewTarget.GetPreviewHeight
                
            If srcDIB.GetDIBWidth < srcWidth Then
                srcWidth = srcDIB.GetDIBWidth
                If srcDIB.GetDIBHeight < srcHeight Then srcHeight = srcDIB.GetDIBHeight
            ElseIf srcDIB.GetDIBHeight < srcHeight Then
                srcHeight = srcDIB.GetDIBHeight
            End If
            
        End If
        
        'Destination width/height are generally the dimensions of the preview box, taking into account aspect ratio.  The only
        ' exception to this is when the image is actually smaller than the preview area - in that case use the whole image.
        dstWidth = previewTarget.GetPreviewWidth
        dstHeight = previewTarget.GetPreviewHeight
                
        If (srcWidth > dstWidth) Or (srcHeight > dstHeight) Then
            ConvertAspectRatio srcWidth, srcHeight, dstWidth, dstHeight, newWidth, newHeight
        Else
            newWidth = srcWidth
            newHeight = srcHeight
        End If
        
        'The area may be offset from the (0, 0) position if the user has elected to drag the preview area
        Dim hOffset As Long, vOffset As Long
        
        'Next, we will create the temporary object (called "workingDIB") at the calculated preview dimensions.  All editing
        ' actions are applied to this DIB; if the user does not cancel the action, that DIB will be copied over the
        ' primary image.  If they cancel, we'll simply discard the temporary DIB.
            
        'The user is using "fit full image on-screen" mode for this preview.  Retrieve a tiny version of the image
        If previewTarget.ViewportFitFullImage Then
            workingDIB.CreateFromExistingDIB srcDIB, newWidth, newHeight, True
            
        'The user is operating at 100% zoom.  Retrieve a subsection of the image, but do not scale it.
        Else
        
            'Calculate offsets, if any, for the image
            hOffset = previewTarget.offsetX
            vOffset = previewTarget.offsetY
            
            If (workingDIB.GetDIBWidth <> newWidth) Or (workingDIB.GetDIBHeight <> newHeight) Or (workingDIB.GetDIBColorDepth <> srcDIB.GetDIBColorDepth) Then
                workingDIB.CreateBlank newWidth, newHeight, srcDIB.GetDIBColorDepth
            Else
                workingDIB.ResetDIB
            End If
            
            BitBlt workingDIB.GetDIBDC, 0, 0, newWidth, newHeight, srcDIB.GetDIBDC, hOffset, vOffset, vbSrcCopy
            workingDIB.SetInitialAlphaPremultiplicationState srcDIB.GetAlphaPremultiplication
            
        End If
        
        'Give the preview object a copy of this original, unmodified image data so it can show it to the user if requested
        If Not previewTarget.HasOriginalImage Then previewTarget.SetOriginalImage workingDIB
        
        'Make a note of this preview target's unique ID value.  We can use this in the future to avoid regenerating workingDIB
        ' from scratch.
        m_PreviousPreviewID = previewTarget.GetUniqueID
        
        'Also, make a backup copy of our completed workingDIB
        If (m_PreviousPreviewCopy Is Nothing) Then Set m_PreviousPreviewCopy = New pdDIB
        m_PreviousPreviewCopy.CreateFromExistingDIB workingDIB
        
    End If
    
    'We're also going to apply the requested alpha premultiplication in advance, which saves us some time on
    ' subsequent requests (assuming the caller always wants the same alpha state for a given filter).
    If (workingDIB.GetDIBColorDepth = 32) And (Not leaveAlphaPremultiplied) Then workingDIB.SetAlphaPremultiplication False
    
    'Finally, populate the ubiquitous curDIBValues variable with everything a filter might want to know
    With curDIBValues
        .Left = 0
        .Top = 0
        .Right = workingDIB.GetDIBWidth - 1
        .Bottom = workingDIB.GetDIBHeight - 1
        .Width = workingDIB.GetDIBWidth
        .Height = workingDIB.GetDIBHeight
        .minX = 0
        .minY = 0
        .maxX = workingDIB.GetDIBWidth - 1
        .maxY = workingDIB.GetDIBHeight - 1
        .colorDepth = workingDIB.GetDIBColorDepth
        .BytesPerPixel = (workingDIB.GetDIBColorDepth \ 8)
        .dibX = 0
        .dibY = 0
        If previewTarget.ViewportFitFullImage Then
            If (srcDIB.GetDIBWidth <> 0) Then
                .previewModifier = workingDIB.GetDIBWidth / srcDIB.GetDIBWidth
            Else
                .previewModifier = 1#
            End If
        Else
            .previewModifier = 1#
        End If
    End With
    
    'With our temporary DIB successfully created, populate the relevant SafeArray variable
    PrepSafeArray tmpSA, workingDIB
    
    'If desired, the statement below can be used to verify that the function created a working DIB at the proper dimensions
    'Debug.Print "previewNonStandardImage worked: " & workingDIB.getDIBHeight & ", " & workingDIB.getDIBWidth & " (" & workingDIB.GetDIBStride & ")" & ", " & workingDIB.GetDIBPointer

End Sub

'The counterpart to previewNonStandardImage, above
Public Sub FinalizeNonstandardPreview(ByRef previewTarget As pdFxPreviewCtl, Optional ByVal alphaAlreadyPremultiplied As Boolean = False)
    
    'Fix premultiplied alpha if necessary
    If (workingDIB.GetDIBColorDepth = 32) And (Not alphaAlreadyPremultiplied) Then
        workingDIB.SetAlphaPremultiplication True
    Else
        workingDIB.SetInitialAlphaPremultiplicationState True
    End If
    
    'Pass the modified image on to the specified preview control
    previewTarget.SetFXImage workingDIB
    
End Sub

'PrepImageData's job is to copy the relevant DIB (or part of a DIB, if a selection is active) into a temporary object,
' which individual filters and effects can then operate on.  prepImageData() also populates a relevant SafeArray object and
' a host of other variables, which filters and effects can copy locally to ensure the fastest possible runtime speed.
'
'In one of the better triumphs of PD's design, this function is used for both previews and actual filter applications.
' The isPreview parameter is used to notify the function of the intended purpose of a given call.  If isPreview is TRUE,
' the image will automatically be scaled to the size of the preview area, which allows the tool dialog to render much faster.
' Note that for this to work, an pdFxPreview control must be passed to the function.
'
'Finally, the calling routine can optionally specify a different maximum progress bar value.  By default, this is the current
' DIB's width, but some routines run vertically and the progress bar maximum needs to be changed accordingly.
Public Sub PrepImageData(ByRef tmpSA As SAFEARRAY2D, Optional isPreview As Boolean = False, Optional previewTarget As pdFxPreviewCtl, Optional newProgBarMax As Long = -1, Optional ByVal doNotTouchProgressBar As Boolean = False, Optional ByVal doNotUnPremultiplyAlpha As Boolean = False)

    'Reset the public "cancel current action" tracker
    g_cancelCurrentAction = False
    
    'The new Layers design sometimes requires us to apply actions outside of a layer's actual boundary.
    ' (For example: a selected area that extends outside the boundary of the current image.)  When this
    ' happens, we have to do some extra handling to render a correct image; basically, we must null-pad
    ' the current layer DIB to the size of the image, then extract the relevant bits after the fact.
    Dim tmpLayer As pdLayer, selBounds As RECTF
    If (pdImages(g_CurrentImage).selectionActive And pdImages(g_CurrentImage).mainSelection.IsLockedIn) Then
        Set tmpLayer = New pdLayer
        selBounds = pdImages(g_CurrentImage).mainSelection.GetBoundaryRect
    End If
            
    'If this is a preview, we need to calculate new width and height for the image that will appear in the preview window.
    Dim dstWidth As Double, dstHeight As Double
    Dim srcWidth As Double, srcHeight As Double
    Dim newWidth As Long, newHeight As Long
    
    'If this is not a preview, simply copy the current DIB without modification
    If (Not isPreview) Then
        
        'Prepare our temporary DIB
        If (workingDIB Is Nothing) Then Set workingDIB = New pdDIB
        
        'Check for an active selection; if one is present, use that instead of the full DIB.  Note that no special processing is
        ' applied to the selected area - a full rectangle is passed to the source function, with no accounting for non-rectangular
        ' boundaries or feathering.  All that work is handled *after* the processing is complete.
        If (pdImages(g_CurrentImage).selectionActive And pdImages(g_CurrentImage).mainSelection.IsLockedIn) Then
            
            'Before proceeding further, null-pad the layer in question.  This will allow any possible selection to work,
            ' regardless of the layer's actual area.
            tmpLayer.CopyExistingLayer pdImages(g_CurrentImage).GetActiveLayer
            tmpLayer.ConvertToNullPaddedLayer pdImages(g_CurrentImage).Width, pdImages(g_CurrentImage).Height
            
            'Now we can proceed to crop out the relevant parts of the layer from the selection boundary.
            workingDIB.CreateBlank selBounds.Width, selBounds.Height, pdImages(g_CurrentImage).GetActiveDIB().GetDIBColorDepth
            BitBlt workingDIB.GetDIBDC, 0, 0, selBounds.Width, selBounds.Height, tmpLayer.layerDIB.GetDIBDC, selBounds.Left, selBounds.Top, vbSrcCopy
            workingDIB.SetInitialAlphaPremultiplicationState pdImages(g_CurrentImage).GetActiveLayer.layerDIB.GetAlphaPremultiplication
            
        Else
            workingDIB.CreateFromExistingDIB pdImages(g_CurrentImage).GetActiveDIB()
        End If
        
    'This IS a preview, meaning more work is involved.  We must prepare a unique copy of the active layer that matches
    ' the requested dimensions of the preview area (which are not assumed to be universal), while accounting for the
    ' selection area!  Aaahhh!
    Else
        
        'Before doing anything else, see if we can simply re-use our previous preview image
        If (m_PreviousPreviewID = previewTarget.GetUniqueID) And (Not workingDIB Is Nothing) And (Not m_PreviousPreviewCopy Is Nothing) Then
        
            'We know workingDIB and m_PreviousPreviewCopy are NOT nothing, thanks to the check above, so no DIB instantation is required.
            
            'Simply copy m_PreviousPreviewCopy into workingDIB
            workingDIB.CreateFromExistingDIB m_PreviousPreviewCopy
            
        'Something has changed, so we must regenerate our preview image from scratch.  (This is time-consuming and complicated, so we try
        ' to avoid it whenever possible.)
        Else
        
            'Prepare our temporary DIB
            If (workingDIB Is Nothing) Then Set workingDIB = New pdDIB
            
            'Start by calculating the source area for the preview.  This changes based on several criteria:
            ' 1) Is the preview area set to "fit full image" or "100% zoom"?
            ' 2) Is a selection is active?  If so, we only want to preview the selected area.  (I may change this behavior in the future,
            '     so the user can actually see the fully composited result of any changes.)
            
            'The full image is being previewed.  Retrieve the entire thing.
            If previewTarget.ViewportFitFullImage Then
            
                If (pdImages(g_CurrentImage).selectionActive And pdImages(g_CurrentImage).mainSelection.IsLockedIn) Then
                    srcWidth = selBounds.Width
                    srcHeight = selBounds.Height
                Else
                    srcWidth = pdImages(g_CurrentImage).GetActiveDIB().GetDIBWidth
                    srcHeight = pdImages(g_CurrentImage).GetActiveDIB().GetDIBHeight
                End If
            
            'Only a section of the image is being preview (at 100% zoom).  Retrieve just that section.
            Else
            
                srcWidth = previewTarget.GetPreviewWidth
                srcHeight = previewTarget.GetPreviewHeight
                
                'If a selection is active, and the selected area is smaller than the preview window, constrain the source area
                ' to the selection boundaries.
                If (pdImages(g_CurrentImage).selectionActive And pdImages(g_CurrentImage).mainSelection.IsLockedIn) Then
                
                    If (selBounds.Width < srcWidth) Then
                        srcWidth = selBounds.Width
                        If (selBounds.Height < srcHeight) Then srcHeight = selBounds.Height
                    ElseIf (selBounds.Height < srcHeight) Then
                        srcHeight = selBounds.Height
                    End If
                    
                Else
                    
                    If (pdImages(g_CurrentImage).GetActiveDIB().GetDIBWidth < srcWidth) Then
                        srcWidth = pdImages(g_CurrentImage).GetActiveDIB().GetDIBWidth
                        If (pdImages(g_CurrentImage).GetActiveDIB().GetDIBHeight < srcHeight) Then srcHeight = pdImages(g_CurrentImage).GetActiveDIB().GetDIBHeight
                    ElseIf (pdImages(g_CurrentImage).GetActiveDIB().GetDIBHeight < srcHeight) Then
                        srcHeight = pdImages(g_CurrentImage).GetActiveDIB().GetDIBHeight
                    End If
                    
                End If
                
            End If
            
            'Destination width/height are generally the dimensions of the preview box, taking into account aspect ratio.  The only
            ' exception to this is when the image is actually smaller than the preview area - in that case use the whole image.
            dstWidth = previewTarget.GetPreviewWidth
            dstHeight = previewTarget.GetPreviewHeight
                    
            If (srcWidth > dstWidth) Or (srcHeight > dstHeight) Then
                ConvertAspectRatio srcWidth, srcHeight, dstWidth, dstHeight, newWidth, newHeight
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
            If pdImages(g_CurrentImage).selectionActive And pdImages(g_CurrentImage).mainSelection.IsLockedIn Then
            
                'Start by chopping out the full rectangular bounding area of the selection, and placing it inside a temporary object.
                ' This is done at the same color depth as the source image.  (Note that we do not do any preprocessing of the selection
                ' area at this juncture.  The full bounding rect of the selection is processed as-is, and it as at the *draw* step
                ' that we do any further processing.)
                Dim tmpDIB As pdDIB
                Set tmpDIB = New pdDIB
                tmpDIB.CreateBlank selBounds.Width, selBounds.Height, pdImages(g_CurrentImage).GetActiveDIB().GetDIBColorDepth
                
                'Before proceeding further, make a copy of the active layer, and null-pad it to the size of the parent image.
                ' This will allow any possible selection to work, regardless of a layer's actual area.
                tmpLayer.CopyExistingLayer pdImages(g_CurrentImage).GetActiveLayer
                tmpLayer.ConvertToNullPaddedLayer pdImages(g_CurrentImage).Width, pdImages(g_CurrentImage).Height
                
                'NOW we can copy over the active layer's data, within the bounding box of the active selection
                BitBlt tmpDIB.GetDIBDC, 0, 0, tmpDIB.GetDIBWidth, tmpDIB.GetDIBHeight, tmpLayer.layerDIB.GetDIBDC, selBounds.Left, selBounds.Top, vbSrcCopy
                
                'The user is using "fit full image on-screen" mode for this preview.  Retrieve a tiny version of the selection
                If previewTarget.ViewportFitFullImage Then
                    workingDIB.CreateFromExistingDIB tmpDIB, newWidth, newHeight
                
                'The user is operating at 100% zoom.  Retrieve a subsection of the selected area, but do not scale it.
                Else
                
                    'Calculate offsets, if any, for the selected area
                    hOffset = previewTarget.offsetX
                    vOffset = previewTarget.offsetY
                    
                    If ((workingDIB.GetDIBWidth <> newWidth) Or (workingDIB.GetDIBHeight <> newHeight)) Then
                        workingDIB.CreateBlank newWidth, newHeight, pdImages(g_CurrentImage).GetActiveDIB().GetDIBColorDepth
                    Else
                        workingDIB.ResetDIB
                    End If
                    
                    BitBlt workingDIB.GetDIBDC, 0, 0, dstWidth, dstHeight, tmpDIB.GetDIBDC, hOffset, vOffset, vbSrcCopy
                    workingDIB.SetInitialAlphaPremultiplicationState pdImages(g_CurrentImage).GetActiveDIB().GetAlphaPremultiplication
                
                End If
                
                'Release our temporary DIB
                tmpDIB.EraseDIB
                Set tmpDIB = Nothing
            
            'If a selection is not currently active, this step is incredibly simple!
            Else
                
                'The user is using "fit full image on-screen" mode for this preview.  Retrieve a tiny version of the image
                If previewTarget.ViewportFitFullImage Then
                    workingDIB.CreateFromExistingDIB pdImages(g_CurrentImage).GetActiveDIB(), newWidth, newHeight, True
                    
                'The user is operating at 100% zoom.  Retrieve a subsection of the image, but do not scale it.
                Else
                
                    'Calculate offsets, if any, for the image
                    hOffset = previewTarget.offsetX
                    vOffset = previewTarget.offsetY
                    
                    If ((workingDIB.GetDIBWidth <> newWidth) Or (workingDIB.GetDIBHeight <> newHeight)) Then
                        workingDIB.CreateBlank newWidth, newHeight, pdImages(g_CurrentImage).GetActiveDIB().GetDIBColorDepth
                    Else
                        workingDIB.ResetDIB
                    End If
                    
                    BitBlt workingDIB.GetDIBDC, 0, 0, dstWidth, dstHeight, pdImages(g_CurrentImage).GetActiveDIB().GetDIBDC, hOffset, vOffset, vbSrcCopy
                    workingDIB.SetInitialAlphaPremultiplicationState pdImages(g_CurrentImage).GetActiveDIB().GetAlphaPremultiplication
                    
                End If
                
            End If
            
            'Give the preview object a copy of this original, unmodified image data so it can show it to the user if requested
            If (Not previewTarget.HasOriginalImage) Then previewTarget.SetOriginalImage workingDIB
            
            'We're also going to apply the requested alpha premultiplication in advance, which saves us some time on
            ' subsequent requests (assuming the caller always wants the same alpha state for a given filter).
            If (workingDIB.GetDIBColorDepth = 32) And (workingDIB.GetAlphaPremultiplication <> doNotUnPremultiplyAlpha) Then
                workingDIB.SetAlphaPremultiplication doNotUnPremultiplyAlpha
            End If
            
            'Make a note of this preview target's unique ID value.  We can use this in the future to avoid regenerating workingDIB
            ' from scratch.
            m_PreviousPreviewID = previewTarget.GetUniqueID
            
            'Also, make a backup copy of our completed workingDIB
            If (m_PreviousPreviewCopy Is Nothing) Then Set m_PreviousPreviewCopy = New pdDIB
            m_PreviousPreviewCopy.CreateFromExistingDIB workingDIB
            
        'End "preview copy is valid" vs "preview must be regenerated from scratch" handling
        End If
        
    'End non-preview vs preview mode handling
    End If
    
    'Premultiplied alpha is typically removed prior to processing; this allows various tools to return proper results.
    ' Note that individual tools can override this behavior - this is helpful in certain cases, e.g. area filters like
    ' blur, where *not* premultiplying alpha causes the black RGB values from transparent areas to be "picked up"
    ' by the area handling.
    If (workingDIB.GetDIBColorDepth = 32) And (workingDIB.GetAlphaPremultiplication <> doNotUnPremultiplyAlpha) Then workingDIB.SetAlphaPremultiplication doNotUnPremultiplyAlpha
    
    'If a selection is active, make a backup of the selected area.  (We do this regardless of whether the current
    ' action is a preview or not.)
    If pdImages(g_CurrentImage).selectionActive And pdImages(g_CurrentImage).mainSelection.IsLockedIn Then
        If (workingDIBBackup Is Nothing) Then Set workingDIBBackup = New pdDIB
        workingDIBBackup.CreateFromExistingDIB workingDIB
    End If
    
    'Finally, populate the ubiquitous curDIBValues variable with everything a filter might want to know
    With curDIBValues
        .Left = 0
        .Top = 0
        .Right = workingDIB.GetDIBWidth - 1
        .Bottom = workingDIB.GetDIBHeight - 1
        .Width = workingDIB.GetDIBWidth
        .Height = workingDIB.GetDIBHeight
        .minX = 0
        .minY = 0
        .maxX = workingDIB.GetDIBWidth - 1
        .maxY = workingDIB.GetDIBHeight - 1
        .colorDepth = workingDIB.GetDIBColorDepth
        .BytesPerPixel = (workingDIB.GetDIBColorDepth \ 8)
        .dibX = 0
        .dibY = 0
        If isPreview Then
            If previewTarget.ViewportFitFullImage Then
                .previewModifier = workingDIB.GetDIBWidth / pdImages(g_CurrentImage).GetActiveDIB().GetDIBWidth
            Else
                .previewModifier = 1#
            End If
        Else
            .previewModifier = 1#
        End If
    End With
    
    'With our temporary DIB successfully created, populate the relevant SafeArray variable, so the caller has direct access to the DIB
    PrepSafeArray tmpSA, workingDIB
    
    'Set up the progress bar (only if this is NOT a preview, mind you - during previews, the progress bar is not touched)
    If (Not isPreview) And (Not doNotTouchProgressBar) Then
        If newProgBarMax = -1 Then
            SetProgBarMax (curDIBValues.Left + curDIBValues.Width)
        Else
            SetProgBarMax newProgBarMax
        End If
    End If
    
    'If desired, the statement below can be used to verify that the function created a working DIB at the proper dimensions
    'Debug.Print "prepImageData worked: " & workingDIB.getDIBHeight & ", " & workingDIB.getDIBWidth & " (" & workingDIB.GetDIBStride & ")" & ", " & workingDIB.GetDIBPointer

End Sub

'The counterpart to prepImageData, finalizeImageData copies the working DIB back into the source image, then renders
' everything to the screen.  Like prepImageData(), a preview target can also be named.  In this case, finalizeImageData
' will rely on the preview-related values calculated by prepImageData(), as it's presumed that preImageData will ALWAYS
' be called before this routine.
'
'Unlike prepImageData, this function has to do quite a bit of processing when selections are active.  The selection
' mask must be scanned for each pixel, and the results blended with the original image as appropriate.  For 32bpp images
' this is especially ugly.  (This is the price we pay for full selection feathering support!)
Public Sub FinalizeImageData(Optional isPreview As Boolean = False, Optional previewTarget As pdFxPreviewCtl, Optional ByVal alphaAlreadyPremultiplied As Boolean = False)

    'If the user canceled the current action, disregard the working DIB and exit immediately.  The central processor
    ' will take care of additional clean-up.
    If (Not isPreview) And g_cancelCurrentAction Then
        workingDIB.EraseDIB
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
    If pdImages(g_CurrentImage).selectionActive And pdImages(g_CurrentImage).mainSelection.IsLockedIn Then
        
        'Retrieve the current selection boundaries
        Dim selBounds As RECTF
        selBounds = pdImages(g_CurrentImage).mainSelection.GetBoundaryRect
        
        'Before continuing further, create a copy of the selection mask at the relevant image size; note that "relevant size"
        ' is obviously calculated differently for previews.
        Dim selMaskCopy As pdDIB
        Set selMaskCopy = New pdDIB
        selMaskCopy.CreateBlank selBounds.Width, selBounds.Height
        BitBlt selMaskCopy.GetDIBDC, 0, 0, selMaskCopy.GetDIBWidth, selMaskCopy.GetDIBHeight, pdImages(g_CurrentImage).mainSelection.GetMaskDC(), selBounds.Left, selBounds.Top, vbSrcCopy
        
        'If this is a preview, resize the selection mask to match the preview size
        If isPreview Then
            Dim tmpDIB As pdDIB
            Set tmpDIB = New pdDIB
            tmpDIB.CreateFromExistingDIB selMaskCopy
            
            'The preview is a shrunk version of the full image.  Shrink the selection mask to match.
            If previewTarget.ViewportFitFullImage Then
                GDIPlusResizeDIB selMaskCopy, 0, 0, workingDIB.GetDIBWidth, workingDIB.GetDIBHeight, tmpDIB, 0, 0, tmpDIB.GetDIBWidth, tmpDIB.GetDIBHeight, GP_IM_HighQualityBicubic
            
            'The preview is a 100% copy of the image.  Copy only the relevant part of the selection mask into the
            ' selection processing DIB.
            Else
                
                Dim hOffset As Long, vOffset As Long
                hOffset = previewTarget.offsetX
                vOffset = previewTarget.offsetY
                
                selMaskCopy.CreateBlank workingDIB.GetDIBWidth, workingDIB.GetDIBHeight
                BitBlt selMaskCopy.GetDIBDC, 0, 0, selMaskCopy.GetDIBWidth, selMaskCopy.GetDIBHeight, pdImages(g_CurrentImage).mainSelection.GetMaskDC(), selBounds.Left + hOffset, selBounds.Top + vOffset, vbSrcCopy
                
            End If
            
            tmpDIB.EraseDIB
            Set tmpDIB = Nothing
        End If
        
        'We now have a DIB that represents the selection mask at the same offset and size as the workingDIB.  This allows
        ' us to process the selected area identically, regardless of whether this is a preview or a true full-DIB operation.
        
        'A few rare functions actually change the color depth of the image.  Check for that now, and make sure the workingDIB
        ' and workingDIBBackup DIBs are the same bit-depth.
        If (workingDIB.GetDIBColorDepth <> workingDIBBackup.GetDIBColorDepth) Then
            If workingDIB.GetDIBColorDepth = 24 Then
                workingDIBBackup.ConvertTo24bpp
            Else
                workingDIBBackup.ConvertTo32bpp
            End If
        End If
        
        'Before applying the selected area back onto the image, we need to null-pad the original layer.  (This is not done
        ' by prepImageData, because the user may elect to cancel a running action - and if they do that, we want to leave
        ' the original image untouched!  Thus, only the workingLayer has been null-padded.)
        pdImages(g_CurrentImage).GetActiveLayer.ConvertToNullPaddedLayer pdImages(g_CurrentImage).Width, pdImages(g_CurrentImage).Height
        
        'Next, point three arrays at three images: the original image, the newly modified image, and the selection mask copy
        ' we just created.
        PrepSafeArray wlSA, workingDIB
        CopyMemory ByVal VarPtrArray(wlImageData()), VarPtr(wlSA), 4
        
        PrepSafeArray selSA, selMaskCopy
        CopyMemory ByVal VarPtrArray(selImageData()), VarPtr(selSA), 4
        
        Dim dstImageData() As Byte
        Dim dstSA As SAFEARRAY2D
        PrepSafeArray dstSA, workingDIBBackup
        CopyMemory ByVal VarPtrArray(dstImageData()), VarPtr(dstSA), 4
        
        Dim i As Long, thisAlpha As Long, blendAlpha As Double
        
        'TODO: figure out why the selection mask is always being copied at 24-bpp
        Dim selMaskDepth As Long
        
        Dim dstQuickVal As Long, srcQuickX As Long
        dstQuickVal = pdImages(g_CurrentImage).GetActiveDIB().GetDIBColorDepth \ 8
            
        Dim workingDIBCD As Long
        workingDIBCD = workingDIB.GetDIBColorDepth \ 8
        
        For x = 0 To workingDIB.GetDIBWidth - 1
            srcQuickX = x * 3
        For y = 0 To workingDIB.GetDIBHeight - 1
            
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
    If (Not isPreview) Then
        
        'If a selection is active, copy the processed area into its proper place.
        If pdImages(g_CurrentImage).selectionActive And pdImages(g_CurrentImage).mainSelection.IsLockedIn Then
        
            If (workingDIBBackup.GetDIBColorDepth = 32) And (Not alphaAlreadyPremultiplied) Then workingDIBBackup.SetAlphaPremultiplication True
            BitBlt pdImages(g_CurrentImage).GetActiveDIB().GetDIBDC, selBounds.Left, selBounds.Top, selBounds.Width, selBounds.Height, workingDIBBackup.GetDIBDC, 0, 0, vbSrcCopy
            
            'Un-pad any null pixels we may have added as part of the selection interaction
            pdImages(g_CurrentImage).GetActiveLayer.CropNullPaddedLayer
        
        'If a selection is not active, replace the entire DIB with the contents of the working DIB
        Else
            
            If (workingDIB.GetDIBColorDepth = 32) And (Not alphaAlreadyPremultiplied) Then
                workingDIB.SetAlphaPremultiplication True
            Else
                workingDIB.SetInitialAlphaPremultiplicationState True
            End If
            
            pdImages(g_CurrentImage).GetActiveDIB().CreateFromExistingDIB workingDIB
            
        End If
        
        'workingDIB and its backup have served their purposes, so erase them from memory
        Set workingDIB = Nothing
        Set workingDIBBackup = Nothing
        
        'If we're setting data to the screen, we can reasonably assume that the progress bar should be reset
        ReleaseProgressBar
        
        'Notify the parent of the target layer of our changes
        pdImages(g_CurrentImage).NotifyImageChanged UNDO_LAYER, pdImages(g_CurrentImage).GetActiveLayerIndex
        
        'Pass control to the viewport renderer, which will perform the actual rendering
        Viewport_Engine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), FormMain.mainCanvas(0)
        
        Message "Finished."
    
    'If this is a preview, we need to repaint a preview box
    Else
        
        'If a selection is active, use the contents of workingDIBBackup instead of workingDIB to render the preview
        If pdImages(g_CurrentImage).selectionActive And pdImages(g_CurrentImage).mainSelection.IsLockedIn Then
            
            If (workingDIBBackup.GetDIBColorDepth = 32) And (Not alphaAlreadyPremultiplied) Then
                workingDIBBackup.SetAlphaPremultiplication True
            Else
                workingDIBBackup.SetInitialAlphaPremultiplicationState True
            End If
            
            previewTarget.SetFXImage workingDIBBackup
        
        Else
            
            'Prior to premultiplying alpha, apply color management.  (It is more efficient to do it here, prior to premultiplying
            ' alpha values, then to let the preview control handle it manually - as it must undo premultiplication to calculate
            ' a proper result.)
            Dim weCanHandleCM As Boolean: weCanHandleCM = False
            
            If (workingDIB.GetDIBColorDepth = 32) And (Not alphaAlreadyPremultiplied) Then
                weCanHandleCM = True
                ColorManagement.ApplyDisplayColorManagement workingDIB
                workingDIB.SetAlphaPremultiplication True
            Else
                workingDIB.SetInitialAlphaPremultiplicationState True
            End If
            
            previewTarget.SetFXImage workingDIB, weCanHandleCM
        
        End If
        
    End If
    
End Sub

