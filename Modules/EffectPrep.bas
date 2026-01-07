Attribute VB_Name = "EffectPrep"
'***************************************************************************
'Fast API Graphics Routines Interface
'Copyright 2001-2026 by Tanner Helland
'Created: 12/June/01
'Last updated: 25/July/17
'Last update: greatly optimize effects when an active selection is present
'
'This interface provides API support for the main image interaction routines. It assigns memory data
' into a useable array, and later transfers that array back into memory.  Very fast, very compact, can't
' live without it. These functions are arguably the most integral part of PhotoDemon.
'
'If you want to know more about how DIB sections work - and why they're so fast compared to VB's internal
' .PSet and .Point methods - please visit https://tannerhelland.com/2008/06/18/vb-graphics-programming-3.html
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Any time a tool dialog is in use, the image to be operated on will be stored IN THIS LAYER.
'- In preview mode, workingDIB contains a small, preview-size version of the image.
'- In non-preview mode, workingDIB contains a full-sized copy of the active layer.  PhotoDemon doesn't allow effects
'   and tools to operate on original image data; this is why a user can cancel functions mid-process.
'- If a selection is active, workingDIB contains only the selected portion of the image; unselected regions will be
'   auto-masked with transparency, so individual functions don't need to concern themselves with those details.
Public workingDIB As pdDIB

'When a workingDIB instance is first created, we store a local backup of it.  The backup is used to rebuild the
' full image while accounting for selected regions; we can simply merge selected pixels onto the original copy,
' then copy the composited result back onto the image.  This is easier (anb faster) than attempting to merge the
' area onto the main DIB while doing extra selection processing.
Private workingDIBBackup As pdDIB

'EffectPrep.PrepImageData() is what PhotoDemon adjustments and effects call to request a copy of the current image.
' That function fills a variable of this type (FilterInfo) with everything the effect could possibly want to know
' about the active DIB.
Public Type FilterInfo
    
    'Lowest coordinate the filter is allowed to *write*.  At present, this is (almost?) always (0, 0).
    Left As Long
    Top As Long
    
    'Highest coordinate the filter is allowed to *write*.  Changes past this position will be ignored.
    Right As Long
    Bottom As Long
    
    'Dimensions of the allowable *write* rectangle.  Provided for convenience, only - this value will
    ' always match the Left/Top/Right/Bottom coordinates provided above.
    Width As Long
    Height As Long
    
    'Lowest coordinate the filter is allowed to *read*.  At present, this is always (0, 0).
    minX As Long
    minY As Long
    
    'Highest coordinate the filter is allowed to *read*.  At present, this is always (width, height).
    maxX As Long
    maxY As Long
    
    'The colorDepth of the current DIB, specified as BITS per pixel; at present, this is always 32.
    colorDepth As Long
    
    'bytesPerPixel is simply colorDepth / 8.  It is provided for convenience, to help callers calculate stride.
    bytesPerPixel As Long
    
    'When in preview mode, the on-screen image is typically shrunk to some smaller-than-actual size.  If an
    ' effect or filter operates on a radius (e.g. "blur radius 20"), the previewed radius value must be shrunk
    ' - otherwise, the preview effect will look much stronger than it actually is!  This value is the ratio
    ' between the original image size and it's current size; it can be multiplied by a radius or other value but
    ' ONLY WHEN PREVIEW MODE IS ACTIVE.  Ignore it during non-preview events.
    previewModifier As Double
    
End Type

'Calling functions can use this variable to access all FilterInfo for the current workingDIB copy.
Public curDIBValues As FilterInfo

'In March 2015, I implemented unique preview identifiers.  This gives PD a way to detect when preview operations
' target the same image (or image region) as a previous preview action.  Because a single tool dialog may generate
' thousands of previews (if the user is moving lots of sliders around), PD will attempt to cache a valid preview
' image once, then simply copy it on subsequent calls.  This is much faster than constantly regenerating the preview,
' especially if the source image is large or a selection is active.
Private m_PreviousPreviewID As Double, m_PreviousPreviewCopy As pdDIB, m_PreviewWasRegenerated As Boolean
Private m_SelectionMaskBackup As pdDIB

'When a preview control is unloaded, it can optionally call this to forcibly reset the preview engine's tracking ID.
' This forces a full refresh on the next preview (which is always advised, in case the user switches between images).
Public Sub ResetPreviewIDs()
    m_PreviousPreviewID = 0#
End Sub

'This function can be used to populate a SafeArray2D structure against any arbitrary DIB.
Public Sub PrepSafeArray(ByRef srcSA As SafeArray2D, ByRef srcDIB As pdDIB)
    
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

'For some odd functions (e.g. export JPEG dialog), it's helpful to have the full power of prepImageData,
' but to apply it against something other than the current image's active layer.  This function is roughly
' equivalent to prepImageData, below, but stripped down and specifically designed for PREVIEWS ONLY.
' A source DIB must be explicitly supplied.
Public Sub PreviewNonStandardImage(ByRef tmpSA As SafeArray2D, ByRef srcDIB As pdDIB, ByRef previewTarget As pdFxPreviewCtl, Optional ByVal leaveAlphaPremultiplied As Boolean = False)
    
    'Before doing anything else, see if we can simply re-use our previous preview image
    If (m_PreviousPreviewID = previewTarget.GetUniqueID) And (m_PreviousPreviewID <> 0) And (Not workingDIB Is Nothing) And (Not m_PreviousPreviewCopy Is Nothing) Then
    
        'We know workingDIB and m_PreviousPreviewCopy are NOT nothing, thanks to the check above, so no DIB instantation is required.
        'Simply copy m_PreviousPreviewCopy into workingDIB
        workingDIB.CreateFromExistingDIB m_PreviousPreviewCopy
        
    'Something has changed, so we must regenerate our preview image from scratch.  (This is time-consuming and complicated,
    ' so we try to avoid it whenever possible.)
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
                
            If (srcDIB.GetDIBWidth < srcWidth) Then
                srcWidth = srcDIB.GetDIBWidth
                If (srcDIB.GetDIBHeight < srcHeight) Then srcHeight = srcDIB.GetDIBHeight
            ElseIf (srcDIB.GetDIBHeight < srcHeight) Then
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
            workingDIB.CreateFromExistingDIB srcDIB, newWidth, newHeight
            
        'The user is operating at 100% zoom.  Retrieve a subsection of the image, but do not scale it.
        Else
        
            'Calculate offsets, if any, for the image
            hOffset = previewTarget.GetOffsetX
            vOffset = previewTarget.GetOffsetY
            
            If (workingDIB.GetDIBWidth <> newWidth) Or (workingDIB.GetDIBHeight <> newHeight) Or (workingDIB.GetDIBColorDepth <> srcDIB.GetDIBColorDepth) Then
                workingDIB.CreateBlank newWidth, newHeight, srcDIB.GetDIBColorDepth
            Else
                workingDIB.ResetDIB
            End If
            
            GDI.BitBltWrapper workingDIB.GetDIBDC, 0, 0, newWidth, newHeight, srcDIB.GetDIBDC, hOffset, vOffset, vbSrcCopy
            workingDIB.SetInitialAlphaPremultiplicationState srcDIB.GetAlphaPremultiplication
            
        End If
        
        'Give the preview object a copy of this original, unmodified image data so it can show it to the user if requested
        If (Not previewTarget.HasOriginalImage) Then previewTarget.SetOriginalImage workingDIB
        
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
        .bytesPerPixel = (workingDIB.GetDIBColorDepth \ 8)
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

'PrepImageData is responsible for copying the relevant layer (or portion of a layer, if a selection is active)
' into a temporary pdDIB object, one that external filters and effects can freely modify.  This function also
' populates a relevant SafeArray struct and other variables related to image processing; filters and effects
' can use these values to figure out how to handle the associated preview DIB.
'
'In one of the better triumphs of PD's design, this function is shared between effect previews and effect
' finalization (e.g. permanently applying the effect). The isPreview parameter is used to notify effect
' functions of the source of a given request.  When isPreview is TRUE, the source image is automatically scaled
' to the size of the preview area; this means many fewer pixels to process, which in turn allows the tool dialog
' to render much faster.  (Importantly, for this to work, a valid pdFxPreview reference *must* be passed;
' it's queried for things like on-screen size and source zoom values.)
'
'The calling routine can optionally specify a different maximum progress bar value.  By default, this is
' the current layer (or selection's) width, but some routines run vertically and the progress bar maximum needs
' to be changed to match.
'
'The optional "ignoreSelection" parameter should always be set to FALSE in internal PD code.  The parameter exists
' for external Photoshop plugins (8bf) which may handle selection masking themselves.  When used, PD will not
' attempt to blend selection results itself and will instead rely on the target 8bf's native selection handling.
Public Sub PrepImageData(ByRef tmpSA As SafeArray2D, Optional isPreview As Boolean = False, Optional previewTarget As pdFxPreviewCtl, Optional newProgBarMax As Long = -1, Optional ByVal doNotTouchProgressBar As Boolean = False, Optional ByVal doNotUnPremultiplyAlpha As Boolean = False, Optional ByVal ignoreSelection As Boolean = False)

    'Reset the public "cancel current action" tracker
    g_cancelCurrentAction = False
    
    'Check for an active selection
    Dim selToolActive As Boolean
    selToolActive = PDImages.GetActiveImage.IsSelectionActive And PDImages.GetActiveImage.MainSelection.IsLockedIn And (Not ignoreSelection)
    
    'When selections are active, we may need to process pixels outside the active layer's boundaries.
    ' This requires special handling to render a correct image; basically, we must null-pad the current layer's
    ' pixel data to the size of its parent image.  When the filter completes, we'll extract the relevant portion
    ' of the effect and merge it automatically.
    Dim tmpLayer As pdLayer, selBounds As RectF
    If selToolActive Then
        Set tmpLayer = New pdLayer
        selBounds = PDImages.GetActiveImage.MainSelection.GetCompositeBoundaryRect
    End If
            
    'If this effect is just a preview, we need to calculate a new width and height relative to the size of the
    ' target preview window.
    Dim dstWidth As Double, dstHeight As Double
    Dim srcWidth As Double, srcHeight As Double
    Dim newWidth As Long, newHeight As Long
    
    '"This effect is not a preview"
    If (Not isPreview) Then
        
        If (workingDIB Is Nothing) Then Set workingDIB = New pdDIB
        m_PreviewWasRegenerated = True
        
        'Check for an active selection.  If one exists, we need to use it instead of the full layer DIB.
        ' (Note that no special processing is applied to the selected area - a full rectangle is passed to the
        ' source function, with no accounting for non-rectangular boundaries or feathering.  Those features are
        ' handled *after* the effect has already been processed.)
        If selToolActive Then
            
            'Before proceeding further, null-pad the layer in question.  This allows selections to work, even if
            ' they extend beyond the layer's borders.
            tmpLayer.CopyExistingLayer PDImages.GetActiveImage.GetActiveLayer
            tmpLayer.ConvertToNullPaddedLayer PDImages.GetActiveImage.Width, PDImages.GetActiveImage.Height
            
            'Crop the relevant portion of the layer out, using the selection boundary as a guide.  (We will only
            ' process the selection rectangle area for the effect; pixels outside that rectangle will be ignored.)
            workingDIB.CreateBlank selBounds.Width, selBounds.Height, PDImages.GetActiveImage.GetActiveDIB().GetDIBColorDepth
            GDI.BitBltWrapper workingDIB.GetDIBDC, 0, 0, selBounds.Width, selBounds.Height, tmpLayer.GetLayerDIB.GetDIBDC, selBounds.Left, selBounds.Top, vbSrcCopy
            workingDIB.SetInitialAlphaPremultiplicationState PDImages.GetActiveImage.GetActiveLayer.GetLayerDIB.GetAlphaPremultiplication
            
        Else
            workingDIB.CreateFromExistingDIB PDImages.GetActiveImage.GetActiveDIB()
        End If
        
    'This is an effect preview, meaning we need to prepare a custom image for operation.  Basically, we need to create
    ' a copy of the active layer that matches the dimensions of the preview window, while also accounting for things
    ' like boundary changes defined by an active selection.
    Else
        
        'Before doing anything else, see if we can simply re-use our previous preview image
        If (m_PreviousPreviewID = previewTarget.GetUniqueID) And (Not workingDIB Is Nothing) And (Not m_PreviousPreviewCopy Is Nothing) Then
            workingDIB.CreateFromExistingDIB m_PreviousPreviewCopy
            m_PreviewWasRegenerated = False
            
        'The preview target has changed, so we need regenerate our cached preview image from scratch.
        ' (This is time-consuming and complicated, so try to avoid it whenever possible.)
        Else
        
            If (workingDIB Is Nothing) Then Set workingDIB = New pdDIB
            m_PreviewWasRegenerated = True
            
            'Start by calculating the source area for the preview.  This changes based on several criteria:
            ' 1) Is the preview area set to "fit full image" or "100% zoom"?
            ' 2) Is a selection active?  If so, we only want to preview the selected area.  (I may change this
            '     behavior in the future, so the user can actually see the fully composited result of any changes.)
            
            'The full image is being previewed.  Retrieve the entire thing.
            If previewTarget.ViewportFitFullImage Then
            
                If selToolActive Then
                    srcWidth = selBounds.Width
                    srcHeight = selBounds.Height
                Else
                    srcWidth = PDImages.GetActiveImage.GetActiveDIB().GetDIBWidth
                    srcHeight = PDImages.GetActiveImage.GetActiveDIB().GetDIBHeight
                End If
            
            'Only a section of the image is being preview (at 100% zoom).  Retrieve just that section.
            Else
            
                srcWidth = previewTarget.GetPreviewWidth
                srcHeight = previewTarget.GetPreviewHeight
                
                'If a selection is active, and the selected area is smaller than the preview window,
                ' constrain the source area to the selection boundaries.
                If selToolActive Then
                
                    If (selBounds.Width < srcWidth) Then
                        srcWidth = selBounds.Width
                        If (selBounds.Height < srcHeight) Then srcHeight = selBounds.Height
                    ElseIf (selBounds.Height < srcHeight) Then
                        srcHeight = selBounds.Height
                    End If
                    
                Else
                    
                    If (PDImages.GetActiveImage.GetActiveDIB().GetDIBWidth < srcWidth) Then
                        srcWidth = PDImages.GetActiveImage.GetActiveDIB().GetDIBWidth
                        If (PDImages.GetActiveImage.GetActiveDIB().GetDIBHeight < srcHeight) Then srcHeight = PDImages.GetActiveImage.GetActiveDIB().GetDIBHeight
                    ElseIf (PDImages.GetActiveImage.GetActiveDIB().GetDIBHeight < srcHeight) Then
                        srcHeight = PDImages.GetActiveImage.GetActiveDIB().GetDIBHeight
                    End If
                    
                End If
                
            End If
            
            'Destination width/height are generally the dimensions of the preview box, taking into account aspect ratio.
            ' The only exception to this is when the image is actually smaller than the preview area - in that case
            ' we just use the whole image.
            dstWidth = previewTarget.GetPreviewWidth
            dstHeight = previewTarget.GetPreviewHeight
                    
            If (srcWidth > dstWidth) Or (srcHeight > dstHeight) Then
                PDMath.ConvertAspectRatio srcWidth, srcHeight, dstWidth, dstHeight, newWidth, newHeight
            Else
                newWidth = srcWidth
                newHeight = srcHeight
            End If
            
            'The area may be offset from the (0, 0) position if the user has elected to drag the preview area
            Dim hOffset As Long, vOffset As Long
            
            'Next, create the temporary object (called "workingDIB") at the calculated preview dimensions.  All editing
            ' actions are applied to this DIB; if the user does not cancel the action, that DIB will be copied over the
            ' primary image.  If they cancel, we'll simply discard the temporary DIB.
            
            'Just like with a full image, if a selection is active, we only want to process the selected area.
            If selToolActive Then
                
                'Before proceeding further, make a copy of the active layer, and null-pad it to the size of the parent image.
                ' This will allow any possible selection to work, regardless of a layer's actual area.
                tmpLayer.CopyExistingLayer PDImages.GetActiveImage.GetActiveLayer
                tmpLayer.ConvertToNullPaddedLayer PDImages.GetActiveImage.Width, PDImages.GetActiveImage.Height
                
                'The user is using "fit full image on-screen" mode for this preview.  Retrieve a tiny version of the selection.
                If previewTarget.ViewportFitFullImage Then
                    workingDIB.CreateBlank newWidth, newHeight, 32, 0, 0
                    GDI_Plus.GDIPlus_StretchBlt workingDIB, 0, 0, newWidth, newHeight, tmpLayer.GetLayerDIB, selBounds.Left, selBounds.Top, selBounds.Width, selBounds.Height, , GP_IM_Bilinear
                    
                'The user is operating at 100% zoom.  Retrieve a subsection of the selected area, but do not scale it.
                Else
                
                    'Calculate offsets, if any, for the selected area
                    hOffset = previewTarget.GetOffsetX
                    vOffset = previewTarget.GetOffsetY
                    workingDIB.CreateBlank newWidth, newHeight, 32, 0, 0
                    
                    GDI.BitBltWrapper workingDIB.GetDIBDC, 0, 0, dstWidth, dstHeight, tmpLayer.GetLayerDIB.GetDIBDC, hOffset + selBounds.Left, vOffset + selBounds.Top, vbSrcCopy
                    workingDIB.SetInitialAlphaPremultiplicationState tmpLayer.GetLayerDIB.GetAlphaPremultiplication
                    
                End If
                
            'If a selection is not currently active, this step is incredibly simple!
            Else
                
                'The user is using "fit full image on-screen" mode for this preview.  Retrieve a tiny version of the image
                If previewTarget.ViewportFitFullImage Then
                    workingDIB.CreateFromExistingDIB PDImages.GetActiveImage.GetActiveDIB(), newWidth, newHeight
                    
                'The user is operating at 100% zoom.  Retrieve a subsection of the image, but do not scale it.
                Else
                
                    'Calculate offsets, if any, for the image
                    hOffset = previewTarget.GetOffsetX
                    vOffset = previewTarget.GetOffsetY
                    workingDIB.CreateBlank newWidth, newHeight, 32, 0, 0
                    
                    GDI.BitBltWrapper workingDIB.GetDIBDC, 0, 0, dstWidth, dstHeight, PDImages.GetActiveImage.GetActiveDIB().GetDIBDC, hOffset, vOffset, vbSrcCopy
                    workingDIB.SetInitialAlphaPremultiplicationState PDImages.GetActiveImage.GetActiveDIB().GetAlphaPremultiplication
                    
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
    ' Note that individual tools can override this behavior - this is helpful in certain cases, e.g. area filters
    ' like blur, where *not* premultiplying alpha causes black RGB values from transparent areas to be "picked up" by
    ' other areas.
    If (workingDIB.GetDIBColorDepth = 32) And (workingDIB.GetAlphaPremultiplication <> doNotUnPremultiplyAlpha) Then workingDIB.SetAlphaPremultiplication doNotUnPremultiplyAlpha
    
    'If a selection is active, make a backup of the selected area.  (We do this regardless of whether the current
    ' action is a preview or not.)
    If selToolActive Then
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
        .bytesPerPixel = (workingDIB.GetDIBColorDepth \ 8)
        If isPreview Then
            If previewTarget.ViewportFitFullImage Then
                .previewModifier = workingDIB.GetDIBWidth / PDImages.GetActiveImage.GetActiveDIB().GetDIBWidth
            Else
                .previewModifier = 1#
            End If
        Else
            .previewModifier = 1#
        End If
    End With
    
    'With our temporary DIB successfully created, populate the relevant SafeArray variable, so the caller has direct access to the DIB.
    ' (TODO: move this into individual function calls.  It's stupid to do it here.)
    PrepSafeArray tmpSA, workingDIB
    
    'Set up the progress bar (only if this is NOT a preview, mind you - during previews, the progress bar is not touched)
    If (Not isPreview) And (Not doNotTouchProgressBar) Then
        If (newProgBarMax = -1) Then
            SetProgBarMax (curDIBValues.Left + curDIBValues.Width)
        Else
            SetProgBarMax newProgBarMax
        End If
    End If
    
    'If desired, the statement below can be used to verify that the function created a working DIB at the proper dimensions
    'Debug.Print "prepImageData worked: " & workingDIB.getDIBHeight & ", " & workingDIB.getDIBWidth & " (" & workingDIB.GetDIBStride & ")" & ", " & workingDIB.GetDIBPointer
    
End Sub

'The counterpart to PrepImageData, FinalizeImageData copies the working DIB back into the source image,
' then renders everything to the screen.  Like PrepImageData(), a preview target can also be named;
' if it is, FinalizeImageData will rely on all preview-related values calculated by PrepImageData()
' because these two functions must always be called in-order as a pair.
'
'Note that unlike PrepImageData, this function has to do a lot of extra processing when selections are active.
' The selection mask must be scanned for each pixel, and the results blended with the original image as appropriate.
' (This is the price we pay for full selection feathering support.)  This step can be forcibly bypassed by setting
' the optional ignoreSelection parameter to TRUE, but this should never be used for internal PD tools.  (The setting
' exists only for Photoshop-style 8bf plugins, which are allowed to use their own selection blending logic.)
Public Sub FinalizeImageData(Optional isPreview As Boolean = False, Optional previewTarget As pdFxPreviewCtl, Optional ByVal alphaAlreadyPremultiplied As Boolean = False, Optional ByVal ignoreSelection As Boolean = False)
    
    'If the user canceled the current action, disregard the working DIB and exit immediately.  The central processor
    ' will take care of additional clean-up.
    If (Not isPreview) And g_cancelCurrentAction Then
        Set workingDIB = Nothing
        Exit Sub
    End If
    
    'Regardless of whether or not this is a preview, we process selections identically - by merging the newly modified
    ' workingDIB with its original version (as stored in workingDIBBackup), while accounting for any selection intricacies.
    Dim selToolActive As Boolean
    selToolActive = PDImages.GetActiveImage.IsSelectionActive
    If selToolActive Then selToolActive = PDImages.GetActiveImage.MainSelection.IsLockedIn And (Not ignoreSelection)
    
    If selToolActive Then
        
        'Retrieve the current selection boundaries
        Dim selBounds As RectF
        selBounds = PDImages.GetActiveImage.MainSelection.GetCompositeBoundaryRect
        
        'Before continuing further, create a copy of the selection mask at the relevant image size.
        ' (Note that "relevant size" differs between effect previews, which tend to be much smaller, and the final
        '  application of the effect to the current layer.)
        Dim selMaskCopy As pdDIB
        
        'If this is a preview, resize the selection mask to match the preview size
        If isPreview Then
            
            If m_PreviewWasRegenerated Or (m_SelectionMaskBackup Is Nothing) Then
            
                If (selMaskCopy Is Nothing) Then Set selMaskCopy = New pdDIB
                selMaskCopy.CreateBlank workingDIB.GetDIBWidth, workingDIB.GetDIBHeight, 32, 0, 0
                
                'The preview is a shrunk version of the full image.  Shrink the selection mask to match.
                If previewTarget.ViewportFitFullImage Then
                    GDI_Plus.GDIPlus_StretchBlt selMaskCopy, 0, 0, workingDIB.GetDIBWidth, workingDIB.GetDIBHeight, PDImages.GetActiveImage.MainSelection.GetCompositeMaskDIB, selBounds.Left, selBounds.Top, selBounds.Width, selBounds.Height, interpolationType:=GP_IM_Default, dstCopyIsOkay:=True
                
                'The preview is a 100% zoom portion of the image.  Copy only the relevant part of the selection mask into the
                ' selection processing DIB.
                Else
                    GDI.BitBltWrapper selMaskCopy.GetDIBDC, 0, 0, selMaskCopy.GetDIBWidth, selMaskCopy.GetDIBHeight, PDImages.GetActiveImage.MainSelection.GetMaskDC(), selBounds.Left + previewTarget.GetOffsetX, selBounds.Top + previewTarget.GetOffsetY, vbSrcCopy
                End If
                
                Set m_SelectionMaskBackup = selMaskCopy
                
            Else
                Set selMaskCopy = m_SelectionMaskBackup
            End If
            
        'If this is *not* a preview, simply crop out the portion of the selection mask matching the current preview area.
        ' (TODO: this copy really isn't necessary; just point the array at the actual selection mask, instead!)
        Else
            If (selMaskCopy Is Nothing) Then Set selMaskCopy = New pdDIB
            selMaskCopy.CreateBlank selBounds.Width, selBounds.Height, 32, 0, 0
            GDI.BitBltWrapper selMaskCopy.GetDIBDC, 0, 0, selMaskCopy.GetDIBWidth, selMaskCopy.GetDIBHeight, PDImages.GetActiveImage.MainSelection.GetMaskDC(), selBounds.Left, selBounds.Top, vbSrcCopy
        End If
        
        'We now have a DIB that represents the selection mask at the same offset and size as the workingDIB.  This allows
        ' us to process the selected area identically, regardless of whether this is a preview or a true full-DIB operation.
        
        'A few rare functions actually change the color depth of the image.  Check for that now,
        ' and make sure the workingDIB and workingDIBBackup DIBs are the same bit-depth.
        '(TODO: I believe PD has fully transitioned to 32-bpp images, but just in case,
        ' let's report any discrepancies .)
        If (workingDIB.GetDIBColorDepth <> workingDIBBackup.GetDIBColorDepth) Then
            PDDebug.LogAction "WARNING!  24-bpp image has been forcefully generated by effect.  Revisit!"
            workingDIBBackup.ConvertTo32bpp
        End If
        
        'If the current working DIB (or its backup copy, which stores a result of how the region looked before we
        ' applied the effect) is *not* premultiplied, premultiply it now.  This allows us to blend the results of
        ' the new effect onto the destination image much more quickly.
        If (Not workingDIB.GetAlphaPremultiplication) Then workingDIB.SetAlphaPremultiplication True
        If (Not workingDIBBackup.GetAlphaPremultiplication) Then workingDIBBackup.SetAlphaPremultiplication True
        
        'Next, point three arrays at three images: the original image, the newly modified image, and the selection mask copy
        ' we just created.  (We need a copy of the original image, because selection feathering requires us to blend pixels
        ' between the original copy, and the new effect-processed copy.)
        
        'Prepare a few image arrays (and array headers) in advance.
        Dim pxEffect() As Byte, saEffect As SafeArray1D, ptrWD As Long, strideWD As Long
        workingDIB.WrapArrayAroundScanline pxEffect, saEffect, 0
        ptrWD = saEffect.pvData
        strideWD = saEffect.cElements
        
        Dim pxSelection() As Byte, saSelection As SafeArray1D, ptrSel As Long, strideSel As Long
        selMaskCopy.WrapArrayAroundScanline pxSelection, saSelection, 0
        ptrSel = saSelection.pvData
        strideSel = saSelection.cElements
        
        Dim pxDst() As Byte, saDst As SafeArray1D, ptrDst As Long, strideDst As Long
        workingDIBBackup.WrapArrayAroundScanline pxDst, saDst, 0
        ptrDst = saDst.pvData
        strideDst = saDst.cElements
        
        Dim x As Long, y As Long
        Dim thisAlpha As Long, blendAlpha As Double
        
        Dim dstPxWidth As Long, xOffsetDst As Long, effectX As Long
        dstPxWidth = workingDIBBackup.GetDIBColorDepth \ 8
        
        Dim workingDIBCD As Long
        workingDIBCD = workingDIB.GetDIBColorDepth \ 8
        
        Const ONE_DIV_255 As Double = 1# / 255#
        
        Dim newR As Long, newG As Long, newB As Long, newA As Long
        Dim oldR As Long, oldG As Long, oldB As Long, oldA As Long
        
        For y = 0 To workingDIB.GetDIBHeight - 1
            
            'Update all array pointers to point at the current line in each array
            saEffect.pvData = ptrWD + strideWD * y
            saSelection.pvData = ptrSel + strideSel * y
            saDst.pvData = ptrDst + strideDst * y
            
        For x = 0 To workingDIB.GetDIBWidth - 1
            
            xOffsetDst = x * dstPxWidth
            effectX = x * workingDIBCD
            
            'Retrieve the selection mask value at this position.  Its value determines how this pixel is handled.
            thisAlpha = pxSelection(x * 4)
            
            'Transparent mask pixels are completely ignored (e.g. they are not part of the selected area)
            If (thisAlpha <> 0) Then
                
                newB = pxEffect(effectX)
                newG = pxEffect(effectX + 1)
                newR = pxEffect(effectX + 2)
                newA = pxEffect(effectX + 3)
                
                'Fully selected pixels are replaced wholesale by the effect results.
                If (thisAlpha = 255) Then
                    pxDst(xOffsetDst) = newB
                    pxDst(xOffsetDst + 1) = newG
                    pxDst(xOffsetDst + 2) = newR
                    pxDst(xOffsetDst + 3) = newA
                    
                'Partially selected pixels are calculated as a weighted average of the old and new pixels.
                ' (Note that this is *not* an alpha-blend operation!  It is a weighted average between the old and
                '  new pixel results, which produces a totally different output.)
                Else
                    
                    blendAlpha = thisAlpha * ONE_DIV_255
                    
                    'Retrieve the old (original, unmodified) RGB values
                    oldB = pxDst(xOffsetDst)
                    oldG = pxDst(xOffsetDst + 1)
                    oldR = pxDst(xOffsetDst + 2)
                    oldA = pxDst(xOffsetDst + 3)
                    
                    'Calculate a weighted blend of the old and new pixel values.  Because they are premultiplied, we do not
                    ' need to deal with the effect this has on alpha values.
                    pxDst(xOffsetDst) = (blendAlpha * newB) + oldB * (1# - blendAlpha)
                    pxDst(xOffsetDst + 1) = (blendAlpha * newG) + oldG * (1# - blendAlpha)
                    pxDst(xOffsetDst + 2) = (blendAlpha * newR) + oldR * (1# - blendAlpha)
                    pxDst(xOffsetDst + 3) = (blendAlpha * newA) + oldA * (1# - blendAlpha)
                    
                End If
                    
            End If
            
        Next x
        Next y
        
        'Safely deallocate all image arrays
        workingDIB.UnwrapArrayFromDIB pxEffect
        selMaskCopy.UnwrapArrayFromDIB pxSelection
        workingDIBBackup.UnwrapArrayFromDIB pxDst
            
    End If
        
    'Processing past this point is contingent on whether or not the current action is a preview.
        
    'If this is not a preview, simply copy the processed data back into the active DIB
    If (Not isPreview) Then
        
        'If a selection is active, copy the processed area into its proper place.
        If selToolActive Then
            
            'Before applying the selected area back onto the image, we need to null-pad the original layer.  (This is not done
            ' by prepImageData, because the user may elect to cancel a running action - and if they do that, we want to leave
            ' the original image untouched!  Thus, only the workingLayer has been null-padded.)
            '
            'TODO: figure out if this is necessary for normal layers, and not just ones with affine transforms.
            PDImages.GetActiveImage.GetActiveLayer.ConvertToNullPaddedLayer PDImages.GetActiveImage.Width, PDImages.GetActiveImage.Height
            
            'Un-pad any null pixels we may have added as part of the selection interaction
            If (Not workingDIBBackup Is Nothing) Then
                If (workingDIBBackup.GetDIBColorDepth = 32) And (Not alphaAlreadyPremultiplied) And (Not workingDIBBackup.GetAlphaPremultiplication) Then workingDIBBackup.SetAlphaPremultiplication True
                GDI.BitBltWrapper PDImages.GetActiveImage.GetActiveDIB().GetDIBDC, selBounds.Left, selBounds.Top, selBounds.Width, selBounds.Height, workingDIBBackup.GetDIBDC, 0, 0, vbSrcCopy
                PDImages.GetActiveImage.GetActiveLayer.CropNullPaddedLayer
            End If
        
        'If a selection is not active, replace the entire DIB with the contents of the working DIB
        Else
            
            If (workingDIB.GetDIBColorDepth = 32) And (Not alphaAlreadyPremultiplied) And (Not workingDIB.GetAlphaPremultiplication) Then
                workingDIB.SetAlphaPremultiplication True
            Else
                workingDIB.SetInitialAlphaPremultiplicationState True
            End If
            
            PDImages.GetActiveImage.GetActiveDIB().CreateFromExistingDIB workingDIB
            
        End If
        
        'workingDIB and its backup have served their purposes, so erase them from memory
        Set workingDIB = Nothing
        Set workingDIBBackup = Nothing
        
        'If we're setting data to the screen, we can reasonably assume that the progress bar should be reset
        ReleaseProgressBar
        
        'Notify the parent of the target layer of our changes
        PDImages.GetActiveImage.NotifyImageChanged UNDO_Layer, PDImages.GetActiveImage.GetActiveLayerIndex
        
        'Pass control to the viewport renderer, which will perform the actual rendering
        Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), FormMain.MainCanvas(0)
        
        Message "Finished."
    
    'If this *is* a preview, we need to repaint a preview box
    Else
        
        'If a selection is active, use the contents of workingDIBBackup instead of workingDIB to render the preview
        If selToolActive Then
            
            If (workingDIBBackup.GetDIBColorDepth = 32) And (Not alphaAlreadyPremultiplied) And (Not workingDIBBackup.GetAlphaPremultiplication) Then
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
                If (Not workingDIB.GetAlphaPremultiplication) Then workingDIB.SetAlphaPremultiplication True
            Else
                workingDIB.SetInitialAlphaPremultiplicationState True
            End If
            
            previewTarget.SetFXImage workingDIB, weCanHandleCM
        
        End If
        
    End If
    
End Sub
