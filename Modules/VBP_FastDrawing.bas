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
' projects IF you provide attribution.  For more information, please visit http://www.tannerhelland.com/photodemon/#license
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
    MinX As Long
    MinY As Long
    
    'The highest coordinate the filter is allowed to check.  This is almost always (width, height).
    MaxX As Long
    MaxY As Long
    
    'The colorDepth of the current layer, specified as BITS per pixel; this will always be 24 or 32
    ColorDepth As Long
    
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


'prepImageData's job is to copy the relevant layer into a temporary object, which individual filters and effects will
' then operate on.  prepImageData() also populates the relevant SafeArray object and a host of other variables, which
' filters and effects can then copy locally to ensure the fastest possible runtime speed.
'
'If the filter will be rendering a preview only, it can specify the fxPreview control that will receive the preview effect.
' This function will automatically adjust its parameters accordingly, and the filter routine will not have to make any
' modifications to its code.
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
    
        'Check for an active selection; if one is present, use that instead of the full layer
        If pdImages(CurrentImage).selectionActive Then
            
            'Make a working copy of the image data within the selection
            workingLayer.createBlank pdImages(CurrentImage).mainSelection.boundWidth, pdImages(CurrentImage).mainSelection.boundHeight, pdImages(CurrentImage).mainLayer.getLayerColorDepth
            BitBlt workingLayer.getLayerDC, 0, 0, pdImages(CurrentImage).mainSelection.boundWidth, pdImages(CurrentImage).mainSelection.boundHeight, pdImages(CurrentImage).mainLayer.getLayerDC, pdImages(CurrentImage).mainSelection.boundLeft, pdImages(CurrentImage).mainSelection.boundTop, vbSrcCopy
            
        Else
            workingLayer.createFromExistingLayer pdImages(CurrentImage).mainLayer
        End If
    
    'If this IS a preview, more work is involved.
    Else
    
        'The destination width and height is the width and height of the preview picture box.
        dstWidth = previewTarget.getPreviewPic.ScaleWidth
        dstHeight = previewTarget.getPreviewPic.ScaleHeight
            
        'The source values need to be adjusted contingent on whether this is a selection or a full-image preview.
        If pdImages(CurrentImage).selectionActive Then
            srcWidth = pdImages(CurrentImage).mainSelection.boundWidth
            srcHeight = pdImages(CurrentImage).mainSelection.boundHeight
        Else
            srcWidth = pdImages(CurrentImage).mainLayer.getLayerWidth
            srcHeight = pdImages(CurrentImage).mainLayer.getLayerHeight
        End If
            
        'Now, use that aspect ratio to determine a proper size for our temporary layer
        If (srcWidth > dstWidth) Or (srcHeight > dstHeight) Then
            convertAspectRatio srcWidth, srcHeight, dstWidth, dstHeight, newWidth, newHeight
        Else
            newWidth = srcWidth
            newHeight = srcHeight
        End If
        
        'Now, create the workingLayer object using the calculated dimensions.
        
        'If a selection is active, quite a bit of additional work must be applied.  The selected area must be "chopped" out of the image,
        ' and converted to 32bpp (so that feathering and antialiased selections are properly handled).
        If pdImages(CurrentImage).selectionActive Then
        
            'Start by chopping out the full rectangular bounding area of the selection, and placing it inside the workingLayer object.
            ' This is done at the same color depth as the source image.
            
            'Note that we do this in two steps.  First, we create a temporary layer that contains the rectangular bounding area at its
            ' original size.  Next, we create the working layer, which is the PROPERLY SIZED version of the data (e.g. shrunk if this
            ' is a preview).  These steps could be combined into one.
            Dim copyLayer As pdLayer
            Set copyLayer = New pdLayer
            copyLayer.createBlank pdImages(CurrentImage).mainSelection.boundWidth, pdImages(CurrentImage).mainSelection.boundHeight, pdImages(CurrentImage).mainLayer.getLayerColorDepth
            BitBlt copyLayer.getLayerDC, 0, 0, pdImages(CurrentImage).mainSelection.boundWidth, pdImages(CurrentImage).mainSelection.boundHeight, pdImages(CurrentImage).mainLayer.getLayerDC, pdImages(CurrentImage).mainSelection.boundLeft, pdImages(CurrentImage).mainSelection.boundTop, vbSrcCopy
            workingLayer.createFromExistingLayer copyLayer, newWidth, newHeight
            copyLayer.eraseLayer
            Set copyLayer = Nothing
            
            'Next, make a copy of the selection mask at the same dimensions as the preview.  We will use this to remove the sections of the
            ' selection that are not selected.  (Say that 10 times fast...lol)
            Dim tmpSelectionMaskCopy As pdLayer
            Set tmpSelectionMaskCopy = New pdLayer
            tmpSelectionMaskCopy.createBlank pdImages(CurrentImage).mainSelection.boundWidth, pdImages(CurrentImage).mainSelection.boundHeight, 24
            BitBlt tmpSelectionMaskCopy.getLayerDC, 0, 0, pdImages(CurrentImage).mainSelection.boundWidth, pdImages(CurrentImage).mainSelection.boundHeight, pdImages(CurrentImage).mainSelection.selMask.getLayerDC, pdImages(CurrentImage).mainSelection.boundLeft, pdImages(CurrentImage).mainSelection.boundTop, vbSrcCopy
            Set tmpSelectionMask = New pdLayer
            tmpSelectionMask.createFromExistingLayer tmpSelectionMaskCopy, newWidth, newHeight
            tmpSelectionMaskCopy.eraseLayer
            Set tmpSelectionMaskCopy = Nothing
            
            'Now, convert the working layer to 32bpp.  Unselected areas must be made transparent, so 24bpp won't work.
            Dim already32bpp As Boolean
            
            If workingLayer.getLayerColorDepth = 32 Then
                already32bpp = True
            Else
                already32bpp = False
                workingLayer.convertTo32bpp
            End If
            
            'Next, we are going to remove any pixels that are not part of the selected area.  We use the selection mask
            ' (as stored in tmpSelectionMaskCopy) as our guide.
            Dim wlImageData() As Byte
            Dim wlSA As SAFEARRAY2D
            prepSafeArray wlSA, workingLayer
            CopyMemory ByVal VarPtrArray(wlImageData()), VarPtr(wlSA), 4
            
            Dim selImageData() As Byte
            Dim selSA As SAFEARRAY2D
            prepSafeArray selSA, tmpSelectionMask
            CopyMemory ByVal VarPtrArray(selImageData()), VarPtr(selSA), 4
                        
            Dim x As Long, y As Long
            For x = 0 To workingLayer.getLayerWidth - 1
            For y = 0 To workingLayer.getLayerHeight - 1
                
                'If the image is already 32bpp, instead of relying solely on the selection mask values, we need to blend any
                ' transparent pixels with the selection mask's transparency - this will give an accurate portrayal of how
                ' the final processed area will look.
                If already32bpp Then
                    wlImageData(x * 4 + 3, y) = wlImageData(x * 4 + 3, y) * (selImageData(x * 3, y) / 255)
                Else
                    wlImageData(x * 4 + 3, y) = selImageData(x * 3, y)
                End If
                
            Next y
            Next x
            
            'Working layer is now a 32bpp image that accurately represents the selected area.
            
            'With our work complete, point both ImageData() arrays away from their DIBs and deallocate them
            CopyMemory ByVal VarPtrArray(wlImageData), 0&, 4
            Erase wlImageData
            CopyMemory ByVal VarPtrArray(selImageData), 0&, 4
            Erase selImageData
            
        
        'If a selection is not currently active, this step is incredibly simple!
        Else
            workingLayer.createFromExistingLayer pdImages(CurrentImage).mainLayer, newWidth, newHeight
        End If
        
        'Give the preview object a copy of this image data so it can show it to the user if requested
        If Not previewTarget.hasOriginalImage Then previewTarget.setOriginalImage workingLayer
        
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
        .MinX = 0
        .MinY = 0
        .MaxX = workingLayer.getLayerWidth - 1
        .MaxY = workingLayer.getLayerHeight - 1
        .ColorDepth = workingLayer.getLayerColorDepth
        .BytesPerPixel = (workingLayer.getLayerColorDepth \ 8)
        .LayerX = 0
        .LayerY = 0
        .previewModifier = workingLayer.getLayerWidth / pdImages(CurrentImage).mainLayer.getLayerWidth
    End With

    'Set up the progress bar (only if this is NOT a preview, mind you - during previews, the progress bar is not touched)
    If Not isPreview Then
        If newProgBarMax = -1 Then
            SetProgBarMax (curLayerValues.Left + curLayerValues.Width)
        Else
            SetProgBarMax newProgBarMax
        End If
    End If
    
    'pdMsgBox "prepImageData worked: " & workingLayer.getLayerHeight & ", " & workingLayer.getLayerWidth & " (" & workingLayer.getLayerArrayWidth & ")" & ", " & workingLayer.getLayerDIBits

End Sub


'The counterpart to prepImageData, finalizeImageData copies the working layer back into the source image, then renders
' everything to the screen.  Like prepImageData(), a preview target can also be named.  In this case, finalizeImageData
' will rely on the preview-related values calculated by prepImageData(), as it's presumed that preImageData will ALWAYS
' be called before this routine.
Public Sub finalizeImageData(Optional isPreview As Boolean = False, Optional previewTarget As fxPreviewCtl)

    'If the user has canceled the current action, disregard the working layer and exit immediately.  The central processor
    ' will take care of additional clean-up.
    If (Not isPreview) And cancelCurrentAction Then
        
        workingLayer.eraseLayer
        Set workingLayer = Nothing
        
        Exit Sub
        
    End If

    Dim wlImageData() As Byte
    Dim wlSA As SAFEARRAY2D
    
    Dim selImageData() As Byte
    Dim selSA As SAFEARRAY2D
    
    Dim x As Long, y As Long
    
    'If this is not a preview, our job is simple - get the newly processed DIB rendered to the screen.
    If Not isPreview Then
        
        Message "Rendering image to screen..."
        
        'If a selection is active, we need to paint the selected area back onto the image.  This can be simple (e.g. a square selection),
        ' or hideously complex (e.g. a "magic wand" selection).
        If pdImages(CurrentImage).selectionActive Then
        
            'In the future, we could optimize this function to check for plain square selections.  If found, we can simply BitBlt the selected
            ' area onto the main image, rather than doing a complex check for partial selections or non-square selection regions.
            'If pdImages(CurrentImage).mainSelection.isPlainOldSquare Then...
            ' Simple: BitBlt pdImages(CurrentImage).mainLayer.getLayerDC, pdImages(CurrentImage).mainSelection.selLeft, pdImages(CurrentImage).mainSelection.selTop, workingLayer.getLayerWidth, workingLayer.getLayerHeight, workingLayer.getLayerDC, 0, 0, vbSrcCopy
            'Else
            prepSafeArray wlSA, workingLayer
            CopyMemory ByVal VarPtrArray(wlImageData()), VarPtr(wlSA), 4
            
            prepSafeArray selSA, pdImages(CurrentImage).mainSelection.selMask
            CopyMemory ByVal VarPtrArray(selImageData()), VarPtr(selSA), 4
            
            Dim dstImageData() As Byte
            Dim dstSA As SAFEARRAY2D
            prepSafeArray dstSA, pdImages(CurrentImage).mainLayer
            CopyMemory ByVal VarPtrArray(dstImageData()), VarPtr(dstSA), 4
            
            Dim leftOffset As Long, topOffset As Long
            leftOffset = pdImages(CurrentImage).mainSelection.boundLeft
            topOffset = pdImages(CurrentImage).mainSelection.boundTop
                        
            Dim i As Long
            Dim thisAlpha As Long
            Dim blendAlpha As Double
            
            Dim dstQuickVal As Long
            dstQuickVal = pdImages(CurrentImage).mainLayer.getLayerColorDepth \ 8
            
            Dim workingLayerCD As Long
            workingLayerCD = workingLayer.getLayerColorDepth \ 8
            
            For x = 0 To workingLayer.getLayerWidth - 1
            For y = 0 To workingLayer.getLayerHeight - 1
                
                thisAlpha = selImageData((leftOffset + x) * 3, topOffset + y)
                
                Select Case thisAlpha
                    
                    'This pixel is not part of the selection, so ignore it
                    Case 0
                    
                    'This pixel completely replaces the destination one, so simply copy it over
                    Case 255
                        For i = 0 To dstQuickVal - 1
                            dstImageData((leftOffset + x) * dstQuickVal + i, topOffset + y) = wlImageData(x * workingLayerCD + i, y)
                        Next i
                        
                    'This pixel is antialiased or feathered, so it needs to be blended with the destination at the level specified
                    ' by the selection mask.
                    Case Else
                        blendAlpha = thisAlpha / 255
                        For i = 0 To dstQuickVal - 1
                            dstImageData((leftOffset + x) * dstQuickVal + i, topOffset + y) = BlendColors(dstImageData((leftOffset + x) * dstQuickVal + i, topOffset + y), wlImageData(x * workingLayerCD + i, y), blendAlpha)
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
            
            
        Else
            If workingLayer.getLayerColorDepth = 32 Then pdImages(CurrentImage).mainLayer.convertTo32bpp
            BitBlt pdImages(CurrentImage).mainLayer.getLayerDC, curLayerValues.LayerX, curLayerValues.LayerY, curLayerValues.Width, curLayerValues.Height, workingLayer.getLayerDC, 0, 0, vbSrcCopy
        End If
                
        'workingLayer has served its purpose, so erase it from memory
        workingLayer.eraseLayer
        Set workingLayer = Nothing
                
        'If we're setting data to the screen, we can reasonably assume that the progress bar should be reset
        SetProgBarVal 0
        
        'Pass control to the viewport renderer, which will perform the actual rendering
        ScrollViewport FormMain.ActiveForm
        
        Message "Finished."
        
    Else
        
        'If this is a preview, and a selection mask was in use, and we forced it to 32bpp to remove the unselected areas,
        ' we now need to restore those areas to their original state.  Otherwise the preview will look funky.
        If Not tmpSelectionMask Is Nothing Then
            
            'Next, we are going to remove any pixels that are not part of the selection mask.
            prepSafeArray wlSA, workingLayer
            CopyMemory ByVal VarPtrArray(wlImageData()), VarPtr(wlSA), 4
            
            prepSafeArray selSA, tmpSelectionMask
            CopyMemory ByVal VarPtrArray(selImageData()), VarPtr(selSA), 4
                        
            Dim already32bpp As Boolean
            If pdImages(CurrentImage).mainLayer.getLayerColorDepth = 24 Then already32bpp = False Else already32bpp = True
            
            For x = 0 To workingLayer.getLayerWidth - 1
            For y = 0 To workingLayer.getLayerHeight - 1
                If already32bpp Then
                    If selImageData(x * 3, y) = 0 Then wlImageData(x * 4 + 3, y) = 0
                Else
                    wlImageData(x * 4 + 3, y) = selImageData(x * 3, y)
                End If
            Next y
            Next x
            
            'With our work complete, point both ImageData() arrays away from their DIBs and deallocate them
            CopyMemory ByVal VarPtrArray(wlImageData), 0&, 4
            Erase wlImageData
            CopyMemory ByVal VarPtrArray(selImageData), 0&, 4
            Erase selImageData
            
            'We can now erase our temporary copy of the selection mask
            tmpSelectionMask.eraseLayer
            Set tmpSelectionMask = Nothing
            
        End If
        
        'If the current layer is 32bpp, precomposite it against a checkerboard background before rendering
        If workingLayer.getLayerColorDepth = 32 Then workingLayer.compositeBackgroundColor
            
        'Give the preview object a copy of the layer data used to generate the preview
        previewTarget.setFXImage workingLayer
        
        'workingLayer has served its purpose, so erase it from memory
        workingLayer.eraseLayer
        Set workingLayer = Nothing
        
    End If
    
End Sub

