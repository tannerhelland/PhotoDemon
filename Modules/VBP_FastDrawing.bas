Attribute VB_Name = "FastDrawing"
'***************************************************************************
'Fast API Graphics Routines Interface
'Copyright ©2001-2013 by Tanner Helland
'Created: 12/June/01
'Last updated: 03/May/13
'Last update: rewrote all selection handling code to work with new selection mask code.  All selections are now represented as masks,
'              e.g. a grayscale image whose values correspond to the amount of "selection" a given pixel has, from "completely selected"
'              to "completely unselected".  This allows for advanced selection operations like feathering, arbitrary shapes, etc.
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

Public workingLayer As pdLayer

'prepImageData will fill a variable of this type with everything a filter or effect could possibly want to know
' about the layer it's operating on.  Filters are free to ignore this data, but it is always made available.
Public Type FilterInfo
    Left As Long            'Coordinates of the top-left location the filter is supposed to operate on
    Top As Long
    Right As Long           'Right and Bottom could be inferred, but we do it here to minimize effort on the calling routine's part
    Bottom As Long
    Width As Long           'Dimensions of the area the filter is supposed to operate on
    Height As Long
    MinX As Long            'The lowest coordinate the filter is allowed to check.  This is almost always (0, 0)
    MinY As Long
    MaxX As Long            'The highest coordinate the filter is allowed to check.  This is almost always (width, height)
    MaxY As Long
    colorDepth As Long      'The colorDepth of the current layer; right now, this should always be 24 or 32
    BytesPerPixel As Long   'BPP is colorDepth / 8.  It is provided for convenience.
    LayerX As Long          'Filters shouldn't have to worry about where the layer is physically located, but when it comes
    LayerY As Long          ' time to set the layer back in place, these may be useful (as when previewing, for example)
End Type

Public curLayerValues As FilterInfo

'We may need a temporary copy of the selection mask; if so, it will be stored here
Dim tmpSelectionMask As pdLayer

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


'prepImageData's job is to copy the relevant layer into a temporary object, which is what individual filters and effects
' will operate on.  prepImageData() also populates the relevant SafeArray object and a host of other variables, which
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
    
    'Prepare a reference to a picture box (only needed when previewing)
    Dim previewPictureBox As PictureBox
    
    'If this is a preview, we need to calculate new width and height for the image that will appear in the preview window
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
    
        Set previewPictureBox = previewTarget.getPreviewPic
        dstWidth = previewPictureBox.ScaleWidth
        dstHeight = previewPictureBox.ScaleHeight
            
        'The source values need to be adjusted contingent on whether this is a selection or a full-image preview
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
        
        'And finally, create our workingLayer using these values
        If pdImages(CurrentImage).selectionActive Then
            Dim copyLayer As pdLayer
            Set copyLayer = New pdLayer
            copyLayer.createBlank pdImages(CurrentImage).mainSelection.boundWidth, pdImages(CurrentImage).mainSelection.boundHeight, pdImages(CurrentImage).mainLayer.getLayerColorDepth
            BitBlt copyLayer.getLayerDC, 0, 0, pdImages(CurrentImage).mainSelection.boundWidth, pdImages(CurrentImage).mainSelection.boundHeight, pdImages(CurrentImage).mainLayer.getLayerDC, pdImages(CurrentImage).mainSelection.boundLeft, pdImages(CurrentImage).mainSelection.boundTop, vbSrcCopy
            workingLayer.createFromExistingLayer copyLayer, newWidth, newHeight
            copyLayer.eraseLayer
            Set copyLayer = Nothing
            
            'Also, make a copy of the selection mask at the same dimensions as the preview.  We will use this to remove the sections of the
            ' selection that are not selected.  (Say that 10 times fast...lol)
            Dim tmpSelectionMaskCopy As pdLayer
            Set tmpSelectionMaskCopy = New pdLayer
            tmpSelectionMaskCopy.createBlank pdImages(CurrentImage).mainSelection.boundWidth, pdImages(CurrentImage).mainSelection.boundHeight, 24
            BitBlt tmpSelectionMaskCopy.getLayerDC, 0, 0, pdImages(CurrentImage).mainSelection.boundWidth, pdImages(CurrentImage).mainSelection.boundHeight, pdImages(CurrentImage).mainSelection.selMask.getLayerDC, pdImages(CurrentImage).mainSelection.boundLeft, pdImages(CurrentImage).mainSelection.boundTop, vbSrcCopy
            Set tmpSelectionMask = New pdLayer
            tmpSelectionMask.createFromExistingLayer tmpSelectionMaskCopy, newWidth, newHeight
            tmpSelectionMaskCopy.eraseLayer
            Set tmpSelectionMaskCopy = Nothing
            
            'In the future, we will first check to see if this selection has a complex area, e.g. if it is not a
            ' square or a rectangle.  If it IS complex, we need to modify it further to remove inactive pixels.
            'If pdImages(CurrentImage).mainSelection.isComplicated (or something like this) ...
            
            'For now, however, all selections are fully processed.
            
            'Start by converting the working layer to 32bpp.  Unselected areas must be made transparent, so 24bpp won't work.
            Dim already32bpp As Boolean
            
            If workingLayer.getLayerColorDepth = 32 Then
                already32bpp = True
            Else
                already32bpp = False
                workingLayer.convertTo32bpp
            End If
            
            'Next, we are going to remove any pixels that are not part of the selection mask.
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
            
            'With our work complete, point both ImageData() arrays away from their DIBs and deallocate them
            CopyMemory ByVal VarPtrArray(wlImageData), 0&, 4
            Erase wlImageData
            CopyMemory ByVal VarPtrArray(selImageData), 0&, 4
            Erase selImageData
            
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
        .colorDepth = workingLayer.getLayerColorDepth
        .BytesPerPixel = (workingLayer.getLayerColorDepth \ 8)
        .LayerX = 0
        .LayerY = 0
    End With

    'Set up the progress bar (only if this is NOT a preview, mind you)
    If Not isPreview Then
        If newProgBarMax = -1 Then
            SetProgBarMax (curLayerValues.Left + curLayerValues.Width)
        Else
            SetProgBarMax newProgBarMax
        End If
    End If
    
    'pdMsgBox "prepImageData worked: " & workingLayer.getLayerHeight & ", " & workingLayer.getLayerWidth & " (" & workingLayer.getLayerArrayWidth & ")" & ", " & workingLayer.getLayerDIBits

End Sub


'The counterpart to prepImageData, finalizeImageData copies the working layer back into its source then renders it
' to the screen.  Like prepImageData(), a preview target can also be named.  In this case, finalizeImageData will rely on
' the values calculated by prepImageData(), as it's presumed that preImageData will ALWAYS be called before this routine.
Public Sub finalizeImageData(Optional isPreview As Boolean = False, Optional previewTarget As fxPreviewCtl)

    'If the user has canceled the current action, disregard the working layer and exit immediately
    If cancelCurrentAction Then
        SetProgBarVal 0
        Message "Action canceled."
        
        workingLayer.eraseLayer
        Set workingLayer = Nothing
                
        cancelCurrentAction = False
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
                            dstImageData((leftOffset + x) * dstQuickVal + i, topOffset + y) = BlendColors(dstImageData((leftOffset + x) * dstQuickVal + i, topOffset + y), wlImageData(x * 4 + i, y), blendAlpha)
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
    
        'Allow workingLayer to paint itself to the target picture box
        workingLayer.renderToPictureBox previewTarget.getPreviewPic
        
        'Give the preview object a copy of the layer data used to generate the preview
        previewTarget.setFXImage workingLayer
        
        'workingLayer has served its purpose, so erase it from memory
        workingLayer.eraseLayer
        Set workingLayer = Nothing
        
    End If
    
End Sub

