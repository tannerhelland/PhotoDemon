Attribute VB_Name = "Filters_Rotate"
'***************************************************************************
'Filter (Rotation) Interface
'Copyright ©2000-2012 by Tanner Helland
'Created: 25/January/03
'Last updated: 05/September/12
'Last update: Rewrote all rotation code against the new layer system.
'
'Runs all rotation-style transformations.  Includes flip and mirror as well.
'
'***************************************************************************

Option Explicit

'Flip an image vertically
Public Sub MenuFlip()

    Message "Flipping image..."
    
    StretchBlt pdImages(CurrentImage).mainLayer.getLayerDC, 0, 0, pdImages(CurrentImage).Width, pdImages(CurrentImage).Height, pdImages(CurrentImage).mainLayer.getLayerDC, 0, pdImages(CurrentImage).Height - 1, pdImages(CurrentImage).Width, -pdImages(CurrentImage).Height, vbSrcCopy
        
    Message "Finished. "
    
    ScrollViewport FormMain.ActiveForm
    
End Sub

'Flip an image horizontally
Public Sub MenuMirror()

    Message "Mirroring image..."
    
    StretchBlt pdImages(CurrentImage).mainLayer.getLayerDC, 0, 0, pdImages(CurrentImage).Width, pdImages(CurrentImage).Height, pdImages(CurrentImage).mainLayer.getLayerDC, pdImages(CurrentImage).Width - 1, 0, -pdImages(CurrentImage).Width, pdImages(CurrentImage).Height, vbSrcCopy
    
    Message "Finished. "
    
    ScrollViewport FormMain.ActiveForm
    
End Sub

'Rotate an image 90° clockwise
Public Sub MenuRotate90Clockwise()

    'If a selection is active, remove it.  (This is not the most elegant solution - the elegant solution would be rotating
    ' the selection to match the new image, but we can fix that at a later date.)
    If pdImages(CurrentImage).selectionActive Then
        pdImages(CurrentImage).selectionActive = False
        tInit tSelection, False
    End If

    Message "Rotating image clockwise..."
    
    'Create a local array and point it at the pixel data of the current image
    Dim srcImageData() As Byte
    Dim srcSA As SAFEARRAY2D
    prepImageData srcSA
    CopyMemory ByVal VarPtrArray(srcImageData()), VarPtr(srcSA), 4
    
    'Create a second local array.  This will contain the pixel data of the new, rotated image
    Dim dstImageData() As Byte
    Dim dstSA As SAFEARRAY2D
    
    Dim dstLayer As pdLayer
    Set dstLayer = New pdLayer
    dstLayer.createBlank pdImages(CurrentImage).mainLayer.getLayerHeight, pdImages(CurrentImage).mainLayer.getLayerWidth, pdImages(CurrentImage).mainLayer.getLayerColorDepth
    
    prepSafeArray dstSA, dstLayer
    CopyMemory ByVal VarPtrArray(dstImageData()), VarPtr(dstSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, i As Long
    Dim initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curLayerValues.Left
    initY = curLayerValues.Top
    finalX = curLayerValues.Right
    finalY = curLayerValues.Bottom
    
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long, QuickValY
    qvDepth = curLayerValues.BytesPerPixel
    
    Dim iWidth As Long, iHeight As Long
    iWidth = finalX * qvDepth
    iHeight = finalY * qvDepth
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = findBestProgBarValue()
        
    'Rotate the source image into the destination image, using the arrays provided
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
        QuickValY = y * qvDepth
        
        For i = 0 To qvDepth - 1
            dstImageData(iHeight - QuickValY + i, finalX - x) = srcImageData(iWidth - QuickVal + i, y)
        Next i
        
    Next y
        If (x And progBarCheck) = 0 Then SetProgBarVal x
    Next x
    
    'With our work complete, point both ImageData() arrays away from their respective DIBs and deallocate them
    CopyMemory ByVal VarPtrArray(srcImageData), 0&, 4
    Erase srcImageData
    CopyMemory ByVal VarPtrArray(dstImageData), 0&, 4
    Erase dstImageData
    
    'dstImageData now contains the rotated image.  We need to transfer that back into the current image.
    pdImages(CurrentImage).mainLayer.createFromExistingLayer dstLayer
    
    'With that transfer complete, we can erase our temporary layer
    dstLayer.eraseLayer
    Set dstLayer = Nothing
    
    'Update the current image size
    pdImages(CurrentImage).updateSize
    DisplaySize pdImages(CurrentImage).Width, pdImages(CurrentImage).Height
    
    Message "Finished. "
    
    'Redraw the image
    FitWindowToImage
    
    'Reset the progress bar to zero
    SetProgBarVal 0
    
End Sub

'Rotate an image 180°
Public Sub MenuRotate180()

    'Fun fact: rotating 180 degrees can be accomplished by flipping and then mirroring it.
    Message "Rotating image..."
        
    StretchBlt pdImages(CurrentImage).mainLayer.getLayerDC, 0, 0, pdImages(CurrentImage).Width, pdImages(CurrentImage).Height, pdImages(CurrentImage).mainLayer.getLayerDC, pdImages(CurrentImage).Width - 1, pdImages(CurrentImage).Height - 1, -pdImages(CurrentImage).Width, -pdImages(CurrentImage).Height, vbSrcCopy
        
    Message "Finished. "
    
    ScrollViewport FormMain.ActiveForm
    
End Sub

'Rotate an image 90° counter-clockwise
Public Sub MenuRotate270Clockwise()

    'If a selection is active, remove it.  (This is not the most elegant solution - the elegant solution would be rotating
    ' the selection to match the new image, but we can fix that at a later date.)
    If pdImages(CurrentImage).selectionActive Then
        pdImages(CurrentImage).selectionActive = False
        tInit tSelection, False
    End If

    Message "Rotating image counter-clockwise..."
    
    'Create a local array and point it at the pixel data of the current image
    Dim srcImageData() As Byte
    Dim srcSA As SAFEARRAY2D
    prepImageData srcSA
    CopyMemory ByVal VarPtrArray(srcImageData()), VarPtr(srcSA), 4
    
    'Create a second local array.  This will contain the pixel data of the new, rotated image
    Dim dstImageData() As Byte
    Dim dstSA As SAFEARRAY2D
    
    Dim dstLayer As pdLayer
    Set dstLayer = New pdLayer
    dstLayer.createBlank pdImages(CurrentImage).mainLayer.getLayerHeight, pdImages(CurrentImage).mainLayer.getLayerWidth, pdImages(CurrentImage).mainLayer.getLayerColorDepth
    
    prepSafeArray dstSA, dstLayer
    CopyMemory ByVal VarPtrArray(dstImageData()), VarPtr(dstSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, i As Long
    Dim initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curLayerValues.Left
    initY = curLayerValues.Top
    finalX = curLayerValues.Right
    finalY = curLayerValues.Bottom
    
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long, QuickValY
    qvDepth = curLayerValues.BytesPerPixel
    
    Dim iWidth As Long
    iWidth = finalX * qvDepth
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = findBestProgBarValue()
        
    'Rotate the source image into the destination image, using the arrays provided
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
        QuickValY = y * qvDepth
        
        For i = 0 To qvDepth - 1
            dstImageData(QuickValY + i, x) = srcImageData(iWidth - QuickVal + i, y)
        Next i
        
    Next y
        If (x And progBarCheck) = 0 Then SetProgBarVal x
    Next x
    
    'With our work complete, point both ImageData() arrays away from their respective DIBs and deallocate them
    CopyMemory ByVal VarPtrArray(srcImageData), 0&, 4
    Erase srcImageData
    CopyMemory ByVal VarPtrArray(dstImageData), 0&, 4
    Erase dstImageData
    
    'dstImageData now contains the rotated image.  We need to transfer that back into the current image.
    pdImages(CurrentImage).mainLayer.createFromExistingLayer dstLayer
    
    'With that transfer complete, we can erase our temporary layer
    dstLayer.eraseLayer
    Set dstLayer = Nothing
    
    'Update the current image size
    pdImages(CurrentImage).updateSize
    DisplaySize pdImages(CurrentImage).Width, pdImages(CurrentImage).Height
    
    Message "Finished. "
    
    'Redraw the image
    FitWindowToImage
    
    'Reset the progress bar to zero
    SetProgBarVal 0
    
End Sub
