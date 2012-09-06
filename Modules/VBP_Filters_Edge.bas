Attribute VB_Name = "Filters_Edge"
'***************************************************************************
'Filter (Edge) Interface
'Copyright ©2000-2012 by Tanner Helland
'Created: 12/June/01
'Last updated: 05/September/12
'Last update: rewrote and optimized all filters against the new layer class.
'
'Runs all edge-related filters (edge detection, relief, etc.).
'
'***************************************************************************

Option Explicit

'Redraw the image using a pencil sketch effect.
Public Sub FilterPencil()

    Message "Sketching the image with an imaginary pencil..."
    
    'Create a local array and point it at the pixel data we want to operate on
    Dim ImageData() As Byte
    Dim tmpSA As SAFEARRAY2D
    prepImageData tmpSA
    CopyMemory ByVal VarPtrArray(ImageData()), VarPtr(tmpSA), 4
        
    'We also need a second array that will hold an identical copy of the image.  Use CopyMemory to obtain this quickly.
    Dim srcImageData() As Byte
    ReDim srcImageData(0 To pdImages(CurrentImage).mainLayer.getLayerArrayWidth - 1, 0 To pdImages(CurrentImage).Height - 1) As Byte
    CopyMemory srcImageData(0, 0), ImageData(0, 0), pdImages(CurrentImage).mainLayer.getLayerArrayWidth * pdImages(CurrentImage).Height
    
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    Dim c As Long, d As Long
    initX = curLayerValues.Left
    initY = curLayerValues.Top
    finalX = curLayerValues.Right
    finalY = curLayerValues.Bottom
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, inQuickVal As Long, qvDepth As Long
    qvDepth = curLayerValues.BytesPerPixel
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = findBestProgBarValue()
    
    'Finally, a bunch of variables used in color calculation
    Dim r As Long, g As Long, b As Long
    Dim grayVal As Long
    
    'Because gray values are constant, we can use a look-up table to calculate them
    Dim gLookup(0 To 765) As Byte
    For x = 0 To 765
        gLookup(x) = CByte(x \ 3)
    Next x
        
    'Apply the filter
    For x = initX + 1 To finalX - 1
        QuickVal = x * qvDepth
    For y = initY + 1 To finalY - 1
        
        r = 0
        g = 0
        b = 0
        
        For c = x - 1 To x + 1
            inQuickVal = c * qvDepth
        For d = y - 1 To y + 1
        
            If c = x And d = y Then
                r = r + 8 * srcImageData(inQuickVal + 2, d)
                g = g + 8 * srcImageData(inQuickVal + 1, d)
                b = b + 8 * srcImageData(inQuickVal, d)
            Else
                r = r - srcImageData(inQuickVal + 2, d)
                g = g - srcImageData(inQuickVal + 1, d)
                b = b - srcImageData(inQuickVal, d)
            End If
            
        Next d
        Next c
        
        If r > 255 Then r = 255
        If g > 255 Then g = 255
        If b > 255 Then b = 255
        If r < 0 Then r = 0
        If g < 0 Then g = 0
        If b < 0 Then b = 0
        
        grayVal = 255 - gLookup(r + g + b)
        
        ImageData(QuickVal + 2, y) = grayVal
        ImageData(QuickVal + 1, y) = grayVal
        ImageData(QuickVal, y) = grayVal
        
    Next y
        If (x And progBarCheck) = 0 Then SetProgBarVal x
    Next x
        
    'With our work complete, point ImageData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(ImageData), 0&, 4
    Erase ImageData
    
    'Erase our temporary array as well
    Erase srcImageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    finalizeImageData

End Sub

'A typical relief filter, that makes the image seem pseudo-3D.
Public Sub FilterRelief()
    FilterSize = 3
    ReDim FM(-1 To 1, -1 To 1) As Long
    FM(-1, -1) = 2
    FM(-1, 0) = 1
    FM(0, 1) = 1
    FM(0, 0) = 1
    FM(0, -1) = -1
    FM(1, 0) = -1
    FM(1, 1) = -2
    FilterWeight = 3
    FilterBias = 75
    DoFilter "Relief"
End Sub

'A lighter version of a traditional sharpen filter; it's designed to bring out edge detail without the blowout typical of sharpening
Public Sub FilterEdgeEnhance()
    FilterSize = 3
    ReDim FM(-1 To 1, -1 To 1) As Long
    FM(-1, 0) = -1
    FM(1, 0) = -1
    FM(0, -1) = -1
    FM(0, 1) = -1
    FM(0, 0) = 8
    FilterWeight = 4
    FilterBias = 0
    DoFilter "Edge Enhance"
End Sub
