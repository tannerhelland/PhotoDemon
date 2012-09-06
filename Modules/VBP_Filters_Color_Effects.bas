Attribute VB_Name = "Filters_Color_Effects"
'***************************************************************************
'Filter (Color Effects) Interface
'Copyright ©2000-2012 by Tanner Helland
'Created: 25/January/03
'Last updated: 14/August/12
'Last update: improved comments and organization
'
'Runs all color-related filters (invert, negative, shifting, etc.).
'
'***************************************************************************

Option Explicit

'Invert an image
Public Sub MenuInvert()
        
    Message "Inverting the image..."
    
    'Create a local array and point it at the pixel data we want to operate on
    Dim ImageData() As Byte
    Dim tmpSA As SAFEARRAY2D
    prepImageData tmpSA
    CopyMemory ByVal VarPtrArray(ImageData()), VarPtr(tmpSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curLayerValues.Left
    initY = curLayerValues.Top
    finalX = curLayerValues.Right
    finalY = curLayerValues.Bottom
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long
    qvDepth = curLayerValues.BytesPerPixel
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = findBestProgBarValue()
    
    'After all that work, the Invert code itself is very small and unexciting!
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
        ImageData(QuickVal, y) = 255 Xor ImageData(QuickVal, y)
        ImageData(QuickVal + 1, y) = 255 Xor ImageData(QuickVal + 1, y)
        ImageData(QuickVal + 2, y) = 255 Xor ImageData(QuickVal + 2, y)
    Next y
        If (x And progBarCheck) = 0 Then SetProgBarVal x
    Next x
        
    'With our work complete, point ImageData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(ImageData), 0&, 4
    Erase ImageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    finalizeImageData
    
End Sub

'Rechannel an image (red, green, or blue)
Public Sub MenuRechannel(ByVal RType As Byte)
    
    Message "Rechanneling colors..."
    SetProgBarMax PicWidthL
    
    GetImageData
    Dim QuickVal As Long
    
    For x = 0 To PicWidthL
        QuickVal = x * 3
    For y = 0 To PicHeightL
        'Rechannel Red
        If RType = 0 Then
            ImageData(QuickVal, y) = 0
            ImageData(QuickVal + 1, y) = 0
        'Rechannel Green
        ElseIf RType = 1 Then
            ImageData(QuickVal, y) = 0
            ImageData(QuickVal + 2, y) = 0
        'Rechannel Blue
        Else
            ImageData(QuickVal + 1, y) = 0
            ImageData(QuickVal + 2, y) = 0
        End If
    Next y
        If x Mod 20 = 0 Then SetProgBarVal x
    Next x
    
    setImageData
    
End Sub

'Shift colors (right or left)
Public Sub MenuCShift(ByVal SType As Byte)
    
    If SType = 0 Then
        Message "Shifting RGB values right..."
    Else
        Message "Shifting RGB values left..."
    End If
    SetProgBarMax PicWidthL
    
    Dim tR As Byte, tB As Byte, tG As Byte
    
    GetImageData
    Dim QuickVal As Long
    
    For x = 0 To PicWidthL
        QuickVal = x * 3
    For y = 0 To PicHeightL
        If SType = 0 Then
            tR = ImageData(QuickVal, y)
            tG = ImageData(QuickVal + 2, y)
            tB = ImageData(QuickVal + 1, y)
        Else
            tR = ImageData(QuickVal + 1, y)
            tG = ImageData(QuickVal, y)
            tB = ImageData(QuickVal + 2, y)
        End If
        ImageData(QuickVal + 2, y) = tR
        ImageData(QuickVal + 1, y) = tG
        ImageData(QuickVal, y) = tB
    Next y
        If x Mod 20 = 0 Then SetProgBarVal x
    Next x
    
    setImageData
    
End Sub

'Generate a luminance-negative version of the image
Public Sub MenuNegative()

    Message "Generating film negative..."
    SetProgBarMax PicWidthL
    
    GetImageData
    Dim QuickVal As Long
    
    Dim r As Long, g As Long, b As Long
    Dim HH As Single, SS As Single, LL As Single
    
    For x = 0 To PicWidthL
        QuickVal = x * 3
    For y = 0 To PicHeightL
    
        'Get the original RGB values
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        
        'Use those to calculate hue and saturation
        tRGBToHSL r, g, b, HH, SS, LL
        
        'Convert back to RGB using inverted luminance
        tHSLToRGB HH, SS, 1 - LL, r, g, b
        
        'Assign those values into the array
        ImageData(QuickVal + 2, y) = r
        ImageData(QuickVal + 1, y) = g
        ImageData(QuickVal, y) = b
        
    Next y
        If x Mod 20 = 0 Then SetProgBarVal x
    Next x
    
    setImageData
    
End Sub

'Invert the hue of an image
Public Sub MenuInvertHue()
    
    Message "Inverting image hue..."
    SetProgBarMax PicWidthL
    
    GetImageData
    Dim QuickVal As Long
    
    Dim r As Long, g As Long, b As Long
    Dim tR As Long, tB As Long, tG As Long
    Dim HH As Single, SS As Single, LL As Single
    
    For x = 0 To PicWidthL
        QuickVal = x * 3
    For y = 0 To PicHeightL
    
        'Get the original RGB values
        tR = ImageData(QuickVal + 2, y)
        tG = ImageData(QuickVal + 1, y)
        tB = ImageData(QuickVal, y)
        
        'Use those to calculate hue, saturation, and luminance
        tRGBToHSL tR, tG, tB, HH, SS, LL
        
        'Invert hue
        HH = 6 - (HH + 1) - 1
        
        'Convert the newly calculated HSL values back to RGB
        tHSLToRGB HH, SS, LL, r, g, b
        
        ImageData(QuickVal + 2, y) = r
        ImageData(QuickVal + 1, y) = g
        ImageData(QuickVal, y) = b
    Next y
        If x Mod 20 = 0 Then SetProgBarVal x
    Next x
    
    setImageData
    
End Sub

'I call this effect a "compound invert", but it's not actually anything like an invert operation.  Oh well.
Public Sub MenuCompoundInvert(ByVal Divisor As Long)

    Message "Performing compound inversion..."
    
    'Create a local array and point it at the pixel data we want to operate on
    Dim ImageData() As Byte
    Dim tmpSA As SAFEARRAY2D
    prepImageData tmpSA
    CopyMemory ByVal VarPtrArray(ImageData()), VarPtr(tmpSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curLayerValues.Left
    initY = curLayerValues.Top
    finalX = curLayerValues.Right
    finalY = curLayerValues.Bottom
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long
    qvDepth = curLayerValues.BytesPerPixel
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = findBestProgBarValue()
    
    'Finally, a bunch of variables used in color calculation
    Dim r As Long, g As Long, b As Long
    Dim newR As Long, newG As Long, newB As Long
        
    'Apply the filter
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
        
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        
        If r = 0 Then r = 1
        If g = 0 Then g = 1
        If b = 0 Then b = 1
        
        newR = (g * b) \ Divisor
        newG = (r * b) \ Divisor
        newB = (r * g) \ Divisor
        
        If newR > 255 Then newR = 255
        If newG > 255 Then newG = 255
        If newB > 255 Then newB = 255
        
        ImageData(QuickVal + 2, y) = newR
        ImageData(QuickVal + 1, y) = newG
        ImageData(QuickVal, y) = newB
        
    Next y
        If (x And progBarCheck) = 0 Then SetProgBarVal x
    Next x
        
    'With our work complete, point ImageData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(ImageData), 0&, 4
    Erase ImageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    finalizeImageData

End Sub

'Enhance CONTRAST
Public Sub MenuAutoEnhanceContrast()

    Message "Enhancing image contrast..."
    SetProgBarMax PicWidthL
    
    GetImageData
    Dim QuickVal As Long
    
    Dim r As Long, g As Long, b As Long
    Dim tR As Long, tG As Long, tB As Long, TC As Long

    For x = 0 To PicWidthL
        QuickVal = x * 3
    For y = 0 To PicHeightL
    
        'Get the original RGB values
        tR = ImageData(QuickVal + 2, y)
        tG = ImageData(QuickVal + 1, y)
        tB = ImageData(QuickVal, y)
        
        'Calculate grayscale
        TC = Int((222 * tR + 707 * tG + 71 * tB) \ 1000)
        
        'Spread out the contrast
        r = Abs(tR - TC) + tR
        g = Abs(tG - TC) + tG
        b = Abs(tB - TC) + tB
        
        ByteMeL r
        ByteMeL g
        ByteMeL b
        
        ImageData(QuickVal + 2, y) = r
        ImageData(QuickVal + 1, y) = g
        ImageData(QuickVal, y) = b
    Next y
        If x Mod 20 = 0 Then SetProgBarVal x
    Next x
    
    setImageData
    
End Sub

'Enhance HIGHLIGHTS
Public Sub MenuAutoEnhanceHighlights()
    
    Message "Enhancing image highlights..."
    SetProgBarMax PicWidthL
    
    GetImageData
    Dim QuickVal As Long
    
    Dim r As Long, g As Long, b As Long
    Dim tR As Long, tG As Long, tB As Long, TC As Long
    
    For x = 0 To PicWidthL
        QuickVal = x * 3
    For y = 0 To PicHeightL
    
        'Get the RGB values
        tR = ImageData(QuickVal + 2, y)
        tG = ImageData(QuickVal + 1, y)
        tB = ImageData(QuickVal, y)
        
        'Calculate grayscale
        TC = Int((222 * tR + 707 * tG + 71 * tB) \ 1000)
        
        'Spread out highlights
        r = Abs(tR - TC) + TC
        g = Abs(tG - TC) + TC
        b = Abs(tB - TC) + TC
        
        ByteMeL r
        ByteMeL g
        ByteMeL b
        
        ImageData(QuickVal + 2, y) = r
        ImageData(QuickVal + 1, y) = g
        ImageData(QuickVal, y) = b
    Next y
        If x Mod 20 = 0 Then SetProgBarVal x
    Next x
    
    setImageData
    
End Sub

'Enhance MIDTONES
Public Sub MenuAutoEnhanceMidtones()

    Message "Enhancing image midtones..."
    SetProgBarMax PicWidthL
    
    GetImageData
    Dim QuickVal As Long
    
    Dim r As Long, g As Long, b As Long
    Dim tR As Long, tG As Long, tB As Long, TC As Long
    
    For x = 0 To PicWidthL
        QuickVal = x * 3
    For y = 0 To PicHeightL
        
        'Get RGB values
        tR = ImageData(QuickVal + 2, y)
        tG = ImageData(QuickVal + 1, y)
        tB = ImageData(QuickVal, y)
        
        'Calculate grayscale
        TC = Int((222 * tR + 707 * tG + 71 * tB) \ 1000)
        
        'Spread out midtones
        r = tR - (TC - tR)
        g = tG - (TC - tG)
        b = tB - (TC - tB)
        
        ByteMeL r
        ByteMeL g
        ByteMeL b
        
        ImageData(QuickVal + 2, y) = r
        ImageData(QuickVal + 1, y) = g
        ImageData(QuickVal, y) = b
    Next y
        If x Mod 20 = 0 Then SetProgBarVal x
    Next x
    
    setImageData
    
End Sub

'Enhance SHADOWS
Public Sub MenuAutoEnhanceShadows()

    Message "Enhancing image shadows..."
    SetProgBarMax PicWidthL
    
    GetImageData
    Dim QuickVal As Long
    
    Dim r As Long, g As Long, b As Long
    Dim tR As Long, tG As Long, tB As Long, TC As Long

    For x = 0 To PicWidthL
        QuickVal = x * 3
    For y = 0 To PicHeightL
    
        'Get RGB values
        tR = ImageData(QuickVal + 2, y)
        tG = ImageData(QuickVal + 1, y)
        tB = ImageData(QuickVal, y)
        
        'Calculate grayscale
        TC = Int((222 * tR + 707 * tG + 71 * tB) \ 1000)
        
        'Spread out shadows
        r = tR
        g = tG + Abs(r - TC)
        b = tB + Abs(g - TC)
        r = tR + Abs(b - TC)
        
        ByteMeL r
        ByteMeL g
        ByteMeL b
        
        ImageData(QuickVal + 2, y) = r
        ImageData(QuickVal + 1, y) = g
        ImageData(QuickVal, y) = b
    Next y
        If x Mod 20 = 0 Then SetProgBarVal x
    Next x
    
    setImageData
    
End Sub

