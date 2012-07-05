Attribute VB_Name = "Filters_Color_Effects"
'***************************************************************************
'Filter (Color Effects) Interface
'©2000-2012 Tanner Helland
'Created: 1/25/03
'Last updated: 11/June/12
'Last update: removed all image-stream related code.
'
'Runs all color-related filters (invert, negative, shifting, etc.).
'
'***************************************************************************

Option Explicit

'Invert an image
Public Sub MenuInvert()
        
    Message "Inverting the image..."
    Dim QuickVal As Long

    SetProgBarMax PicWidthL
    GetImageData
    For X = 0 To PicWidthL
        QuickVal = X * 3
    For Y = 0 To PicHeightL
        ImageData(QuickVal, Y) = 255 Xor ImageData(QuickVal, Y)
        ImageData(QuickVal + 1, Y) = 255 Xor ImageData(QuickVal + 1, Y)
        ImageData(QuickVal + 2, Y) = 255 Xor ImageData(QuickVal + 2, Y)
    Next Y
        If X Mod 10 = 0 Then SetProgBarVal X
    Next X
    
    SetImageData

End Sub

'Rechannel an image (red, green, or blue)
Public Sub MenuRechannel(ByVal RType As Byte)
    Dim QuickVal As Long
    Message "Rechanneling colors..."
    SetProgBarMax PicWidthL
    For X = 0 To PicWidthL
        QuickVal = X * 3
    For Y = 0 To PicHeightL
        If RType = 0 Then
            ImageData(QuickVal, Y) = 0
            ImageData(QuickVal + 1, Y) = 0
        ElseIf RType = 1 Then
            ImageData(QuickVal, Y) = 0
            ImageData(QuickVal + 2, Y) = 0
        Else
            ImageData(QuickVal + 1, Y) = 0
            ImageData(QuickVal + 2, Y) = 0
        End If
    Next Y
        If X Mod 20 = 0 Then SetProgBarVal X
    Next X
    SetImageData
End Sub

'Shift colors (right or left)
Public Sub MenuCShift(ByVal SType As Byte)
    Dim tR As Byte, tB As Byte, tG As Byte
    If SType = 0 Then
        Message "Shifting RGB values right..."
    Else
        Message "Shifting RGB values left..."
    End If
    SetProgBarMax PicWidthL
    Dim QuickVal As Long
    For X = 0 To PicWidthL
        QuickVal = X * 3
    For Y = 0 To PicHeightL
        If SType = 0 Then
            tR = ImageData(QuickVal, Y)
            tG = ImageData(QuickVal + 2, Y)
            tB = ImageData(QuickVal + 1, Y)
        Else
            tR = ImageData(QuickVal + 1, Y)
            tG = ImageData(QuickVal, Y)
            tB = ImageData(QuickVal + 2, Y)
        End If
        ImageData(QuickVal + 2, Y) = tR
        ImageData(QuickVal + 1, Y) = tG
        ImageData(QuickVal, Y) = tB
    Next Y
        If X Mod 20 = 0 Then SetProgBarVal X
    Next X
    SetImageData
End Sub

'Generate a luminance-negative version of the image
Public Sub MenuNegative()
    GetImageData
    Dim r As Long, G As Long, B As Long
    Dim HH As Single, SS As Single, LL As Single
    Message "Generating film negative..."
    SetProgBarMax PicWidthL
    SetProgBarVal 0
    Dim QuickVal As Long
    For X = 0 To PicWidthL
        QuickVal = X * 3
    For Y = 0 To PicHeightL
        'Get the temporary values
        r = ImageData(QuickVal + 2, Y)
        G = ImageData(QuickVal + 1, Y)
        B = ImageData(QuickVal, Y)
        'Get the hue and saturation
        tRGBToHSL r, G, B, HH, SS, LL
        'Convert back to RGB using our artificial luminance values
        tHSLToRGB HH, SS, 1 - LL, r, G, B
        'Assign those values into the array
        ImageData(QuickVal + 2, Y) = r
        ImageData(QuickVal + 1, Y) = G
        ImageData(QuickVal, Y) = B
    Next Y
        If X Mod 20 = 0 Then SetProgBarVal X
    Next X
    SetImageData
    Message "Finished."
End Sub

'Invert the hue of an image
Public Sub MenuInvertHue()
    Dim r As Long, G As Long, B As Long
    Dim tR As Long, tB As Long, tG As Long
    Dim HH As Single, SS As Single, LL As Single
    Message "Inverting image hue..."
    SetProgBarMax PicWidthL
    Dim QuickVal As Long
    For X = 0 To PicWidthL
        QuickVal = X * 3
    For Y = 0 To PicHeightL
        tR = ImageData(QuickVal + 2, Y)
        tG = ImageData(QuickVal + 1, Y)
        tB = ImageData(QuickVal, Y)
        
        'Get the hue and saturation
        tRGBToHSL tR, tG, tB, HH, SS, LL
        HH = 6 - (HH + 1) - 1
        'Convert back to RGB using our artificial hue value
        tHSLToRGB HH, SS, LL, r, G, B
        
        ImageData(QuickVal + 2, Y) = r
        ImageData(QuickVal + 1, Y) = G
        ImageData(QuickVal, Y) = B
    Next Y
        If X Mod 20 = 0 Then SetProgBarVal X
    Next X
    SetImageData
End Sub

'Enhance CONTRAST
Public Sub MenuAutoEnhanceContrast()
    GetImageData
    Dim r As Long, G As Long, B As Long
    Dim tR As Long, tG As Long, tB As Long, TC As Long
    Message "Enhancing image contrast..."
    SetProgBarMax PicWidthL
    SetProgBarVal 0
    Dim QuickVal As Long
    For X = 0 To PicWidthL
        QuickVal = X * 3
    For Y = 0 To PicHeightL
        'Get the temporary values
        tR = ImageData(QuickVal + 2, Y)
        tG = ImageData(QuickVal + 1, Y)
        tB = ImageData(QuickVal, Y)
        TC = Int((222 * tR + 707 * tG + 71 * tB) \ 1000)
        r = Abs(tR - TC) + tR
        G = Abs(tG - TC) + tG
        B = Abs(tB - TC) + tB
        ByteMeL r
        ByteMeL G
        ByteMeL B
        'Assign those values into the array
        ImageData(QuickVal + 2, Y) = r
        ImageData(QuickVal + 1, Y) = G
        ImageData(QuickVal, Y) = B
    Next Y
        If X Mod 20 = 0 Then SetProgBarVal X
    Next X
    SetImageData
End Sub

'Enhance HIGHLIGHTS
Public Sub MenuAutoEnhanceHighlights()
    GetImageData
    Dim r As Long, G As Long, B As Long
    Dim tR As Long, tG As Long, tB As Long, TC As Long
    Message "Enhancing image highlights..."
    SetProgBarMax PicWidthL
    SetProgBarVal 0
    Dim QuickVal As Long
    For X = 0 To PicWidthL
        QuickVal = X * 3
    For Y = 0 To PicHeightL
        'Get the temporary values
        tR = ImageData(QuickVal + 2, Y)
        tG = ImageData(QuickVal + 1, Y)
        tB = ImageData(QuickVal, Y)
        TC = Int((222 * tR + 707 * tG + 71 * tB) \ 1000)
        r = Abs(tR - TC) + TC
        G = Abs(tG - TC) + TC
        B = Abs(tB - TC) + TC
        ByteMeL r
        ByteMeL G
        ByteMeL B
        'Assign those values into the array
        ImageData(QuickVal + 2, Y) = r
        ImageData(QuickVal + 1, Y) = G
        ImageData(QuickVal, Y) = B
    Next Y
        If X Mod 20 = 0 Then SetProgBarVal X
    Next X
    SetImageData
End Sub

'Enhance MIDTONES
Public Sub MenuAutoEnhanceMidtones()
    GetImageData
    Dim r As Long, G As Long, B As Long
    Dim tR As Long, tG As Long, tB As Long, TC As Long
    Message "Enhancing image midtones..."
    SetProgBarMax PicWidthL
    SetProgBarVal 0
    Dim QuickVal As Long
    For X = 0 To PicWidthL
        QuickVal = X * 3
    For Y = 0 To PicHeightL
        'Get the temporary values
        tR = ImageData(QuickVal + 2, Y)
        tG = ImageData(QuickVal + 1, Y)
        tB = ImageData(QuickVal, Y)
        TC = Int((222 * tR + 707 * tG + 71 * tB) \ 1000)
        r = tR - (TC - tR)
        G = tG - (TC - tG)
        B = tB - (TC - tB)
        ByteMeL r
        ByteMeL G
        ByteMeL B
        'Assign those values into the array
        ImageData(QuickVal + 2, Y) = r
        ImageData(QuickVal + 1, Y) = G
        ImageData(QuickVal, Y) = B
    Next Y
        If X Mod 20 = 0 Then SetProgBarVal X
    Next X
    SetImageData
End Sub

'Enhance SHADOWS
Public Sub MenuAutoEnhanceShadows()
    GetImageData
    Dim r As Long, G As Long, B As Long
    Dim tR As Long, tG As Long, tB As Long, TC As Long
    Message "Enhancing image shadows..."
    SetProgBarMax PicWidthL
    SetProgBarVal 0
    Dim QuickVal As Long
    For X = 0 To PicWidthL
        QuickVal = X * 3
    For Y = 0 To PicHeightL
        'Get the temporary values
        tR = ImageData(QuickVal + 2, Y)
        tG = ImageData(QuickVal + 1, Y)
        tB = ImageData(QuickVal, Y)
        TC = Int((222 * tR + 707 * tG + 71 * tB) \ 1000)
        r = tR
        G = tG + Abs(r - TC)
        B = tB + Abs(G - TC)
        r = tR + Abs(B - TC)
        ByteMeL r
        ByteMeL G
        ByteMeL B
        'Assign those values into the array
        ImageData(QuickVal + 2, Y) = r
        ImageData(QuickVal + 1, Y) = G
        ImageData(QuickVal, Y) = B
    Next Y
        If X Mod 20 = 0 Then SetProgBarVal X
    Next X
    SetImageData
End Sub

