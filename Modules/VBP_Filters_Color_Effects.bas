Attribute VB_Name = "Filters_Color_Effects"
'***************************************************************************
'Filter (Color Effects) Interface
'Copyright ©2000-2012 by Tanner Helland
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
    For x = 0 To PicWidthL
        QuickVal = x * 3
    For y = 0 To PicHeightL
        ImageData(QuickVal, y) = 255 Xor ImageData(QuickVal, y)
        ImageData(QuickVal + 1, y) = 255 Xor ImageData(QuickVal + 1, y)
        ImageData(QuickVal + 2, y) = 255 Xor ImageData(QuickVal + 2, y)
    Next y
        If x Mod 10 = 0 Then SetProgBarVal x
    Next x
    
    SetImageData

End Sub

'Rechannel an image (red, green, or blue)
Public Sub MenuRechannel(ByVal RType As Byte)
    Dim QuickVal As Long
    Message "Rechanneling colors..."
    SetProgBarMax PicWidthL
    For x = 0 To PicWidthL
        QuickVal = x * 3
    For y = 0 To PicHeightL
        If RType = 0 Then
            ImageData(QuickVal, y) = 0
            ImageData(QuickVal + 1, y) = 0
        ElseIf RType = 1 Then
            ImageData(QuickVal, y) = 0
            ImageData(QuickVal + 2, y) = 0
        Else
            ImageData(QuickVal + 1, y) = 0
            ImageData(QuickVal + 2, y) = 0
        End If
    Next y
        If x Mod 20 = 0 Then SetProgBarVal x
    Next x
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
    SetImageData
End Sub

'Generate a luminance-negative version of the image
Public Sub MenuNegative()
    GetImageData
    Dim r As Long, g As Long, b As Long
    Dim HH As Single, SS As Single, LL As Single
    Message "Generating film negative..."
    SetProgBarMax PicWidthL
    SetProgBarVal 0
    Dim QuickVal As Long
    For x = 0 To PicWidthL
        QuickVal = x * 3
    For y = 0 To PicHeightL
        'Get the temporary values
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        'Get the hue and saturation
        tRGBToHSL r, g, b, HH, SS, LL
        'Convert back to RGB using our artificial luminance values
        tHSLToRGB HH, SS, 1 - LL, r, g, b
        'Assign those values into the array
        ImageData(QuickVal + 2, y) = r
        ImageData(QuickVal + 1, y) = g
        ImageData(QuickVal, y) = b
    Next y
        If x Mod 20 = 0 Then SetProgBarVal x
    Next x
    SetImageData
    Message "Finished."
End Sub

'Invert the hue of an image
Public Sub MenuInvertHue()
    Dim r As Long, g As Long, b As Long
    Dim tR As Long, tB As Long, tG As Long
    Dim HH As Single, SS As Single, LL As Single
    Message "Inverting image hue..."
    SetProgBarMax PicWidthL
    Dim QuickVal As Long
    For x = 0 To PicWidthL
        QuickVal = x * 3
    For y = 0 To PicHeightL
        tR = ImageData(QuickVal + 2, y)
        tG = ImageData(QuickVal + 1, y)
        tB = ImageData(QuickVal, y)
        
        'Get the hue and saturation
        tRGBToHSL tR, tG, tB, HH, SS, LL
        HH = 6 - (HH + 1) - 1
        'Convert back to RGB using our artificial hue value
        tHSLToRGB HH, SS, LL, r, g, b
        
        ImageData(QuickVal + 2, y) = r
        ImageData(QuickVal + 1, y) = g
        ImageData(QuickVal, y) = b
    Next y
        If x Mod 20 = 0 Then SetProgBarVal x
    Next x
    SetImageData
End Sub

'Enhance CONTRAST
Public Sub MenuAutoEnhanceContrast()
    GetImageData
    Dim r As Long, g As Long, b As Long
    Dim tR As Long, tG As Long, tB As Long, TC As Long
    Message "Enhancing image contrast..."
    SetProgBarMax PicWidthL
    SetProgBarVal 0
    Dim QuickVal As Long
    For x = 0 To PicWidthL
        QuickVal = x * 3
    For y = 0 To PicHeightL
        'Get the temporary values
        tR = ImageData(QuickVal + 2, y)
        tG = ImageData(QuickVal + 1, y)
        tB = ImageData(QuickVal, y)
        TC = Int((222 * tR + 707 * tG + 71 * tB) \ 1000)
        r = Abs(tR - TC) + tR
        g = Abs(tG - TC) + tG
        b = Abs(tB - TC) + tB
        ByteMeL r
        ByteMeL g
        ByteMeL b
        'Assign those values into the array
        ImageData(QuickVal + 2, y) = r
        ImageData(QuickVal + 1, y) = g
        ImageData(QuickVal, y) = b
    Next y
        If x Mod 20 = 0 Then SetProgBarVal x
    Next x
    SetImageData
End Sub

'Enhance HIGHLIGHTS
Public Sub MenuAutoEnhanceHighlights()
    GetImageData
    Dim r As Long, g As Long, b As Long
    Dim tR As Long, tG As Long, tB As Long, TC As Long
    Message "Enhancing image highlights..."
    SetProgBarMax PicWidthL
    SetProgBarVal 0
    Dim QuickVal As Long
    For x = 0 To PicWidthL
        QuickVal = x * 3
    For y = 0 To PicHeightL
        'Get the temporary values
        tR = ImageData(QuickVal + 2, y)
        tG = ImageData(QuickVal + 1, y)
        tB = ImageData(QuickVal, y)
        TC = Int((222 * tR + 707 * tG + 71 * tB) \ 1000)
        r = Abs(tR - TC) + TC
        g = Abs(tG - TC) + TC
        b = Abs(tB - TC) + TC
        ByteMeL r
        ByteMeL g
        ByteMeL b
        'Assign those values into the array
        ImageData(QuickVal + 2, y) = r
        ImageData(QuickVal + 1, y) = g
        ImageData(QuickVal, y) = b
    Next y
        If x Mod 20 = 0 Then SetProgBarVal x
    Next x
    SetImageData
End Sub

'Enhance MIDTONES
Public Sub MenuAutoEnhanceMidtones()
    GetImageData
    Dim r As Long, g As Long, b As Long
    Dim tR As Long, tG As Long, tB As Long, TC As Long
    Message "Enhancing image midtones..."
    SetProgBarMax PicWidthL
    SetProgBarVal 0
    Dim QuickVal As Long
    For x = 0 To PicWidthL
        QuickVal = x * 3
    For y = 0 To PicHeightL
        'Get the temporary values
        tR = ImageData(QuickVal + 2, y)
        tG = ImageData(QuickVal + 1, y)
        tB = ImageData(QuickVal, y)
        TC = Int((222 * tR + 707 * tG + 71 * tB) \ 1000)
        r = tR - (TC - tR)
        g = tG - (TC - tG)
        b = tB - (TC - tB)
        ByteMeL r
        ByteMeL g
        ByteMeL b
        'Assign those values into the array
        ImageData(QuickVal + 2, y) = r
        ImageData(QuickVal + 1, y) = g
        ImageData(QuickVal, y) = b
    Next y
        If x Mod 20 = 0 Then SetProgBarVal x
    Next x
    SetImageData
End Sub

'Enhance SHADOWS
Public Sub MenuAutoEnhanceShadows()
    GetImageData
    Dim r As Long, g As Long, b As Long
    Dim tR As Long, tG As Long, tB As Long, TC As Long
    Message "Enhancing image shadows..."
    SetProgBarMax PicWidthL
    SetProgBarVal 0
    Dim QuickVal As Long
    For x = 0 To PicWidthL
        QuickVal = x * 3
    For y = 0 To PicHeightL
        'Get the temporary values
        tR = ImageData(QuickVal + 2, y)
        tG = ImageData(QuickVal + 1, y)
        tB = ImageData(QuickVal, y)
        TC = Int((222 * tR + 707 * tG + 71 * tB) \ 1000)
        r = tR
        g = tG + Abs(r - TC)
        b = tB + Abs(g - TC)
        r = tR + Abs(b - TC)
        ByteMeL r
        ByteMeL g
        ByteMeL b
        'Assign those values into the array
        ImageData(QuickVal + 2, y) = r
        ImageData(QuickVal + 1, y) = g
        ImageData(QuickVal, y) = b
    Next y
        If x Mod 20 = 0 Then SetProgBarVal x
    Next x
    SetImageData
End Sub

