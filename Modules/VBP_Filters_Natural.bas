Attribute VB_Name = "Filters_Natural"
'***************************************************************************
'"Natural" Filters
'©2000-2012 Tanner Helland
'Created: 8/April/02
'Last updated: 12/January/07
'Last update: added the cool "rainbow" effect and finally finished basic
'             optimization of all the functions.
'
'Runs all nature-type filters.  Includes water, steel, burn, rainbow, etc.
'
'***************************************************************************

Option Explicit

Public Sub MenuRainbow()
    Dim R As Long, G As Long, B As Long
    Dim HH As Single, SS As Single, LL As Single
    Dim hVal As Single
    Message "Applying rainbow effect..."
    GetImageData
    SetProgBarMax PicWidthL
    Dim QuickVal As Long
    For x = 0 To PicWidthL
        QuickVal = x * 3
        
        'Based on our x-position, gradient a value between -1 and 5
        hVal = (x / PicWidthL) * 360
        hVal = (hVal - 60) / 60
        
    For y = 0 To PicHeightL
        'Get the temporary values
        R = ImageData(QuickVal + 2, y)
        G = ImageData(QuickVal + 1, y)
        B = ImageData(QuickVal, y)
        'Get the hue and saturation
        tRGBToHSL R, G, B, HH, SS, LL
        'Convert back to RGB using our artificial hue value
        tHSLToRGB hVal, 0.5, LL, R, G, B
        'Assign those values into the array
        ImageData(QuickVal + 2, y) = R
        ImageData(QuickVal + 1, y) = G
        ImageData(QuickVal, y) = B
    Next y
        If x Mod 20 = 0 Then SetProgBarVal x
    Next x
    
    SetImageData
End Sub

Public Sub MenuFogEffect()
    Dim R As Long, G As Long, B As Long
    Dim FogLimit As Long
    'Change this value to change the "thickness" of the fog
    FogLimit = 36
    Dim QuickVal As Long
    Message "Generating fog..."
    SetProgBarMax PicWidthL
    For x = 0 To PicWidthL
        QuickVal = x * 3
    For y = 0 To PicHeightL
        R = ImageData(QuickVal + 2, y)
        G = ImageData(QuickVal + 1, y)
        B = ImageData(QuickVal, y)
        If R > 127 Then
            R = R - FogLimit
            If R < 127 Then R = 127
        Else
            R = R + FogLimit
            If R > 127 Then R = 127
        End If
        If G > 127 Then
            G = G - FogLimit
            If G < 127 Then G = 127
        Else
            G = G + FogLimit
            If G > 127 Then G = 127
        End If
        If B > 127 Then
            B = B - FogLimit
            If B < 127 Then B = 127
        Else
            B = B + FogLimit
            If B > 127 Then B = 127
        End If
        
        ImageData(QuickVal + 2, y) = R
        ImageData(QuickVal + 1, y) = G
        ImageData(QuickVal, y) = B
    Next y
        If x Mod 20 = 0 Then SetProgBarVal x
    Next x
    SetImageData
End Sub

Public Sub MenuWater()
    Dim R As Long, G As Long, B As Long
    Dim TC As Long
    Dim QuickVal As Long
    Message "Running water filter..."
    SetProgBarMax PicWidthL
    For x = 0 To PicWidthL
        QuickVal = x * 3
    For y = 0 To PicHeightL
        R = ImageData(QuickVal + 2, y)
        G = ImageData(QuickVal + 1, y)
        B = ImageData(QuickVal, y)
        TC = ((222 * R + 707 * G + 71 * B) \ 1000)
        R = TC - G - B
        G = TC - R - B
        B = TC - R - G
        ByteMeL R
        ByteMeL G
        ByteMeL B
        ImageData(QuickVal + 2, y) = R
        ImageData(QuickVal + 1, y) = G
        ImageData(QuickVal, y) = B
    Next y
        If x Mod 20 = 0 Then SetProgBarVal x
    Next x
    SetImageData
End Sub

Public Sub MenuAtmospheric()
    Dim R As Integer, G As Integer, B As Integer
    Dim TR As Integer, TB As Integer, TG As Integer
    Dim QuickVal As Long
    Message "Running atmospheric filter..."
    SetProgBarMax PicWidthL
    For x = 0 To PicWidthL
        QuickVal = x * 3
    For y = 0 To PicHeightL
        R = ImageData(QuickVal + 2, y)
        G = ImageData(QuickVal + 1, y)
        B = ImageData(QuickVal, y)
        TR = (G + B) \ 2
        TG = (R + B) \ 2
        TB = (R + G) \ 2
        ImageData(QuickVal + 2, y) = TR
        ImageData(QuickVal + 1, y) = TG
        ImageData(QuickVal, y) = TB
    Next y
        If x Mod 20 = 0 Then SetProgBarVal x
    Next x
    SetImageData
End Sub

Public Sub MenuFrozen()
    Dim R As Integer, G As Integer, B As Integer
    Message "Running freeze filter..."
    SetProgBarMax PicWidthL
    Dim QuickVal As Long
    For x = 0 To PicWidthL
        QuickVal = x * 3
    For y = 0 To PicHeightL
        R = ImageData(QuickVal + 2, y)
        G = ImageData(QuickVal + 1, y)
        B = ImageData(QuickVal, y)
        R = Abs((R - G - B) * 1.5)
        G = Abs((G - B - R) * 1.5)
        B = Abs((B - R - G) * 1.5)
        ByteMe R
        ByteMe G
        ByteMe B
        ImageData(QuickVal + 2, y) = R
        ImageData(QuickVal + 1, y) = G
        ImageData(QuickVal, y) = B
    Next y
        If x Mod 20 = 0 Then SetProgBarVal x
    Next x
    SetImageData
End Sub

Public Sub MenuLava()
    Dim R As Long, G As Long, B As Long
    Dim TC As Long
    Message "Running lava filter..."
    SetProgBarMax PicWidthL
    Dim QuickVal As Long
    For x = 0 To PicWidthL
        QuickVal = x * 3
    For y = 0 To PicHeightL
        R = ImageData(QuickVal + 2, y)
        G = ImageData(QuickVal + 1, y)
        B = ImageData(QuickVal, y)
        TC = Int((222 * R + 707 * G + 71 * B) \ 1000)
        R = TC
        G = Abs(B - 128)
        B = Abs(B - 128)
        ImageData(QuickVal + 2, y) = R
        ImageData(QuickVal + 1, y) = G
        ImageData(QuickVal, y) = B
    Next y
        If x Mod 20 = 0 Then SetProgBarVal x
    Next x
    SetImageData
End Sub

Public Sub MenuBurn()
    Dim R As Long, G As Long, B As Long
    Dim TC As Long
    Message "Running burn filter..."
    SetProgBarMax PicWidthL
    Dim QuickVal As Long
    For x = 0 To PicWidthL
        QuickVal = x * 3
    For y = 0 To PicHeightL
        R = ImageData(QuickVal + 2, y)
        G = ImageData(QuickVal + 1, y)
        B = ImageData(QuickVal, y)
        TC = Int((222 * R + 707 * G + 71 * B) \ 1000)
        R = TC * 3
        G = TC
        B = TC \ 3
        ByteMeL R
        ByteMeL G
        ByteMeL B
        ImageData(QuickVal + 2, y) = R
        ImageData(QuickVal + 1, y) = G
        ImageData(QuickVal, y) = B
    Next y
        If x Mod 20 = 0 Then SetProgBarVal x
    Next x
    SetImageData
End Sub

Public Sub MenuOcean()
    Dim R As Long, G As Long, B As Long
    Dim TC As Long
    Message "Running ocean filter..."
    SetProgBarMax PicWidthL
    Dim QuickVal As Long
    For x = 0 To PicWidthL
        QuickVal = x * 3
    For y = 0 To PicHeightL
        R = ImageData(QuickVal + 2, y)
        G = ImageData(QuickVal + 1, y)
        B = ImageData(QuickVal, y)
        TC = Int((222 * R + 707 * G + 71 * B) \ 1000)
        R = TC \ 3
        G = TC
        B = TC * 3
        ByteMeL R
        ByteMeL G
        ByteMeL B
        ImageData(QuickVal + 2, y) = R
        ImageData(QuickVal + 1, y) = G
        ImageData(QuickVal, y) = B
    Next y
        If x Mod 20 = 0 Then SetProgBarVal x
    Next x
    SetImageData
End Sub

Public Sub MenuSteel()
    Dim R As Long, G As Long, B As Long
    Dim TC As Long, TR As Long
    Message "Running steel filter..."
    SetProgBarMax PicWidthL
    Dim QuickVal As Long
    For x = 0 To PicWidthL
        QuickVal = x * 3
    For y = 0 To PicHeightL
        TR = ImageData(QuickVal + 2, y)
        R = Abs(TR - 64)
        G = Abs(R - 64)
        B = Abs(G - 64)
        TC = Int((222 * R + 707 * G + 71 * B) \ 1000)
        R = TC + 70
        R = R + (((R - 128) * 100) \ 100)
        G = Abs(TC - 6) + 70
        G = G + (((G - 128) * 100) \ 100)
        B = (TC + 5) + 70
        B = B + (((B - 128) * 100) \ 100)
        ByteMeL R
        ByteMeL G
        ByteMeL B
        ImageData(QuickVal + 2, y) = R
        ImageData(QuickVal + 1, y) = G
        ImageData(QuickVal, y) = B
    Next y
        If x Mod 20 = 0 Then SetProgBarVal x
    Next x
    SetImageData
End Sub
