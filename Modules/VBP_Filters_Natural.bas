Attribute VB_Name = "Filters_Natural"
'***************************************************************************
'"Natural" Filters
'Copyright ©2000-2012 by Tanner Helland
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
    Dim r As Long, g As Long, b As Long
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
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        'Get the hue and saturation
        tRGBToHSL r, g, b, HH, SS, LL
        'Convert back to RGB using our artificial hue value
        tHSLToRGB hVal, 0.5, LL, r, g, b
        'Assign those values into the array
        ImageData(QuickVal + 2, y) = r
        ImageData(QuickVal + 1, y) = g
        ImageData(QuickVal, y) = b
    Next y
        If x Mod 20 = 0 Then SetProgBarVal x
    Next x
    
    SetImageData
End Sub

Public Sub MenuFogEffect()
    Dim r As Long, g As Long, b As Long
    Dim FogLimit As Long
    'Change this value to change the "thickness" of the fog
    FogLimit = 36
    Dim QuickVal As Long
    Message "Generating fog..."
    SetProgBarMax PicWidthL
    For x = 0 To PicWidthL
        QuickVal = x * 3
    For y = 0 To PicHeightL
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        If r > 127 Then
            r = r - FogLimit
            If r < 127 Then r = 127
        Else
            r = r + FogLimit
            If r > 127 Then r = 127
        End If
        If g > 127 Then
            g = g - FogLimit
            If g < 127 Then g = 127
        Else
            g = g + FogLimit
            If g > 127 Then g = 127
        End If
        If b > 127 Then
            b = b - FogLimit
            If b < 127 Then b = 127
        Else
            b = b + FogLimit
            If b > 127 Then b = 127
        End If
        
        ImageData(QuickVal + 2, y) = r
        ImageData(QuickVal + 1, y) = g
        ImageData(QuickVal, y) = b
    Next y
        If x Mod 20 = 0 Then SetProgBarVal x
    Next x
    SetImageData
End Sub

Public Sub MenuWater()
    Dim r As Long, g As Long, b As Long
    Dim TC As Long
    Dim QuickVal As Long
    Message "Running water filter..."
    SetProgBarMax PicWidthL
    For x = 0 To PicWidthL
        QuickVal = x * 3
    For y = 0 To PicHeightL
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        TC = ((222 * r + 707 * g + 71 * b) \ 1000)
        r = TC - g - b
        g = TC - r - b
        b = TC - r - g
        ByteMeL r
        ByteMeL g
        ByteMeL b
        ImageData(QuickVal + 2, y) = r
        ImageData(QuickVal + 1, y) = g
        ImageData(QuickVal, y) = b
    Next y
        If x Mod 20 = 0 Then SetProgBarVal x
    Next x
    SetImageData
End Sub

Public Sub MenuAtmospheric()
    Dim r As Integer, g As Integer, b As Integer
    Dim tR As Integer, tB As Integer, tG As Integer
    Dim QuickVal As Long
    Message "Running atmospheric filter..."
    SetProgBarMax PicWidthL
    For x = 0 To PicWidthL
        QuickVal = x * 3
    For y = 0 To PicHeightL
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        tR = (g + b) \ 2
        tG = (r + b) \ 2
        tB = (r + g) \ 2
        ImageData(QuickVal + 2, y) = tR
        ImageData(QuickVal + 1, y) = tG
        ImageData(QuickVal, y) = tB
    Next y
        If x Mod 20 = 0 Then SetProgBarVal x
    Next x
    SetImageData
End Sub

Public Sub MenuFrozen()
    Dim r As Integer, g As Integer, b As Integer
    Message "Running freeze filter..."
    SetProgBarMax PicWidthL
    Dim QuickVal As Long
    For x = 0 To PicWidthL
        QuickVal = x * 3
    For y = 0 To PicHeightL
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        r = Abs((r - g - b) * 1.5)
        g = Abs((g - b - r) * 1.5)
        b = Abs((b - r - g) * 1.5)
        ByteMe r
        ByteMe g
        ByteMe b
        ImageData(QuickVal + 2, y) = r
        ImageData(QuickVal + 1, y) = g
        ImageData(QuickVal, y) = b
    Next y
        If x Mod 20 = 0 Then SetProgBarVal x
    Next x
    SetImageData
End Sub

Public Sub MenuLava()
    Dim r As Long, g As Long, b As Long
    Dim TC As Long
    Message "Running lava filter..."
    SetProgBarMax PicWidthL
    Dim QuickVal As Long
    For x = 0 To PicWidthL
        QuickVal = x * 3
    For y = 0 To PicHeightL
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        TC = Int((222 * r + 707 * g + 71 * b) \ 1000)
        r = TC
        g = Abs(b - 128)
        b = Abs(b - 128)
        ImageData(QuickVal + 2, y) = r
        ImageData(QuickVal + 1, y) = g
        ImageData(QuickVal, y) = b
    Next y
        If x Mod 20 = 0 Then SetProgBarVal x
    Next x
    SetImageData
End Sub

Public Sub MenuBurn()
    Dim r As Long, g As Long, b As Long
    Dim TC As Long
    Message "Running burn filter..."
    SetProgBarMax PicWidthL
    Dim QuickVal As Long
    For x = 0 To PicWidthL
        QuickVal = x * 3
    For y = 0 To PicHeightL
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        TC = Int((222 * r + 707 * g + 71 * b) \ 1000)
        r = TC * 3
        g = TC
        b = TC \ 3
        ByteMeL r
        ByteMeL g
        ByteMeL b
        ImageData(QuickVal + 2, y) = r
        ImageData(QuickVal + 1, y) = g
        ImageData(QuickVal, y) = b
    Next y
        If x Mod 20 = 0 Then SetProgBarVal x
    Next x
    SetImageData
End Sub

Public Sub MenuOcean()
    Dim r As Long, g As Long, b As Long
    Dim TC As Long
    Message "Running ocean filter..."
    SetProgBarMax PicWidthL
    Dim QuickVal As Long
    For x = 0 To PicWidthL
        QuickVal = x * 3
    For y = 0 To PicHeightL
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        TC = Int((222 * r + 707 * g + 71 * b) \ 1000)
        r = TC \ 3
        g = TC
        b = TC * 3
        ByteMeL r
        ByteMeL g
        ByteMeL b
        ImageData(QuickVal + 2, y) = r
        ImageData(QuickVal + 1, y) = g
        ImageData(QuickVal, y) = b
    Next y
        If x Mod 20 = 0 Then SetProgBarVal x
    Next x
    SetImageData
End Sub

Public Sub MenuSteel()
    Dim r As Long, g As Long, b As Long
    Dim TC As Long, tR As Long
    Message "Running steel filter..."
    SetProgBarMax PicWidthL
    Dim QuickVal As Long
    For x = 0 To PicWidthL
        QuickVal = x * 3
    For y = 0 To PicHeightL
        tR = ImageData(QuickVal + 2, y)
        r = Abs(tR - 64)
        g = Abs(r - 64)
        b = Abs(g - 64)
        TC = Int((222 * r + 707 * g + 71 * b) \ 1000)
        r = TC + 70
        r = r + (((r - 128) * 100) \ 100)
        g = Abs(TC - 6) + 70
        g = g + (((g - 128) * 100) \ 100)
        b = (TC + 5) + 70
        b = b + (((b - 128) * 100) \ 100)
        ByteMeL r
        ByteMeL g
        ByteMeL b
        ImageData(QuickVal + 2, y) = r
        ImageData(QuickVal + 1, y) = g
        ImageData(QuickVal, y) = b
    Next y
        If x Mod 20 = 0 Then SetProgBarVal x
    Next x
    SetImageData
End Sub
