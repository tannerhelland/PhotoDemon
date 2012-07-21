Attribute VB_Name = "Filters_Miscellaneous"
'***************************************************************************
'Filter Module
'Copyright ©2000-2012 by Tanner Helland
'Created: 10/13/00
'Last updated: 11/January/07
'
'The general image filter module; contains unorganized routines at present.
'
'***************************************************************************

Option Explicit

'Loads the last Undo file and alpha-blends it with the current image
Public Sub MenuFadeLastEffect()
    Message "Fading last effect..."
    'Load the last undo file into the temporary picture box
    FormMain.ActiveForm.BackBuffer2.AutoSize = True
    FormMain.ActiveForm.BackBuffer2.Picture = LoadPicture("")
    FormMain.ActiveForm.BackBuffer2.Picture = LoadPicture(GetLastUndoFile())
    FormMain.ActiveForm.BackBuffer2.Picture = FormMain.ActiveForm.BackBuffer2.Image
    FormMain.ActiveForm.BackBuffer2.Refresh
    'Get that picture's information
    GetImageData2
    
    'Use these to determine minimum image sizes
    Dim minWidth As Long, minHeight As Long
    minWidth = PicWidthL
    minHeight = PicHeightL
        
    'Get the current picture's information
    GetImageData
    
    'Find the smallest dimensions (in case the picture has been at all resized -
    ' this prevents grievous 'out of subscript' errors)
    If minWidth > PicWidthL Then minWidth = PicWidthL
    If minHeight > PicHeightL Then minHeight = PicHeightL
    
    SetProgBarMax PicWidthL
    
    'Run a loop through the two arrays, blending each byte individually
    Dim QuickVal As Long
    For x = 0 To minWidth
        QuickVal = x * 3
    For y = 0 To minHeight
     For z = 0 To 2
        ImageData(QuickVal + z, y) = MixColors(ImageData(QuickVal + z, y), ImageData2(QuickVal + z, y), 50)
     Next z
    Next y
        If x Mod 20 = 0 Then SetProgBarVal x
    Next x
    
    SetProgBarVal PicWidthL
    
    'Clear out the temporary buffer & picture array (try to conserve SOME memory, heh)
    FormMain.ActiveForm.BackBuffer2.Picture = LoadPicture("")
    Erase ImageData2()
    
    'Post the new data
    SetImageData
    
End Sub

'Right now this is a work in progress; it's somewhat based off... <description forthcoming>
Public Sub MenuAnimate()

    MsgBox "Sorry, but this filter is still under heavy development.  It's disabled right now due to some instability in the code.  Stay tuned for updates!", vbInformation + vbOKOnly + vbApplicationModal, "Animate filter disabled... for now"
    
    Message "Animate filter canceled"
    
    Exit Sub

End Sub

Public Sub MenuSynthesize()
    Dim r As Long, g As Long, b As Long
    Dim TC As Long
    Message "Synthesizing image..."
    SetProgBarMax PicWidthL
    Dim QuickVal As Long
    
    For x = 0 To PicWidthL
        QuickVal = x * 3
    For y = 0 To PicHeightL
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        TC = ((222 * r + 707 * g + 71 * b) \ 1000)
        r = g + b - TC
        g = r + b - TC
        b = r + g - TC
        ByteMeL r
        ByteMeL g
        ByteMeL b
        ImageData(QuickVal + 2, y) = r
        ImageData(QuickVal + 1, y) = g
        ImageData(QuickVal, y) = b
    Next y
        If x Mod 20 = 0 Then SetProgBarVal x
    Next x
    Message "Image synthesized.  Redrawing new image..."
    SetImageData
End Sub

Public Sub MenuAlien()
    Dim r As Long, g As Long, b As Long
    Dim tR As Integer, tB As Integer, tG As Integer
    Message "Generating alien colors..."
    SetProgBarMax PicWidthL
    Dim QuickVal As Long
    For x = 0 To PicWidthL
        QuickVal = x * 3
    For y = 0 To PicHeightL
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        tR = b + g - r
        tG = r + b - g
        tB = r + g - b
        ByteMe tR
        ByteMe tG
        ByteMe tB
        ImageData(x * 3 + 2, y) = tR
        ImageData(x * 3 + 1, y) = tG
        ImageData(x * 3, y) = tB
    Next y
        If x Mod 20 = 0 Then SetProgBarVal x
    Next x
    SetImageData
End Sub

Public Sub MenuAntique()
    Dim r As Long, g As Long, b As Long
    Dim tR As Long, tB As Long, tG As Long
    Message "Running antique filter..."
    SetProgBarMax PicWidthL
    
    'Build a basic gamma conversion table (otherwise, the image is really
    'dim and hard to see)
    Dim LookUp(0 To 255) As Integer
    Dim TempVal As Single
    For x = 0 To 255
        TempVal = x / 255
        TempVal = TempVal ^ (1 / 1.6)
        TempVal = TempVal * 255
        If TempVal > 255 Then TempVal = 255
        If TempVal < 0 Then TempVal = 0
        LookUp(x) = TempVal
    Next x
    
    'Run a loop through the entire image, converting to sepia as we go
    Dim QuickVal As Long
    For x = 0 To PicWidthL
        QuickVal = x * 3
    For y = 0 To PicHeightL
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        tR = (r + g + b) / 3
        r = (r + tR) / 2
        g = (g + tR) / 2
        b = (b + tR) / 2
        r = (g * b) \ 256
        g = (b * r) \ 256
        b = (r * g) \ 256
        tR = r * 1.75
        tG = g * 1.75
        tB = b * 1.75
        If tR > 255 Then tR = 255
        If tG > 255 Then tG = 255
        If tB > 255 Then tB = 255
        tR = LookUp(tR)
        tG = LookUp(tG)
        tB = LookUp(tB)
        ImageData(QuickVal + 2, y) = tR
        ImageData(QuickVal + 1, y) = tG
        ImageData(QuickVal, y) = tB
    Next y
        If x Mod 20 = 0 Then SetProgBarVal x
    Next x
    SetImageData
End Sub

'Makes the picture appear like it has been shaken
Public Sub MenuVibrate()
    FilterSize = 5
    ReDim FM(-2 To 2, -2 To 2) As Long
    FM(-2, -2) = 1
    FM(-1, -1) = -1
    FM(0, 0) = 1
    FM(1, 1) = -1
    FM(2, 2) = 1
    FM(-1, 1) = 1
    FM(-2, 2) = -1
    FM(1, -1) = 1
    FM(2, -2) = -1
    FilterWeight = 1
    FilterBias = 0
    DoFilter "Vibrate"
End Sub

Public Sub MenuDream()
    Dim r As Integer, g As Integer, b As Integer
    Dim tR As Long, tB As Long, tG As Long, TC As Long
    Message "Applying dream filter..."
    SetProgBarMax PicWidthL
    Dim QuickVal As Long
    For x = 0 To PicWidthL
        QuickVal = x * 3
    For y = 0 To PicHeightL
        tR = ImageData(QuickVal + 2, y)
        tG = ImageData(QuickVal + 1, y)
        tB = ImageData(QuickVal, y)
        TC = Int((222 * tR + 707 * tG + 71 * tB) \ 1000)
        r = Abs(tR - TC) + Abs(tR - tG) + Abs(tR - tB) + (tR \ 2)
        g = Abs(tG - TC) + Abs(tG - tB) + Abs(tG - tR) + (tG \ 2)
        b = Abs(tB - TC) + Abs(tB - tR) + Abs(tB - tG) + (tB \ 2)
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

Public Sub MenuCompoundInvert(ByVal Divisor As Integer)
    Dim r As Long, g As Long, b As Long
    Dim tR As Long, tB As Long, tG As Long
    Message "Running compound inversion filter..."
    SetProgBarMax PicWidthL
    Dim QuickVal As Long
    
    For x = 0 To PicWidthL
        QuickVal = x * 3
    For y = 0 To PicHeightL
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        If r = 0 Then r = 1
        If g = 0 Then g = 1
        If b = 0 Then b = 1
        tR = (g * b) \ Divisor
        tG = (r * b) \ Divisor
        tB = (r * g) \ Divisor
        If tR > 255 Then tR = 255
        If tG > 255 Then tG = 255
        If tB > 255 Then tB = 255
        ImageData(QuickVal + 2, y) = tR
        ImageData(QuickVal + 1, y) = tG
        ImageData(QuickVal, y) = tB
    Next y
        If x Mod 20 = 0 Then SetProgBarVal x
    Next x
    SetImageData
End Sub

Public Sub MenuRadioactive()
    Dim r As Long, g As Long, b As Long
    Dim tR As Long, tB As Long, tG As Long
    Message "Running radioactive filter..."
    SetProgBarMax PicWidthL
    Dim QuickVal As Long
    For x = 0 To PicWidthL
        QuickVal = x * 3
    For y = 0 To PicHeightL
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        If r = 0 Then r = 1
        If g = 0 Then g = 1
        If b = 0 Then b = 1
        tR = (g * b) \ r
        tG = (r * b) \ g
        tB = (r * g) \ b
        If tR > 255 Then tR = 255
        If tG > 255 Then tG = 255
        If tB > 255 Then tB = 255
        tG = 255 - tG
        ImageData(QuickVal + 2, y) = tR
        ImageData(QuickVal + 1, y) = tG
        ImageData(QuickVal, y) = tB
    Next y
        If x Mod 20 = 0 Then SetProgBarVal x
    Next x
    SetImageData
End Sub

Public Sub MenuComicBook()
    Dim r As Long, g As Long, b As Long
    Dim TC As Long
    Message "Running comic book filter..."
    SetProgBarMax PicWidthL
    Dim QuickVal As Long
    
    For x = 0 To PicWidthL
        QuickVal = x * 3
    For y = 0 To PicHeightL
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        r = Abs(r * (g - b + g + r)) / 256
        g = Abs(r * (b - g + b + r)) / 256
        b = Abs(g * (b - g + b + r)) / 256
        TC = (222 * r + 707 * g + 71 * b) \ 1000
        ByteMeL TC
        ImageData(QuickVal + 2, y) = TC
        ImageData(QuickVal + 1, y) = TC
        ImageData(QuickVal, y) = TC
    Next y
        If x Mod 20 = 0 Then SetProgBarVal x
    Next x
    SetImageData
End Sub

'Subroutine for counting the number of unique colors in an image
Public Sub MenuCountColors()
    'No Undo necessary
    GetImageData
    Dim TC As Long, tR As Long, tB As Long, tG As Long
    Dim totalVal As Long
    
    Message "Counting image colors..."
    
    'Array for storing whether or not a color has been seen already
    Dim UniqueColors() As Boolean
    ReDim UniqueColors(0 To 16777216) As Boolean
    
    'Total number of unique colors
    totalVal = 0
    
    Dim QuickVal As Long
    
    SetProgBarMax PicWidthL
    For x = 0 To PicWidthL
        QuickVal = x * 3
    For y = 0 To PicHeightL
        tR = ImageData(QuickVal + 2, y)
        tG = ImageData(QuickVal + 1, y)
        tB = ImageData(QuickVal, y)
        TC = RGB(tR, tG, tB)
        If UniqueColors(TC) = False Then
            totalVal = totalVal + 1
            UniqueColors(TC) = True
        End If
    Next y
        If x Mod 20 = 0 Then SetProgBarVal x
    Next x
    SetProgBarVal PicWidthL
    Message "Total number of unique colors: " & totalVal
    MsgBox "Total number of unique colors: " & totalVal, vbOKOnly + vbApplicationModal + vbInformation, "Count Image Colors"
    
    'Attempt to free up whatever memory our dynamic array used
    Erase UniqueColors
    
    SetProgBarVal 0
End Sub

Public Sub MenuTest()
    
    'Currently testing plugin downloading
    zLibEnabled = False
    ScanEnabled = False
    FreeImageEnabled = False
    FormPluginDownloader.Show 1, FormMain
    
    Exit Sub
    
    BuildImageRestore
    GetImageData
    Dim r As Long, g As Long, b As Long
    Dim TC As Long, tR As Long, tB As Long, tG As Long
    Dim bR As Byte, bG As Byte, bB As Byte, bC As Byte
    Dim HH As Single, SS As Single, LL As Single
    Dim tH As Single, tS As Single, tL As Single
    Dim xCalc As Long, yCalc As Long
    Dim totalVal As Long
    totalVal = 0
    
    SetProgBarMax PicWidthL
    
    For x = 0 To PicWidthL
    For y = 0 To PicHeightL
        tR = ImageData(x * 3 + 2, y)
        tG = ImageData(x * 3 + 1, y)
        tB = ImageData(x * 3, y)
        TC = (tR + tG + tB) \ 3
        
        'If TC = 0 Then TC = 1
        'bR = TR
        'bG = TG
        'bB = TB
        'bC = TC
        
        'bR = bR Xor bC
        'bG = bG Xor bC
        'bB = bB Xor bC
        tR = TC * 1.2
        tG = TC
        tB = TC * 0.8
        
        ByteMeL tR
        ByteMeL tG
        ByteMeL tB
        
        ImageData(x * 3 + 2, y) = tR
        ImageData(x * 3 + 1, y) = tG
        ImageData(x * 3, y) = tB
    Next y
        If x Mod 20 = 0 Then SetProgBarVal x
    Next x
    
    SetProgBarVal 0
    
    SetImageData
    
    
    
    
    

    
    
    Exit Sub
    
    
    
    
    
    
    
    
    'xCalc = 10
    'yCalc = 5
    'Randomize
    Message "Running test filter..."
    SetProgBarMax PicWidthL
    xCalc = PicWidthL + PicHeightL
    For x = 0 To PicWidthL
    For y = 0 To PicHeightL
        tR = ImageData(x * 3 + 2, y)
        tG = ImageData(x * 3 + 1, y)
        tB = ImageData(x * 3, y)
        TC = Int((222 * tR + 707 * tG + 71 * tB) \ 1000)
        'Get the hue and saturation
        'tRGBToHSL TR, TG, TB, HH, SS, LL

        'Working on night-vision
        r = TC \ 2
        b = r
        g = tG + ((tR + tB) \ 4)
        
        'Convert back to RGB using our artificial luminance values
        'tHSLToRGB HH, SS, LL, r, g, b
        
        ByteMeL r
        ByteMeL g
        ByteMeL b
        ImageData(x * 3 + 2, y) = r
        ImageData(x * 3 + 1, y) = g
        ImageData(x * 3, y) = b
    Next y
        If x Mod 20 = 0 Then SetProgBarVal x
    Next x
    SetImageData
End Sub

'Here's all the filters I haven't yet found a home for
Public Sub TempHolderForUnplacedFilters()
        Dim r As Long, g As Long, b As Long
        Dim TC As Long, tR As Long, tG As Long, tB As Long
        Dim xCalc As Single, yCalc As Single
        
        'Purpleize
        r = tR
        g = tG + Abs(r - TC)
        b = tB + Abs(g - TC)
        r = tR + Abs(b - TC)
        
        'Trigonometric noise
        xCalc = 6
        yCalc = xCalc \ 2
        'horizontal
        r = tR * Cos((x Mod xCalc - yCalc) / yCalc)
        g = tG * Cos((x Mod xCalc - yCalc) / yCalc)
        b = tB * Cos((x Mod xCalc - yCalc) / yCalc)
        'vertical
        r = tR * Cos(((x + y) Mod xCalc - yCalc) / yCalc)
        g = tG * Cos(((x + y) Mod xCalc - yCalc) / yCalc)
        b = tB * Cos(((x + y) Mod xCalc - yCalc) / yCalc)
        
        'Two-tiered invert
        If tR < 128 Then r = Abs(tR - 128) Else r = Abs(tR - 255)
        If tG < 128 Then g = Abs(tG - 128) Else g = Abs(tG - 255)
        If tB < 128 Then b = Abs(tB - 128) Else b = Abs(tB - 255)
        
        'Difference between colors (can cycle between colors)
        r = Abs(tR - tB)
        g = Abs(tG - tR)
        b = Abs(tB - tG)
        
        'Psycho versions...
        If TC = 0 Then TC = 1
        r = tR Mod TC
        g = tG Mod TC
        b = tB Mod TC
        'and all the bitwise operators
        
        'AutoEnhance clone...?
        r = (tR - TC) + tR
        g = (tG - TC) + tG
        b = (tB - TC) + tB
        'another AutoEnhance clone?
        If TC = 0 Then TC = 1
        r = tR / TC * tR
        g = tG / TC * TC
        b = tB / TC * tB
        'Yet another AutoEnhance version...
        r = Abs(tR - TC) - TC + 2 * tR
        g = Abs(tG - TC) - TC + 2 * tG
        b = Abs(tB - TC) - TC + 2 * tB
        
        'Strange infrared effect
        r = Abs(tR - 64)
        g = Abs(r - 64)
        b = Abs(g - 64)
        TC = Int((222 * tR + 707 * tG + 71 * tB) \ 1000)
        r = TC + 70
        r = r + (((r - 128) * 100) \ 100)
        g = Abs(TC - 6) + 70
        g = g + (((g - 128) * 100) \ 100)
        b = (TC + 4) + 70
        b = b + (((b - 128) * 100) \ 100)
        r = (r - TC) * 5
        g = (g - TC) * 5
        b = (b - TC) * 5
End Sub




'HSL conversion crap
Public Sub tRGBToHSL(r As Long, g As Long, b As Long, h As Single, s As Single, l As Single)
Dim Max As Single
Dim Min As Single
Dim delta As Single
Dim rR As Single, rG As Single, rB As Single
   rR = r / 255: rG = g / 255: rB = b / 255

'{Given: rgb each in [0,1].
' Desired: h in [0,360] and s in [0,1], except if s=0, then h=UNDEFINED.}
'It actually gives h from [-1,5]
        Max = Maximum(rR, rG, rB)
        Min = Minimum(rR, rG, rB)
        l = (Max + Min) / 2    '{This is the lightness}
        '{Next calculate saturation}
        If Max = Min Then
            'begin {Acrhomatic case}
            s = 0
            h = 0
           'end {Acrhomatic case}
        Else
           'begin {Chromatic case}
                '{First calculate the saturation.}
           If l <= 0.5 Then
               s = (Max - Min) / (Max + Min)
           Else
               s = (Max - Min) / (2 - Max - Min)
            End If
            '{Next calculate the hue.}
            delta = Max - Min
           If rR = Max Then
                h = (rG - rB) / delta    '{Resulting color is between yellow and magenta}
           ElseIf rG = Max Then
                h = 2 + (rB - rR) / delta '{Resulting color is between cyan and yellow}
           ElseIf rB = Max Then
                h = 4 + (rR - rG) / delta '{Resulting color is between magenta and cyan}
           End If
            'Debug.Print h
            'h = h * 60
           'If h < 0# Then
           '     h = h + 360            '{Make degrees be nonnegative}
           'End If
        'end {Chromatic Case}
      End If

'Tanner's hack: transfer the values into ones I can use, dangit; this gives me
'hue on [0,240], saturation on [0,255], and luminance on [0,255]
    'H = Int(H * 40 + 40)
    'S = Int(S * 255)
    'L = Int(L * 255)
End Sub

Public Sub tHSLToRGB(h As Single, s As Single, l As Single, r As Long, g As Long, b As Long)
Dim rR As Single, rG As Single, rB As Single
Dim Min As Single, Max As Single
'This one requires the stupid values; such is life

   If s = 0 Then
      ' Achromatic case:
      rR = l: rG = l: rB = l
   Else
      ' Chromatic case:
      ' delta = Max-Min
      If l <= 0.5 Then
         's = (Max - Min) / (Max + Min)
         ' Get Min value:
         Min = l * (1 - s)
      Else
         's = (Max - Min) / (2 - Max - Min)
         ' Get Min value:
         Min = l - s * (1 - l)
      End If
      ' Get the Max value:
      Max = 2 * l - Min
      
      ' Now depending on sector we can evaluate the h,l,s:
      If (h < 1) Then
         rR = Max
         If (h < 0) Then
            rG = Min
            rB = rG - h * (Max - Min)
         Else
            rB = Min
            rG = h * (Max - Min) + rB
         End If
      ElseIf (h < 3) Then
         rG = Max
         If (h < 2) Then
            rB = Min
            rR = rB - (h - 2) * (Max - Min)
         Else
            rR = Min
            rB = (h - 2) * (Max - Min) + rR
         End If
      Else
         rB = Max
         If (h < 4) Then
            rR = Min
            rG = rR - (h - 4) * (Max - Min)
         Else
            rG = Min
            rR = (h - 4) * (Max - Min) + rG
         End If
         
      End If
            
   End If
   r = rR * 255: g = rG * 255: b = rB * 255
End Sub

'Return the maximum of three variables
Private Function Maximum(rR As Single, rG As Single, rB As Single) As Single
   If (rR > rG) Then
      If (rR > rB) Then
         Maximum = rR
      Else
         Maximum = rB
      End If
   Else
      If (rB > rG) Then
         Maximum = rB
      Else
         Maximum = rG
      End If
   End If
End Function

'Return the minimum of three variables
Private Function Minimum(rR As Single, rG As Single, rB As Single) As Single
   If (rR < rG) Then
      If (rR < rB) Then
         Minimum = rR
      Else
         Minimum = rB
      End If
   Else
      If (rB < rG) Then
         Minimum = rB
      Else
         Minimum = rG
      End If
   End If
End Function
