Attribute VB_Name = "Filters_BlackAndWhite"
'***************************************************************************
'Black and White Conversion Handler
'©2000-2012 Tanner Helland
'Created: 22/December/01
'Last updated: 13/January/07
'Last update: Finished basic optimizations for each routine
'
'All of my 1-bit conversion routines.  Each one has its intrinsic strengths
'and weaknesses, so I've included each in the program.  Similarly, some are pretty
'advanced while others are simple, but each could - feasibly - serve a unique
'purpose.
'
'***************************************************************************

Option Explicit

'A standard threshold-based reduction.  Very simple.
Public Sub MenuThreshold(ByVal Threshold As Long)
        Dim r As Integer, g As Integer, B As Integer
        Dim CurrentColor As Integer
        Dim QuickVal As Long
        Message "Generating threshold data..."
        SetProgBarMax PicWidthL
        For x = 0 To PicWidthL
            QuickVal = x * 3
        For y = 0 To PicHeightL
            r = ImageData(QuickVal + 2, y)
            g = ImageData(QuickVal + 1, y)
            B = ImageData(QuickVal, y)
            CurrentColor = (r + g + B) / 3
            If CurrentColor > Threshold Then
                ImageData(QuickVal, y) = 255
                ImageData(QuickVal + 1, y) = 255
                ImageData(QuickVal + 2, y) = 255
            Else
                ImageData(QuickVal, y) = 0
                ImageData(QuickVal + 1, y) = 0
                ImageData(QuickVal + 2, y) = 0
            End If
        Next y
            If x Mod 20 = 0 Then SetProgBarVal x
        Next x
    SetImageData
End Sub

'Nearest color conversion - about as simple as it gets.
Public Sub MenuBWNearestColor()
    Dim r As Integer, g As Integer, B As Integer
    Dim TC As Integer
    Dim QuickVal As Long
    Message "Converting to black and white..."
    SetProgBarMax PicWidthL
    For x = 0 To PicWidthL
        QuickVal = x * 3
    For y = 0 To PicHeightL
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        B = ImageData(QuickVal, y)
        TC = (r + g + B) \ 3
        If TC < 128 Then
            ImageData(QuickVal + 2, y) = 0
            ImageData(QuickVal + 1, y) = 0
            ImageData(QuickVal, y) = 0
        Else
            ImageData(QuickVal + 2, y) = 255
            ImageData(QuickVal + 1, y) = 255
            ImageData(QuickVal, y) = 255
        End If
    Next y
        If x Mod 20 = 0 Then SetProgBarVal x
    Next x
    SetImageData
End Sub

'This is a strange method I found on the net somewhere.  It's white-weighted, setting the
'pixel to black only if every component is less than 128.
Public Sub MenuBWNearestColor2()
    Dim r As Integer, g As Integer, B As Integer
    Dim QuickVal As Long
    Message "Converting to black and white..."
    SetProgBarMax PicWidthL
    For x = 0 To PicWidthL
        QuickVal = x * 3
    For y = 0 To PicHeightL
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        B = ImageData(QuickVal, y)
        If r < 128 And g < 128 And B < 128 Then
            ImageData(QuickVal + 2, y) = 0
            ImageData(QuickVal + 1, y) = 0
            ImageData(QuickVal, y) = 0
        Else
            ImageData(QuickVal + 2, y) = 255
            ImageData(QuickVal + 1, y) = 255
            ImageData(QuickVal, y) = 255
        End If
    Next y
        If x Mod 20 = 0 Then SetProgBarVal x
    Next x
    SetImageData
End Sub

'Standard ordered dither.  Coefficients derived from http://en.wikipedia.org/wiki/Ordered_dithering
' As you'd expect, this routine is quite fast, but it gives an image that "obviously dithered" Windows 3.1 look.
Public Sub MenuBWOrderedDither()
    Dim r As Integer, g As Integer, B As Integer
    Dim DitherTable(1 To 4, 1 To 4) As Byte
    Dim dTable(0 To 765) As Byte
    DitherTable(1, 1) = 1
    DitherTable(1, 2) = 9
    DitherTable(1, 3) = 3
    DitherTable(1, 4) = 11
    DitherTable(2, 1) = 13
    DitherTable(2, 2) = 5
    DitherTable(2, 3) = 15
    DitherTable(2, 4) = 7
    DitherTable(3, 1) = 4
    DitherTable(3, 2) = 12
    DitherTable(3, 3) = 2
    DitherTable(3, 4) = 10
    DitherTable(4, 1) = 16
    DitherTable(4, 2) = 8
    DitherTable(4, 3) = 14
    DitherTable(4, 4) = 6
    For x = 0 To 765
        dTable(x) = 1 + (x \ 3) \ 17
    Next x

    Dim TC As Integer
    Dim TX As Integer, TY As Integer
    Dim QuickVal As Long
    Message "Converting to black and white..."
    SetProgBarMax PicWidthL
    For x = 0 To PicWidthL
        QuickVal = x * 3
    For y = 0 To PicHeightL
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        B = ImageData(QuickVal, y)
        TC = dTable(r + g + B)
        TX = 1 + (x Mod 4)
        TY = 1 + (y Mod 4)
        If TC < DitherTable(TX, TY) Then
            ImageData(QuickVal + 2, y) = 0
            ImageData(QuickVal + 1, y) = 0
            ImageData(QuickVal, y) = 0
        Else
            ImageData(QuickVal + 2, y) = 255
            ImageData(QuickVal + 1, y) = 255
            ImageData(QuickVal, y) = 255
        End If
    Next y
        If x Mod 20 = 0 Then SetProgBarVal x
    Next x
    SetImageData
End Sub

'Standard error diffusion using my self-written diffusion algorithm (hacked for 1-bit usage)
'Includes a rounding function for optimum diffusing.  It's the fastest diffusion algorithm
'I think it's possible to write within VB, but it destroys all details.
'Works well enough for me, though. :)
Public Sub MenuBWDiffusionDither()
    Message "Converting image to black and white..."
    SetProgBarMax PicHeightL
    'Track the error value
    Dim EV As Long
    'Lots of color values
    Dim CC As Long, CurrentColor As Long
    Dim r1 As Long, g1 As Long, b1 As Long
    Dim QuickVal As Long
    For y = 0 To PicHeightL
    For x = 0 To PicWidthL
        'Strip out pixel data
        QuickVal = x * 3
        r1 = ImageData(QuickVal + 2, y)
        g1 = ImageData(QuickVal + 1, y)
        b1 = ImageData(QuickVal, y)
        'Grayscale value of the pixel
        CurrentColor = (222 * r1 + 707 * g1 + 71 * b1) \ 1000
        'Add the running error to the pixel
        CurrentColor = CurrentColor + EV
        'Convert this pixel to black or white, depending on its "luminance"
        CC = Int((CSng(CurrentColor) / 255) + 0.5) * 255
        'Error value is the difference between the two
        EV = (CurrentColor - CC)
        'Make all the pixels the same value
        ByteMeL CC
        ImageData(QuickVal + 2, y) = CC
        ImageData(QuickVal + 1, y) = CC
        ImageData(QuickVal, y) = CC
        'Reduce color bleeding by reducing the error as we go (optional)
        'EV = (EV \ 3) * 2
    Next x
        'Discard the error at the end of every line
        EV = 0
        If y Mod 20 = 0 Then SetProgBarVal y
    Next y
    SetImageData
End Sub

'A freakish converter that uses a nice every-other sampling routine.  I like it.
Public Sub MenuBWImpressionist()
    Dim r As Integer, g As Integer, B As Integer
    Dim tR As Long
    Dim QuickVal As Long
    Message "Converting to black and white..."
    SetProgBarMax PicWidthL
    For x = 0 To PicWidthL
        QuickVal = x * 3
    For y = 0 To PicHeightL
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        B = ImageData(QuickVal, y)
        tR = Int((222 * CLng(r) + 707 * CLng(g) + 71 * CLng(B)) \ 1000)
        If tR < 64 Then
            r = 0
            g = 0
            B = 0
        ElseIf tR >= 64 And tR < 128 Then
            r = 255
            g = 255
            B = 255
        ElseIf tR >= 128 And tR < 192 Then
            r = 0
            g = 0
            B = 0
        ElseIf tR >= 192 Then
            r = 255
            g = 255
            B = 255
        End If
        ImageData(QuickVal + 2, y) = r
        ImageData(QuickVal + 1, y) = g
        ImageData(QuickVal, y) = B
    Next y
        If x Mod 20 = 0 Then SetProgBarVal x
    Next x
    SetImageData
End Sub

'This is based off Manuel Augusto Santos's "Enhanced Dither" algoritm; a link to his original code can be found on the
' "About" page of this project
Public Sub MenuBWEnhancedDither()
  Dim r As Long, g As Long, B As Long
  Dim ErrorDif As Long, nColors As Long
  Dim gray As Long
  Dim AveColor As Long
    Message "Preparing conversion data..."
    SetProgBarMax PicHeightL
    Dim LookUp(0 To 765) As Integer
  For x = 0 To 765
    LookUp(x) = x \ 3
  Next x
  GetImageData
  AveColor = 0
  nColors = 0
  
  Dim QuickVal As Long
  
  For x = 0 To PicWidthL
    QuickVal = x * 3
  For y = 0 To PicHeightL
      B = ImageData(QuickVal, y)
      g = ImageData(QuickVal + 1, y)
      r = ImageData(QuickVal + 2, y)
      gray = LookUp(r + g + B)
      AveColor = AveColor + gray
      nColors = nColors + 1
  Next y
  Next x
  AveColor = AveColor \ nColors
  
    ReDim tData(0 To PicWidthL * 3 + 3, 0 To PicHeightL + 1)
  
    Message "Converting image to black and white..."
  For y = 0 To PicHeightL
    For x = 0 To PicWidthL
        QuickVal = x * 3
      If (x > 0) And (y > 0) And (x < PicWidthL - 1) And (y < PicHeightL - 1) Then
        B = CLng(ImageData((x - 1) * 3, y - 1)) + CLng(ImageData((x - 1) * 3, y)) + CLng(ImageData((x - 1) * 3, y + 1)) + CLng(ImageData(QuickVal, y - 1)) + CLng(ImageData(QuickVal, y + 1)) + CLng(ImageData((x + 1) * 3, y - 1)) + CLng(ImageData((x + 1) * 3, y)) + CLng(ImageData((x + 1) * 3, y + 1))
        g = CLng(ImageData((x - 1) * 3 + 1, y - 1)) + CLng(ImageData((x - 1) * 3 + 1, y)) + CLng(ImageData((x - 1) * 3 + 1, y + 1)) + CLng(ImageData(QuickVal + 1, y - 1)) + CLng(ImageData(QuickVal + 1, y + 1)) + CLng(ImageData((x + 1) * 3 + 1, y - 1)) + CLng(ImageData((x + 1) * 3 + 1, y)) + CLng(ImageData((x + 1) * 3 + 1, y + 1))
        r = CLng(ImageData((x - 1) * 3 + 2, y - 1)) + CLng(ImageData((x - 1) * 3 + 2, y)) + CLng(ImageData((x - 1) * 3 + 2, y + 1)) + CLng(ImageData(QuickVal + 2, y - 1)) + CLng(ImageData(QuickVal + 2, y + 1)) + CLng(ImageData((x + 1) * 3 + 2, y - 1)) + CLng(ImageData((x + 1) * 3 + 2, y)) + CLng(ImageData((x + 1) * 3 + 2, y + 1))
        B = (10 * CLng(ImageData(QuickVal, y)) - B) \ 2
        g = (10 * CLng(ImageData(QuickVal + 1, y)) - g) \ 2
        r = (10 * CLng(ImageData(QuickVal + 2, y)) - r) \ 2
        ByteMeL r
        ByteMeL g
        ByteMeL B
      Else
        B = ImageData(QuickVal, y)
        g = ImageData(QuickVal + 1, y)
        r = ImageData(QuickVal + 2, y)
      End If
      gray = LookUp(r + g + B)
      gray = gray + ErrorDif
      ByteMeL gray
      If gray < AveColor Then nColors = 0 Else nColors = 255
      ErrorDif = (gray - nColors) \ 4
      tData(QuickVal, y) = nColors
      tData(QuickVal + 1, y) = nColors
      tData(QuickVal + 2, y) = nColors
    Next x
    If y Mod 20 = 0 Then SetProgBarVal y
  Next y
  TransferImageData
  SetImageData
  SetProgBarVal 0
End Sub

'A modified version of Floyd-Steinberg dithering.  Very good results, but slow.  I tried to fix this by
' using the mother of all look-up tables, but there are limits to what VB can do with processes this complex :)
Public Sub MenuBWFloydSteinberg()
    SetProgBarMax PicHeightL
    Dim EV As Long
    Dim CurrentColor As Long
    Dim r1 As Long, g1 As Long, b1 As Long
    Dim TV As Integer
    
    'One freaking huge look-up table; note that I apply some reduction to the error values - this is done to limit color
    ' bleeding and produce a slightly sharper picture.  This makes it not the "true" Floyd-Steinberg approach, but since
    ' few people (if any) could pick out a visual a difference I stick with the more aesthetically pleasing mechanism.
    Message "Generating look-up tables..."
    Dim LookUps(0 To 3, -127 To 127, 0 To 255) As Integer
    For x = -127 To 127
    For y = 0 To 255
        TV = (CSng(x) * 0.4375) \ 1.35 + y
        ByteMe TV
        LookUps(0, x, y) = TV
        TV = (CSng(x) * 0.1875) \ 1.35 + y
        ByteMe TV
        LookUps(1, x, y) = TV
        TV = (CSng(x) * 0.3125) \ 1.35 + y
        ByteMe TV
        LookUps(2, x, y) = TV
        TV = (CSng(x) * 0.0625) \ 1.35 + y
        ByteMe TV
        LookUps(3, x, y) = TV
    Next y
    Next x
    
    'Applying the wonderful look-up values
    Message "Converting image to black and white..."
    Dim QuickVal As Long, QuickValUp As Long, QuickValDown As Long
    For y = 0 To PicHeightL
    For x = 0 To PicWidthL
        QuickVal = x * 3
        QuickValUp = (x + 1) * 3
        QuickValDown = (x - 1) * 3
        r1 = ImageData(QuickVal + 2, y)
        g1 = ImageData(QuickVal + 1, y)
        b1 = ImageData(QuickVal, y)
        CurrentColor = (222 * r1 + 707 * g1 + 71 * b1) \ 1000
        If CurrentColor < 128 Then
            EV = CurrentColor
            ImageData(QuickVal + 2, y) = 0
            ImageData(QuickVal + 1, y) = 0
            ImageData(QuickVal, y) = 0
        Else
            EV = CurrentColor - 255
            ImageData(QuickVal + 2, y) = 255
            ImageData(QuickVal + 1, y) = 255
            ImageData(QuickVal, y) = 255
        End If
        'Additional error reduction may be used here at the programmer's discretion.
        'EV = (EV * 2) \ 3
        'Diffuse right
        If x <> PicWidthL Then
            ImageData(QuickValUp + 2, y) = LookUps(0, EV, ImageData(QuickValUp + 2, y))
            ImageData(QuickValUp + 1, y) = LookUps(0, EV, ImageData(QuickValUp + 1, y))
            ImageData(QuickValUp, y) = LookUps(0, EV, ImageData(QuickValUp, y))
        End If
        'Diffuse down and left
        If ((y + 1) <= PicHeightL) Then
        If x <> 0 Then
            ImageData(QuickValDown + 2, y + 1) = LookUps(1, EV, ImageData(QuickValDown + 2, y + 1))
            ImageData(QuickValDown + 1, y + 1) = LookUps(1, EV, ImageData(QuickValDown + 1, y + 1))
            ImageData(QuickValDown, y + 1) = LookUps(1, EV, ImageData(QuickValDown, y + 1))
        End If
        'Diffuse down
            ImageData(QuickVal + 2, y + 1) = LookUps(2, EV, ImageData(QuickVal + 2, y + 1))
            ImageData(QuickVal + 1, y + 1) = LookUps(2, EV, ImageData(QuickVal + 1, y + 1))
            ImageData(QuickVal, y + 1) = LookUps(2, EV, ImageData(QuickVal, y + 1))
        'Diffuse down and right
        If x <> PicWidthL Then
            ImageData(QuickValUp + 2, y + 1) = LookUps(3, EV, ImageData(QuickValUp + 2, y + 1))
            ImageData(QuickValUp + 1, y + 1) = LookUps(3, EV, ImageData(QuickValUp + 1, y + 1))
            ImageData(QuickValUp, y + 1) = LookUps(3, EV, ImageData(QuickValUp, y + 1))
        End If
        End If
    Next x
        EV = 0
        If y Mod 20 = 0 Then SetProgBarVal y
    Next y
    'Clear out the memory as best we can within VB
    Erase LookUps
    
    SetImageData
End Sub
