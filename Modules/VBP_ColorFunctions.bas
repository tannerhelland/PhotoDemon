Attribute VB_Name = "Color_Functions"
'***************************************************************************
'Miscellaneous Color Functions
'Copyright ©2013-2014 by Tanner Helland
'Created: 13/June/13
'Last updated: 13/August/13
'Last update: added XYZ and CieLAB color conversions
'
'Many of these functions are older than the create date above, but I did not organize them into a consistent module
' until June 2014.  This module is now used to store all the random bits of specialized color processing code
' required by the program.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Convert a system color (such as "button face" or "inactive window") to a literal RGB value
Private Declare Function OleTranslateColor Lib "olepro32" (ByVal oColor As OLE_COLOR, ByVal HPALETTE As Long, ByRef cColorRef As Long) As Long

'Given a DIB, fill a Single-type array with a L*a*b* representation of the image
Public Function convertEntireDIBToLabColor(ByRef srcDIB As pdDIB, ByRef dstArray() As Single) As Boolean

    'This only works on 24bpp images; exit prematurely on 32bpp encounters
    If srcDIB.getDIBColorDepth = 32 Then
        convertEntireDIBToLabColor = False
        Exit Function
    End If

    'Redim the destination array to proper dimensions
    ReDim dstArray(0 To srcDIB.getDIBArrayWidth, 0 To srcDIB.getDIBHeight) As Single
    
    'Request a pointer to the source dib
    Dim tmpSA As SAFEARRAY2D
    prepSafeArray tmpSA, srcDIB
    
    Dim ImageData() As Byte
    CopyMemory ByVal VarPtrArray(ImageData()), VarPtr(tmpSA), 4
    
    'Iterate through the image, converting colors as we go
    Dim x As Long, y As Long, finalX As Long, finalY As Long, QuickX As Long
    
    finalX = srcDIB.getDIBWidth - 1
    finalY = srcDIB.getDIBHeight - 1
    
    Dim r As Long, g As Long, b As Long
    Dim labL As Double, labA As Double, labB As Double
    
    For x = 0 To finalX
        QuickX = x * 3
    For y = 0 To finalY
    
        'Get the source pixel color values
        r = ImageData(QuickX + 2, y)
        g = ImageData(QuickX + 1, y)
        b = ImageData(QuickX, y)
        
        'Convert the color to the L*a*b* color space
        RGBtoLAB r, g, b, labL, labA, labB
        
        'Store the L*a*b* values
        dstArray(QuickX, y) = labL
        dstArray(QuickX + 1, y) = labA
        dstArray(QuickX + 2, y) = labB
    
    Next y
    Next x
    
    'With our work complete, point ImageData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(ImageData), 0&, 4
    Erase ImageData
    
    convertEntireDIBToLabColor = True

End Function

'Present the user with a color selection dialog.  At present, this is just a thin wrapper to the stock Windows color
' selector, but in the future it will link to a custom PhotoDemon one.
' INPUTS:  1) a Long-type variable that will receive the new color
'          2) an optional initial color
'
' OUTPUTS: 1) TRUE if OK was pressed, FALSE for Cancel
Public Function showColorDialog(ByRef colorReceive As Long, Optional ByVal initialColor As Long = vbWhite, Optional ByRef callingControl As colorSelector) As Boolean

    'Uncomment this code to use the system color selector
    'Dim retColor As Long
    'Dim CD1 As cCommonDialog
    'Set CD1 = New cCommonDialog
    'retColor = initialColor
    '
    'If CD1.VBChooseColor(retColor, True, True, False, getModalOwner(True).hWnd) Then
    '    colorReceive = retColor
    '    showColorDialog = True
    'Else
    '    showColorDialog = False
    'End If
    
    'As of November 2014, PhotoDemon has its own color selector!
    If choosePDColor(initialColor, colorReceive, callingControl) = vbOK Then
        showColorDialog = True
    Else
        showColorDialog = False
    End If
    
End Function

'Given the number of colors in an image (as supplied by getQuickColorCount, below), return the highest color depth
' that includes all those colors and is supported by PhotoDemon (1/4/8/24/32)
Public Function getColorDepthFromColorCount(ByVal srcColors As Long, ByRef refDIB As pdDIB) As Long
    
    If srcColors <= 256 Then
        If srcColors > 16 Then
            getColorDepthFromColorCount = 8
        Else
            
            'FreeImage only supports the writing of 4bpp and 1bpp images if they are grayscale. Thus, only
            ' mark images as 4bpp or 1bpp if they are gray/b&w - otherwise, consider them 8bpp indexed color.
            If (srcColors > 2) Then
                                
                If g_IsImageGray Then getColorDepthFromColorCount = 4 Else getColorDepthFromColorCount = 8
            
            'If there are only two colors, see if they are black and white, other shades of gray, or colors.
            ' Mark the color depth as 1bpp, 4bpp, or 8bpp respectively.
            Else
                If g_IsImageMonochrome Then
                    getColorDepthFromColorCount = 1
                Else
                    If g_IsImageGray Then getColorDepthFromColorCount = 4 Else getColorDepthFromColorCount = 8
                End If
            End If
            
        End If
    Else
        If refDIB.getDIBColorDepth = 24 Then
            getColorDepthFromColorCount = 24
        Else
            getColorDepthFromColorCount = 32
        End If
    End If

End Function

'When images are loaded, this function is used to quickly determine the image's color count. It stops once 257 is reached,
' as at that point the program will automatically treat the image as 24 or 32bpp (contingent on presence of an alpha channel).
Public Function getQuickColorCount(ByVal srcImage As pdImage, Optional ByVal imageID As Long = -1) As Long
    
    Message "Verifying image color count..."
    
    'Mark the image ID to the global tracking variable
    g_LastImageScanned = imageID
    
    'Retrieve a composited version of the target image
    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB
    srcImage.getCompositedImage tmpDIB, True
    
    'Create a local array and point it at the pixel data we want to operate on
    Dim ImageData() As Byte
    Dim tmpSA As SAFEARRAY2D
    prepSafeArray tmpSA, tmpDIB
    CopyMemory ByVal VarPtrArray(ImageData()), VarPtr(tmpSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, finalX As Long, finalY As Long
    finalX = srcImage.Width - 1
    finalY = srcImage.Height - 1
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long
    qvDepth = (srcImage.getActiveDIB().getDIBColorDepth) \ 8
    
    'This array will track whether or not a given color has been detected in the image. (I don't know if powers of two
    ' are allocated more efficiently, but it doesn't hurt to stick to that rule.)
    Dim UniqueColors() As Long
    ReDim UniqueColors(0 To 511) As Long
    
    Dim i As Long
    For i = 0 To 255
        UniqueColors(i) = -1
    Next i
    
    'Total number of unique colors counted so far
    Dim totalCount As Long
    totalCount = 0
    
    'Finally, a bunch of variables used in color calculation
    Dim r As Long, g As Long, b As Long
    Dim chkValue As Long
    Dim colorFound As Boolean
        
    'Apply the filter
    For x = 0 To finalX
        QuickVal = x * qvDepth
    For y = 0 To finalY
        
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        
        chkValue = RGB(r, g, b)
        colorFound = False
        
        'Now, loop through the colors we've accumulated thus far and compare this entry against each of them.
        For i = 0 To totalCount
            If UniqueColors(i) = chkValue Then
                colorFound = True
                Exit For
            End If
        Next i
        
        'If colorFound is still false, store this value in the array and increment our color counter
        If Not colorFound Then
            UniqueColors(totalCount) = chkValue
            totalCount = totalCount + 1
        End If
        
        'If the image has more than 256 colors, treat it as 24/32 bpp
        If totalCount > 256 Then Exit For
        
    Next y
        If totalCount > 256 Then Exit For
    Next x
        
    'With our work complete, point ImageData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(ImageData), 0&, 4
    Erase ImageData
    
    'If the image contains only two colors, check to see if they are pure black and pure white. If so, mark
    ' a global flag accordingly and exit (to save a little bit of extra processing time)
    g_IsImageMonochrome = False
    
    If totalCount = 2 Then
    
        r = ExtractR(UniqueColors(0))
        g = ExtractG(UniqueColors(0))
        b = ExtractB(UniqueColors(0))
        
        If ((r = 0) And (g = 0) And (b = 0)) Or ((r = 255) And (g = 255) And (b = 255)) Then
            
            r = ExtractR(UniqueColors(1))
            g = ExtractG(UniqueColors(1))
            b = ExtractB(UniqueColors(1))
            
            If ((r = 0) And (g = 0) And (b = 0)) Or ((r = 255) And (g = 255) And (b = 255)) Then
                g_IsImageMonochrome = True
                Erase UniqueColors
                g_LastColorCount = totalCount
                getQuickColorCount = totalCount
                Exit Function
            End If
            
        End If
        
    End If
        
    'If we've made it this far, the color count has a maximum value of 257.
    ' If it is less than 257, analyze it to see if it contains all gray values.
    If totalCount <= 256 Then
    
        g_IsImageGray = True
    
        'Loop through all available colors
        For i = 0 To totalCount - 1
        
            r = ExtractR(UniqueColors(i))
            g = ExtractG(UniqueColors(i))
            b = ExtractB(UniqueColors(i))
            
            'If any of the components do not match, this is not a grayscale image
            If (r <> g) Or (g <> b) Or (r <> b) Then
                g_IsImageGray = False
                Exit For
            End If
            
        Next i
    
    'If the image contains more than 256 colors, it is not grayscale
    Else
        g_IsImageGray = False
    End If
    
    Erase UniqueColors
    
    g_LastColorCount = totalCount
    getQuickColorCount = totalCount
        
End Function

'Given an OLE color, return an RGB
Public Function ConvertSystemColor(ByVal colorRef As OLE_COLOR) As Long
    
    'OleTranslateColor returns -1 if it fails; if that happens, default to white
    If OleTranslateColor(colorRef, 0, ConvertSystemColor) Then
        ConvertSystemColor = RGB(255, 255, 255)
    End If
    
End Function

'Extract the red, green, or blue value from an RGB() Long
Public Function ExtractR(ByVal currentColor As Long) As Integer
    ExtractR = currentColor Mod 256
End Function

Public Function ExtractG(ByVal currentColor As Long) As Integer
    ExtractG = (currentColor \ 256) And 255
End Function

Public Function ExtractB(ByVal currentColor As Long) As Integer
    ExtractB = (currentColor \ 65536) And 255
End Function

'Blend byte1 w/ byte2 based on mixRatio. mixRatio is expected to be a value between 0 and 1.
Public Function BlendColors(ByVal Color1 As Byte, ByVal Color2 As Byte, ByRef mixRatio As Double) As Byte
    BlendColors = ((1 - mixRatio) * Color1) + (mixRatio * Color2)
End Function

'This function will return the luminance value of an RGB triplet.  Note that the value will be in the [0,255] range instead
' of the usual [0,1.0] one.
Public Function getLuminance(ByVal r As Long, ByVal g As Long, ByVal b As Long) As Long
    Dim Max As Long, Min As Long
    Max = Max3Int(r, g, b)
    Min = Min3Int(r, g, b)
    getLuminance = (Max + Min) \ 2
End Function

'HSL <-> RGB conversion routines
Public Sub tRGBToHSL(r As Long, g As Long, b As Long, h As Double, s As Double, l As Double)
    
    Dim Max As Double, Min As Double
    Dim Delta As Double
    Dim rR As Double, rG As Double, rB As Double
    
    rR = r / 255
    rG = g / 255
    rB = b / 255

    'Note: HSL are calculated in the following ranges:
    ' Hue: [-1,5]
    ' Saturation: [0,1] (Note that if saturation = 0, hue is technically undefined)
    ' Lightness: [0,1]

    Max = Max3Float(rR, rG, rB)
    Min = Min3Float(rR, rG, rB)
        
    'Calculate luminance
    l = (Max + Min) / 2
        
    'If the maximum and minimum are identical, this image is gray, meaning it has no saturation and an undefined hue.
    If Max = Min Then
        s = 0
        h = 0
    Else
        
        Delta = Max - Min
        
        'Calculate saturation
        If l <= 0.5 Then
            s = Delta / (Max + Min)
        Else
            s = Delta / (2 - Max - Min)
        End If
        
        'Calculate hue
        
        If rR = Max Then
            h = (rG - rB) / Delta    '{Resulting color is between yellow and magenta}
        ElseIf rG = Max Then
            h = 2 + (rB - rR) / Delta '{Resulting color is between cyan and yellow}
        ElseIf rB = Max Then
            h = 4 + (rR - rG) / Delta '{Resulting color is between magenta and cyan}
        End If
        
        'If you prefer hue in the [0,360] range instead of [-1, 5] you can use this code
        'h = h * 60
        'If h < 0 Then h = h + 360

    End If

    'Tanner's final note: if byte values are preferred to floating-point, this code will return hue on [0,240],
    ' saturation on [0,255], and luminance on [0,255]
    'H = Int(H * 40 + 40)
    'S = Int(S * 255)
    'L = Int(L * 255)
    
End Sub

'Convert HSL values to RGB values
Public Sub tHSLToRGB(h As Double, s As Double, l As Double, r As Long, g As Long, b As Long)

    Dim rR As Double, rG As Double, rB As Double
    Dim Min As Double, Max As Double

    'Unsaturated pixels do not technically have hue - they only have luminance
    If s = 0 Then
        rR = l: rG = l: rB = l
    Else
        If l <= 0.5 Then
            Min = l * (1 - s)
        Else
            Min = l - s * (1 - l)
        End If
      
        Max = 2 * l - Min
      
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
   
   r = rR * 255
   g = rG * 255
   b = rB * 255
   
   'Failsafe added 29 August '12
   'This should never return RGB values > 255, but it doesn't hurt to make sure.
   If r > 255 Then r = 255
   If g > 255 Then g = 255
   If b > 255 Then b = 255
   
End Sub

'Convert [0,255] RGB values to [0,1] HSV values, with thanks to easyrgb.com for the conversion math
Public Sub RGBtoHSV(ByVal r As Long, ByVal g As Long, ByVal b As Long, ByRef h As Double, ByRef s As Double, ByRef v As Double)

    Dim fR As Double, fG As Double, fB As Double
    fR = r / 255
    fG = g / 255
    fB = b / 255

    Dim var_Min As Double, var_Max As Double, del_Max As Double
    var_Min = Min3Float(fR, fG, fB)
    var_Max = Max3Float(fR, fG, fB)
    del_Max = var_Max - var_Min
    
    'Value is easy to calculate - it's the largest of R/G/B
    v = var_Max

    'If the max and min are the same, this is a gray pixel
    If del_Max = 0 Then
        h = 0
        s = 0
        
    'If max and min vary, we can calculate a hue component
    Else
    
        s = del_Max / var_Max
        
        Dim del_R As Double, del_G As Double, del_B As Double
        del_R = (((var_Max - fR) / 6) + (del_Max / 2)) / del_Max
        del_G = (((var_Max - fG) / 6) + (del_Max / 2)) / del_Max
        del_B = (((var_Max - fB) / 6) + (del_Max / 2)) / del_Max

        If fR = var_Max Then
            h = del_B - del_G
        ElseIf fG = var_Max Then
            h = (1 / 3) + del_R - del_B
        Else
            h = (2 / 3) + del_G - del_R
        End If

        If h < 0 Then h = h + 1
        If h > 1 Then h = h - 1

    End If

End Sub

'Convert [0,1] HSV values to [0,255] RGB values, with thanks to easyrgb.com for the conversion math
Public Sub HSVtoRGB(ByRef h As Double, ByRef s As Double, ByRef v As Double, ByRef r As Long, ByRef g As Long, ByRef b As Long)

    'If saturation is 0, RGB are calculated identically
    If s = 0 Then
        r = v * 255
        g = v * 255
        b = v * 255
    
    'If saturation is not 0, we have to calculate RGB independently
    Else
       
        Dim var_H As Double
        var_H = h * 6
        
        'To keep our math simple, limit hue to [0, 5.9999999]
        If var_H >= 6 Then var_H = 0
        
        Dim var_I As Long
        var_I = Int(var_H)
        
        Dim var_1 As Double, var_2 As Double, var_3 As Double
        var_1 = v * (1 - s)
        var_2 = v * (1 - s * (var_H - var_I))
        var_3 = v * (1 - s * (1 - (var_H - var_I)))

        Dim var_R As Double, var_G As Double, var_B As Double

        Select Case var_I
        
            Case 0
                var_R = v
                var_G = var_3
                var_B = var_1
                
            Case 1
                var_R = var_2
                var_G = v
                var_B = var_1
                
            Case 2
                var_R = var_1
                var_G = v
                var_B = var_3
                
            Case 3
                var_R = var_1
                var_G = var_2
                var_B = v
            
            Case 4
                var_R = var_3
                var_G = var_1
                var_B = v
                
            Case Else
                var_R = v
                var_G = var_1
                var_B = var_2
                
        End Select

        r = var_R * 255
        g = var_G * 255
        b = var_B * 255
                
    End If

End Sub

'A heavily modified RGB to HSV transform, courtesy of http://lolengine.net/blog/2013/01/13/fast-rgb-to-hsv.
' Note that the code assumes RGB values already in the [0, 1] range, and it will return HSV values in the [0, 1] range.
Public Sub fRGBtoHSV(ByVal r As Double, ByVal g As Double, ByVal b As Double, ByRef h As Double, ByRef s As Double, ByRef v As Double)

    Dim K As Double, tmpSwap As Double, chroma As Double
    
    If (g < b) Then
        tmpSwap = b
        b = g
        g = tmpSwap
        K = -1
    End If
    
    If (r < g) Then
        tmpSwap = g
        g = r
        r = tmpSwap
        K = -(2 / 6) - K
    End If
    
    chroma = r - fMin(g, b)
    h = Abs(K + (g - b) / (6 * chroma + 0.0000001))
    s = chroma / (r + 0.00000001)
    v = r
    
End Sub

'Convert [0,1] HSV values to [0,1] RGB values, with thanks to easyrgb.com for the conversion math
Public Sub fHSVtoRGB(ByRef h As Double, ByRef s As Double, ByRef v As Double, ByRef r As Double, ByRef g As Double, ByRef b As Double)

    'If saturation is 0, RGB are calculated identically
    If s = 0 Then
        r = v
        g = v
        b = v
        Exit Sub
    
    'If saturation is not 0, we have to calculate RGB independently
    Else
       
        Dim var_H As Double
        var_H = h * 6
        
        'To keep our math simple, limit hue to [0, 5.9999999]
        If var_H >= 6 Then var_H = 0
        
        Dim var_I As Long
        var_I = Int(var_H)
        
        Dim var_1 As Double, var_2 As Double, var_3 As Double
        var_1 = v * (1 - s)
        var_2 = v * (1 - s * (var_H - var_I))
        var_3 = v * (1 - s * (1 - (var_H - var_I)))
        
        Select Case var_I
        
            Case 0
                r = v
                g = var_3
                b = var_1
                
            Case 1
                r = var_2
                g = v
                b = var_1
                
            Case 2
                r = var_1
                g = v
                b = var_3
                
            Case 3
                r = var_1
                g = var_2
                b = v
            
            Case 4
                r = var_3
                g = var_1
                b = v
                
            Case Else
                r = v
                g = var_1
                b = var_2
                
        End Select
                
    End If

End Sub

'This function is just a thin wrapper to RGBtoXYZ and XYZtoLAB.  There is no direct conversion from RGB to CieLAB.
Public Sub RGBtoLAB(ByVal r As Long, ByVal g As Long, ByVal b As Long, ByRef labL As Double, ByRef labA As Double, ByRef labB As Double)

    Dim x As Double, y As Double, z As Double
    RGBtoXYZ r, g, b, x, y, z
    XYZtoLab x, y, z, labL, labA, labB

End Sub

'Convert RGB to XYZ space, using an sRGB conversion and the assumption of a D65 (e.g. color temperature of 6500k) illuminant
' Formula adopted from http://www.easyrgb.com/index.php?X=MATH&H=02#text2
Public Sub RGBtoXYZ(ByVal r As Long, ByVal g As Long, ByVal b As Long, ByRef x As Double, ByRef y As Double, ByRef z As Double)

    'Normalize RGB to [0, 1]
    Dim rFloat As Double, gFloat As Double, bFloat As Double
    rFloat = r / 255
    gFloat = g / 255
    bFloat = b / 255
    
    'Convert RGB values to the sRGB color space
    If rFloat > 0.04045 Then
        rFloat = ((rFloat + 0.055) / (1.055)) ^ 2.2
    Else
        rFloat = rFloat / 12.92
    End If
    
    If gFloat > 0.04045 Then
        gFloat = ((gFloat + 0.055) / (1.055)) ^ 2.2
    Else
        gFloat = gFloat / 12.92
    End If
    
    If bFloat > 0.04045 Then
        bFloat = ((bFloat + 0.055) / (1.055)) ^ 2.2
    Else
        bFloat = bFloat / 12.92
    End If
    
    'Calculate XYZ using D65 correction
    x = rFloat * 0.4124 + gFloat * 0.3576 + bFloat * 0.1805
    y = rFloat * 0.2126 + gFloat * 0.7152 + bFloat * 0.0722
    z = rFloat * 0.0193 + gFloat * 0.1192 + bFloat * 0.9505
    
End Sub

'Convert an XYZ color to CIELab.  As with the original XYZ calculation, D65 is assumed.
' Formula adopted from http://www.easyrgb.com/index.php?X=MATH&H=07#text7, with minor changes by me (not re-applying D65 values until after
'  fXYZ has been calculated)
Public Sub XYZtoLab(ByVal x As Double, ByVal y As Double, ByVal z As Double, ByRef l As Double, ByRef a As Double, ByRef b As Double)
    l = 116 * fXYZ(y) - 16
    a = 500 * (fXYZ(x / 0.9505) - fXYZ(y))
    b = 200 * (fXYZ(y) - fXYZ(z / 1.089))
End Sub

Private Function fXYZ(ByVal t As Double) As Double
    If t > 0.008856 Then
        fXYZ = t ^ (1 / 3)
    Else
        fXYZ = (7.787 * t) + (16 / 116)
    End If
End Function

'Return the minimum of two floating-point values
Private Function fMin(x As Double, y As Double) As Double
    If x > y Then fMin = y Else fMin = x
End Function

'Return the maximum of two floating-point values
Private Function fMax(x As Double, y As Double) As Double
    If x < y Then fMax = y Else fMax = x
End Function

'Return the maximum of three floating point values
Private Function fMax3(rR As Double, rG As Double, rB As Double) As Double
   If (rR > rG) Then
      If (rR > rB) Then
         fMax3 = rR
      Else
         fMax3 = rB
      End If
   Else
      If (rB > rG) Then
         fMax3 = rB
      Else
         fMax3 = rG
      End If
   End If
End Function

'Return the minimum of three floating point values
Private Function fMin3(rR As Double, rG As Double, rB As Double) As Double
   If (rR < rG) Then
      If (rR < rB) Then
         fMin3 = rR
      Else
         fMin3 = rB
      End If
   Else
      If (rB < rG) Then
         fMin3 = rB
      Else
         fMin3 = rG
      End If
   End If
End Function


