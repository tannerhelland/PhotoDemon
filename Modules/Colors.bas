Attribute VB_Name = "Colors"
'***************************************************************************
'Miscellaneous Color Functions
'Copyright 2013-2026 by Tanner Helland
'Created: 13/June/13
'Last updated: 28/February/19
'Last update: add support for retrieving colors by SVG color name
'
'Many of these functions are older than the create date above, but I did not organize them into a consistent module
' until June 2013.  This module is now used to store all the random bits of specialized color processing code
' required by the program.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'PhotoDemon tries to support a variety of textual color representations.  Not all of these are implemented at present.
' (TODO: get all formats working)
Public Enum PD_ColorStringType
    ColorInvalid = -1
    ColorUnknown = 0
    ColorHex = 1
    ColorRGB = 2
    ColorRGBA = 3
    ColorHSL = 4
    ColorHSLA = 5
    ColorNamed = 6
End Enum

#If False Then
    Private Const ColorInvalid = -1, ColorUnknown = 0, ColorHex = 1, ColorRGB = 2, ColorRGBA = 3, ColorHSL = 4, ColorHSLA = 5, ColorNamed = 6
#End If

Public Enum PD_BlendMode
    BM_Normal = 0
    BM_Darken = 1
    BM_Multiply = 2
    BM_ColorBurn = 3
    BM_LinearBurn = 4
    BM_Lighten = 5
    BM_Screen = 6
    BM_ColorDodge = 7
    BM_LinearDodge = 8
    BM_Overlay = 9
    BM_SoftLight = 10
    BM_HardLight = 11
    BM_VividLight = 12
    BM_LinearLight = 13
    BM_PinLight = 14
    BM_HardMix = 15
    BM_Difference = 16
    BM_Exclusion = 17
    BM_Subtract = 18
    BM_Divide = 19
    BM_Hue = 20
    BM_Saturation = 21
    BM_Color = 22
    BM_Luminosity = 23
    BM_GrainExtract = 24
    BM_GrainMerge = 25
    BM_Erase = 26
    BM_Behind = 27
    BM_Overwrite = 28
End Enum

#If False Then
    Const BM_Normal = 0, BM_Darken = 1, BM_Multiply = 2, BM_ColorBurn = 3, BM_LinearBurn = 4
    Const BM_Lighten = 5, BM_Screen = 6, BM_ColorDodge = 7, BM_LinearDodge = 8, BM_Overlay = 9
    Const BM_SoftLight = 10, BM_HardLight = 11, BM_VividLight = 12, BM_LinearLight = 13, BM_PinLight = 14
    Const BM_HardMix = 15, BM_Difference = 16, BM_Exclusion = 17, BM_Subtract = 18, BM_Divide = 19
    Const BM_Hue = 20, BM_Saturation = 21, BM_Color = 22, BM_Luminosity = 23, BM_GrainExtract = 24
    Const BM_GrainMerge = 25, BM_Erase = 26, BM_Behind = 27, BM_Overwrite = 28
#End If

'PD supports the notion "alpha modes", including "inheritance," where a layer "inherits" the alpha of a layer beneath it,
' (https://userbase.kde.org/Krita/Tutorial_2#Inherit_Alpha_.28alpha_.3D_transparency.29).  I may or may not choose to
' address masking via this property as well... I'm still deciding the best way to do it.
'
'Also, alpha locking is handled via mode.
Public Enum PD_AlphaMode
    AM_Normal = 0
    AM_Inherit = 1
    AM_Locked = 2
End Enum

#If False Then
    Private Const AM_Normal = 0, AM_Inherit = 1, AM_Locked = 2
#End If

'Convert a system color (such as "button face" or "inactive window") to a literal RGB value
Private Declare Function OleTranslateColor Lib "oleaut32" (ByVal oColor As OLE_COLOR, ByVal hPalette As Long, ByRef cColorRef As Long) As Long

'Constants to improve color space conversion performance
Private Const ONE_DIV_SIX As Double = 0.166666666666667
Private Const HEX_PREFIX As String = "&H"

'Collection of SVG color names, extracted at run-time from a compressed text file in PD's resource segment.
' This collection is *not* populated by default; it is automagically populated at first request.
Private m_SVGColors As pdDictionary

'Present the user with PD's custom color selection dialog.
' INPUTS:  1) a Long-type variable (ByRef, of course) which will receive the new color
'          2) an optional initial color
'          3) an optional pdColorSelector control reference, if this dialog is being raised by a pdColorSelector control.
'             (This reference will be used to provide live updates as the user plays with the color dialog.)
'
' OUTPUTS: 1) TRUE if OK was pressed, FALSE for Cancel
Public Function ShowColorDialog(ByRef colorReceive As Long, Optional ByVal initialColor As Long = vbWhite, Optional ByRef callingControl As pdColorSelector, Optional ByRef callerParent As Form = Nothing) As Boolean
    
    'As of November 2014, PhotoDemon has its own color selector!
    If Dialogs.ChoosePDColor(initialColor, colorReceive, callingControl, callerParent) = vbOK Then
        ShowColorDialog = True
    Else
        ShowColorDialog = False
    End If
    
End Function

'Given an OLE color, return an RGB
Public Function ConvertSystemColor(ByVal colorRef As OLE_COLOR) As Long
    
    'OleTranslateColor returns -1 if it fails; if that happens, default to white
    If OleTranslateColor(colorRef, 0&, ConvertSystemColor) = -1 Then
        ConvertSystemColor = RGB(255, 255, 255)
    End If
    
End Function

'Extract the red, green, or blue value from an RGB() Long
Public Function ExtractRed(ByVal currentColor As Long) As Long
    ExtractRed = currentColor And &HFF&
End Function

Public Function ExtractGreen(ByVal currentColor As Long) As Long
    ExtractGreen = (currentColor \ 256) And &HFF&
End Function

Public Function ExtractBlue(ByVal currentColor As Long) As Long
    ExtractBlue = (currentColor \ 65536) And &HFF&
End Function

'Blend byte1 w/ byte2 based on mixRatio. mixRatio is expected to be a value between 0 and 1.
Public Function BlendColors(ByVal firstColor As Long, ByVal secondColor As Long, ByVal mixRatio As Single) As Byte
    BlendColors = Int((1! - mixRatio) * firstColor + 0.5!) + Int(mixRatio * secondColor)
End Function

'This function will return the luminance value of an RGB triplet.  Note that the value will be in the [0,255] range instead
' of the usual [0,1.0] one.
Public Function GetLuminance(ByVal r As Long, ByVal g As Long, ByVal b As Long) As Long
    Dim Max As Long, Min As Long
    Max = Max3Int(r, g, b)
    Min = Min3Int(r, g, b)
    GetLuminance = (Max + Min) * 0.5
End Function

'This function will return a well-calculated luminance value of an RGB triplet.  Note that the value will be in
' the [0,255] range instead of the usual [0,1.0] one.
Public Function GetHQLuminance(ByVal r As Long, ByVal g As Long, ByVal b As Long) As Long
    GetHQLuminance = (218 * r + 732 * g + 74 * b) \ 1024
End Function

Public Function GetRGBAFromRGBAndA(ByVal rgbLong As Long, ByVal a As Long) As Long
    Dim tmpQuad As RGBQuad
    tmpQuad.Red = Colors.ExtractRed(rgbLong)
    tmpQuad.Green = Colors.ExtractGreen(rgbLong)
    tmpQuad.Blue = Colors.ExtractBlue(rgbLong)
    tmpQuad.Alpha = a
    GetMem4 VarPtr(tmpQuad), GetRGBAFromRGBAndA
End Function

'HSL <-> RGB conversion routines.  H is returned on the (weird) range  [-1, 5], but this allows for
' some optimizations that are otherwise difficult.  Please do not use this function if quality is
' paramount; for that, you'd be better off using PreciseRGBtoHSL, below.
Public Sub ImpreciseRGBtoHSL(r As Long, g As Long, b As Long, h As Double, s As Double, l As Double)
    
    Dim inMax As Double, inMin As Double, inDelta As Double
    Dim rR As Double, rG As Double, rB As Double
    
    Const ONE_DIV_255 As Double = 1# / 255#
    rR = r * ONE_DIV_255
    rG = g * ONE_DIV_255
    rB = b * ONE_DIV_255

    'Note: HSL are calculated in the following ranges:
    ' Hue: [-1, 5]
    ' Saturation: [0, 1] (If saturation = 0.0, hue is technically undefined)
    ' Lightness: [0, 1]
    
    'In-line max/min calculations for performance reasons
    'inMax = Max3Float(rR, rG, rB)
    'inMin = Min3Float(rR, rG, rB)
    If (rR > rG) Then
       If (rR > rB) Then inMax = rR Else inMax = rB
    Else
       If (rB > rG) Then inMax = rB Else inMax = rG
    End If
    If (rR < rG) Then
       If (rR < rB) Then inMin = rR Else inMin = rB
    Else
       If (rB < rG) Then inMin = rB Else inMin = rG
    End If
    
    'Calculate luminance
    l = (inMax + inMin) * 0.5
        
    'If the maximum and minimum are identical, this image is gray, meaning it has no saturation
    ' (and thus an undefined hue).
    If (inMax = inMin) Then
        s = 0#
        h = 0#
    Else
        
        inDelta = inMax - inMin
        
        'Calculate saturation
        If (l <= 0.5) Then
            s = inDelta / (inMax + inMin)
        Else
            s = inDelta / (2# - inMax - inMin)
        End If
        
        'Calculate hue.  This code uses a three-quadrant system which is not especially precise.
        ' (A Cos() system that maps to a true circle would be better.)  However, there are some
        ' tasks where we only need quick HSL estimations, and this method yields a large
        ' performance boost over a "perfect" solution.
        If (rR = inMax) Then
            h = (rG - rB) / inDelta
        ElseIf (rG = inMax) Then
            h = 2# + (rB - rR) / inDelta
        Else
            h = 4# + (rR - rG) / inDelta
        End If
        
        'If you need hue in the [0,360] range instead of [-1, 5] you can add this code to your function:
        'h = h * 60.0
        'If (h < 0.0) Then h = h + 360.0

    End If

    'Similarly, if byte values are preferred to floating-point, this code will modify hue to return
    ' on the range [0, 240], saturation on [0, 255], and luminance on [0, 255]
    'H = Int(H * 40.0 + 40.0)
    'S = Int(S * 255.0)
    'L = Int(L * 255.0)
    
End Sub

'Convert HSL values to RGB values.  *Input ranges are non-standard* - see the comments in
' ImpreciseRGBtoHSL(), above, for details.  This function is preferable when quality is not of
' paramount importance.  Estimations are used to calculate hue, which means color returns will
' be imprecise compared to other methods.
Public Sub ImpreciseHSLtoRGB(h As Double, s As Double, l As Double, r As Long, g As Long, b As Long)

    Dim rR As Double, rG As Double, rB As Double
    Dim inMin As Double, inMax As Double
    
    'Failsafe hue check
    If (h > 5#) Then h = h - 6#
    
    'Unsaturated pixels do not technically have hue - they only have luminance
    If (s = 0#) Then
        rR = l: rG = l: rB = l
    Else
        If (l <= 0.5) Then
            inMin = l * (1# - s)
        Else
            inMin = l - s * (1# - l)
        End If
      
        inMax = 2# * l - inMin
      
        If (h < 1#) Then
            
            rR = inMax
            
            If (h < 0#) Then
                rG = inMin
                rB = rG - h * (inMax - inMin)
            Else
                rB = inMin
                rG = rB + h * (inMax - inMin)
            End If
        
        ElseIf (h < 3#) Then
            
            rG = inMax
         
            If (h < 2#) Then
                rB = inMin
                rR = rB - (h - 2#) * (inMax - inMin)
            Else
                rR = inMin
                rB = rR + (h - 2#) * (inMax - inMin)
            End If
        
        Else
        
            rB = inMax
            
            If (h < 4#) Then
                rR = inMin
                rG = rR - (h - 4#) * (inMax - inMin)
            Else
                rG = inMin
                rR = rG + (h - 4#) * (inMax - inMin)
            End If
         
        End If
            
    End If
    
    r = rR * 255#
    g = rG * 255#
    b = rB * 255#
    
    'Failsafe added 29 August '12
    'This should never return RGB values > 255, but it doesn't hurt to make sure.
    If (r > 255) Then r = 255
    If (g > 255) Then g = 255
    If (b > 255) Then b = 255
   
End Sub

'Floating-point conversion between RGB [0, 1] and HSL [0, 1]
Public Sub PreciseRGBtoHSL(ByVal r As Double, ByVal g As Double, ByVal b As Double, ByRef h As Double, ByRef s As Double, ByRef l As Double)

    Dim minVal As Double, maxVal As Double, delta As Double
    minVal = Min3Float(r, g, b)
    maxVal = Max3Float(r, g, b)
    delta = maxVal - minVal

    l = (maxVal + minVal) * 0.5

    'Check the achromatic case
    If (delta = 0#) Then

        'Hue is technically undefined, but we have to return SOME value...
        h = 0#
        s = 0#
        
    'Chromatic case...
    Else
        
        If (l < 0.5) Then
            s = delta / (maxVal + minVal)
        Else
            s = delta / (2# - maxVal - minVal)
        End If
        
        Dim deltaR As Double, deltaG As Double, deltaB As Double, halfDelta As Double, invDelta As Double
        halfDelta = delta * 0.5
        invDelta = 1# / delta

        deltaR = ((maxVal - r) * ONE_DIV_SIX + halfDelta) * invDelta
        deltaG = ((maxVal - g) * ONE_DIV_SIX + halfDelta) * invDelta
        deltaB = ((maxVal - b) * ONE_DIV_SIX + halfDelta) * invDelta
        
        If (r >= maxVal) Then
            h = deltaB - deltaG
        ElseIf (g >= maxVal) Then
            h = 0.333333333333333 + deltaR - deltaB
        Else
            h = 0.666666666666667 + deltaG - deltaR
        End If
        
        'Lock hue to the [0, 1] range
        If (h < 0#) Then h = h + 1#
        If (h > 1#) Then h = h - 1#
    
    End If
    
End Sub

'Floating-point conversion between HSL [0, 1] and RGB [0, 1]
Public Sub PreciseHSLtoRGB(ByVal h As Double, ByVal s As Double, ByVal l As Double, ByRef r As Double, ByRef g As Double, ByRef b As Double)

    'Check the achromatic case
    If (s = 0#) Then
    
        r = l
        g = l
        b = l
    
    'Check the chromatic case
    Else
        
        'As a failsafe, lock hue to [0, 1].
        ' NOTE!  In PhotoDemon, these values are always round-tripped from the matching PreciseRGBtoHSL function,
        ' which performs the failsafe check there.  As such, we don't need it here.
        'If (h < 0#) Then h = h + 1#
        'If (h > 1#) Then h = h - 1#
        
        Dim var_1 As Double, var_2 As Double
        
        If (l < 0.5) Then
            var_2 = l * (1# + s)
        Else
            var_2 = (l + s) - (s * l)
        End If

        var_1 = 2# * l - var_2

        r = fHueToRGB(var_1, var_2, h + 0.333333333333333)
        g = fHueToRGB(var_1, var_2, h)
        b = fHueToRGB(var_1, var_2, h - 0.333333333333333)
        
        'Failsafe check for underflow
        If (r < 0#) Then r = 0#
        If (g < 0#) Then g = 0#
        If (b < 0#) Then b = 0#
        
    End If

End Sub

Private Function fHueToRGB(ByRef v1 As Double, ByRef v2 As Double, ByRef vH As Double) As Double
    
    If (vH < 0#) Then
        vH = vH + 1#
    ElseIf (vH > 1#) Then
        vH = vH - 1#
    End If
    
    If ((6# * vH) < 1#) Then
        fHueToRGB = v1 + (v2 - v1) * 6# * vH
    ElseIf ((2# * vH) < 1#) Then
        fHueToRGB = v2
    ElseIf ((3# * vH) < 2#) Then
        fHueToRGB = v1 + (v2 - v1) * (0.666666666666667 - vH) * 6#
    Else
        fHueToRGB = v1
    End If

End Function

'Convert [0,255] RGB values to [0,1] HSV values, with thanks to easyrgb.com for the conversion math
Public Sub RGBtoHSV(ByVal r As Long, ByVal g As Long, ByVal b As Long, ByRef h As Double, ByRef s As Double, ByRef v As Double)

    Dim fR As Double, fG As Double, fB As Double
    Const ONE_DIV_255 As Double = 1# / 255#
    fR = r * ONE_DIV_255
    fG = g * ONE_DIV_255
    fB = b * ONE_DIV_255

    Dim var_Min As Double, var_Max As Double, del_Max As Double
    var_Min = PDMath.Min3Float(fR, fG, fB)
    var_Max = PDMath.Max3Float(fR, fG, fB)
    del_Max = var_Max - var_Min
    
    'Value is easy to calculate - it's the largest of R/G/B
    v = var_Max

    'If the max and min are the same, this is a gray pixel
    If (del_Max = 0#) Then
        h = 0#
        s = 0#
        
    'If max and min vary, we can calculate a hue component
    Else
        
        s = del_Max / var_Max
        
        Const ONE_DIV_SIX As Double = 1# / 6#
        
        Dim inv_del_Max As Double, half_del_Max As Double
        inv_del_Max = 1# / del_Max
        half_del_Max = del_Max / 2
        
        Dim del_R As Double, del_G As Double, del_B As Double
        del_R = (((var_Max - fR) * ONE_DIV_SIX) + half_del_Max) * inv_del_Max
        del_G = (((var_Max - fG) * ONE_DIV_SIX) + half_del_Max) * inv_del_Max
        del_B = (((var_Max - fB) * ONE_DIV_SIX) + half_del_Max) * inv_del_Max

        If (fR = var_Max) Then
            h = del_B - del_G
        ElseIf (fG = var_Max) Then
            h = 0.333333333333333 + del_R - del_B
        Else
            h = 0.666666666666667 + del_G - del_R
        End If

        If (h < 0#) Then h = h + 1#
        If (h > 1#) Then h = h - 1#

    End If

End Sub

'Convert [0,1] HSV values to [0,255] RGB values, with thanks to easyrgb.com for the conversion math
Public Sub HSVtoRGB(ByRef h As Double, ByRef s As Double, ByRef v As Double, ByRef r As Long, ByRef g As Long, ByRef b As Long)

    'If saturation is 0, RGB are calculated identically
    If (s <= 0#) Then
        r = v * 255#
        g = v * 255#
        b = v * 255#
    
    'If saturation is not 0, we have to calculate RGB independently
    Else
       
        'To keep our math simple, limit hue to [0, 5.9999999]
        Dim var_H As Double
        var_H = h * 6#
        If (var_H >= 6#) Then var_H = 0#
        
        Dim var_I As Long
        var_I = Int(var_H)
        
        Dim var_1 As Double, var_2 As Double, var_3 As Double
        var_1 = v * (1# - s)
        var_2 = v * (1# - s * (var_H - var_I))
        var_3 = v * (1# - s * (1# - (var_H - var_I)))

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
            Case 5
                var_R = v
                var_G = var_1
                var_B = var_2
        End Select

        r = var_R * 255#
        g = var_G * 255#
        b = var_B * 255#
                
    End If

End Sub

'A heavily modified RGB to HSV transform, courtesy of http://lolengine.net/blog/2013/01/13/fast-rgb-to-hsv.
' Note that the code assumes RGB values already in the [0, 1] range, and it will return HSV values in the [0, 1] range.
Public Sub fRGBtoHSV(ByVal r As Double, ByVal g As Double, ByVal b As Double, ByRef h As Double, ByRef s As Double, ByRef v As Double)

    Dim k As Double, tmpSwap As Double, chroma As Double
    
    If (g < b) Then
        tmpSwap = b
        b = g
        g = tmpSwap
        k = -1#
    End If
    
    If (r < g) Then
        tmpSwap = g
        g = r
        r = tmpSwap
        k = -0.333333333333333 - k
    End If
    
    chroma = r - fMin(g, b)
    h = Abs(k + (g - b) / (6# * chroma + 0.0000001))
    s = chroma / (r + 0.00000001)
    v = r
    
End Sub

'Convert [0,1] HSV values to [0,1] RGB values, with thanks to easyrgb.com for the conversion math
Public Sub fHSVtoRGB(ByRef h As Double, ByRef s As Double, ByRef v As Double, ByRef r As Double, ByRef g As Double, ByRef b As Double)

    'If saturation is 0, RGB are calculated identically
    If (s = 0#) Then
        r = v
        g = v
        b = v
        
    'If saturation is not 0, we have to calculate RGB independently
    Else
       
        Dim var_H As Double
        var_H = h * 6#
        
        'To keep our math simple, limit hue to [0, 5.9999999]
        If (var_H >= 6#) Then var_H = 0#
        
        Dim var_I As Long
        var_I = Int(var_H)
        
        Dim var_1 As Double, var_2 As Double, var_3 As Double
        var_1 = v * (1# - s)
        var_2 = v * (1# - s * (var_H - var_I))
        var_3 = v * (1# - s * (1# - (var_H - var_I)))
        
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
    rFloat = r / 255#
    gFloat = g / 255#
    bFloat = b / 255#
    
    'Convert RGB values to the sRGB color space
    If (rFloat > 0.04045) Then
        rFloat = ((rFloat + 0.055) / (1.055)) ^ 2.2
    Else
        rFloat = rFloat / 12.92
    End If
    
    If (gFloat > 0.04045) Then
        gFloat = ((gFloat + 0.055) / (1.055)) ^ 2.2
    Else
        gFloat = gFloat / 12.92
    End If
    
    If (bFloat > 0.04045) Then
        bFloat = ((bFloat + 0.055) / (1.055)) ^ 2.2
    Else
        bFloat = bFloat / 12.92
    End If
    
    'Calculate XYZ using hard-coded values corresponding to sRGB endpoints
    x = rFloat * 0.4124 + gFloat * 0.3576 + bFloat * 0.1805
    y = rFloat * 0.2126 + gFloat * 0.7152 + bFloat * 0.0722
    z = rFloat * 0.0193 + gFloat * 0.1192 + bFloat * 0.9505
    
End Sub

'Convert an XYZ color to CIELab.  As with the original XYZ calculation, D65 is assumed.
' Formula adopted from http://www.easyrgb.com/index.php?X=MATH&H=07#text7, with minor changes by me (not re-applying D65 values until after
'  fXYZ has been calculated)
Public Sub XYZtoLab(ByVal x As Double, ByVal y As Double, ByVal z As Double, ByRef l As Double, ByRef a As Double, ByRef b As Double)
    l = 116# * fXYZ(y) - 16#
    a = 500# * (fXYZ(x / 0.9505) - fXYZ(y))
    b = 200# * (fXYZ(y) - fXYZ(z / 1.089))
End Sub

'Matrix from http://brucelindbloom.com/index.html?Eqn_RGB_XYZ_Matrix.html
Public Sub XYZtosRGB_Float(ByVal x As Single, ByVal y As Single, ByVal z As Single, ByRef r As Single, ByRef g As Single, ByRef b As Single)
    r = (3.2404542 * x + -1.5371385 * y + -0.4985314 * z)
    g = (-0.969266 * x + 1.8760108 * y + 0.041556 * z)
    b = (0.0556434 * x + -0.2040259 * y + 1.0572252 * z)
End Sub

Private Function fXYZ(ByVal t As Double) As Double
    If (t > 0.008856) Then
        fXYZ = t ^ 0.333333333333333
    Else
        fXYZ = (7.787 * t) + 0.137931034482759  '(16.0 / 116.0)
    End If
End Function

'Return the minimum of two floating-point values
Private Function fMin(x As Double, y As Double) As Double
    If (x > y) Then fMin = y Else fMin = x
End Function

'Given a hex color representation, return a matching RGB Long.
' NOTE: this function does not handle alpha, so incoming hex values must be 1, 3, or 6 chars long
' NOTE: that this function DOES NOT validate the incoming string; as a purely internal function, it's assumed you
'       won't send it gibberish!
Public Function GetRGBLongFromHex(ByVal srcHex As String) As Long
    
    'This function will correctly handle *any* valid color hex string, but we deliberately optimize the
    ' most common scenario ("#RRGGBB) for improved performance.
    If (Left$(srcHex, 1) = "#") Then
        If (Len(srcHex) = 7) Then
            GetRGBLongFromHex = RGB(Val(HEX_PREFIX & Mid$(srcHex, 2, 2)), Val(HEX_PREFIX & Mid$(srcHex, 4, 2)), Val(HEX_PREFIX & Right$(srcHex, 2)))
            Exit Function
            
        'Parsing resumes after removing the left "#"
        Else
            srcHex = (Right$(srcHex, Len(srcHex) - 1))
        End If
    End If
    
    'If short-hand length is in use, expand it to 6 chars now
    If (Len(srcHex) < 6) Then
        
        If (Len(srcHex) = 3) Then
            'Three characters is standard shorthand hex; expand each character as a pair
            srcHex = Left$(srcHex, 1) & Left$(srcHex, 1) & Mid$(srcHex, 2, 1) & Mid$(srcHex, 2, 1) & Right$(srcHex, 1) & Right$(srcHex, 1)
        ElseIf (Len(srcHex) = 1) Then
            'One character is treated as a shade of gray; extend it to six characters.
            srcHex = String$(6, srcHex)
        Else
            'We can't handle this character string!
            'Debug.Print "WARNING! Invalid hex passed to GetRGBLongFromHex: " & srcHex
            Exit Function
        End If
        
    End If
    
    'Parse the string to calculate actual numeric values; we can use VB's Val() function for this.
    Dim r As Long, g As Long, b As Long
    r = Val(HEX_PREFIX & Left$(srcHex, 2))
    g = Val(HEX_PREFIX & Mid$(srcHex, 3, 2))
    b = Val(HEX_PREFIX & Right$(srcHex, 2))
    
    'Return the RGB Long
    GetRGBLongFromHex = RGB(r, g, b)
    
End Function

'Given an 8-char hex string (e.g. "ff00ff00") representing RGBA data, return a long-type RGBA quad
Public Function GetRGBALongFromHex(ByVal srcHex As String) As Long
    
    'To make things simpler, remove variability from the source string
    If (InStr(1, srcHex, "#", vbBinaryCompare) <> 0) Then srcHex = (Right$(srcHex, Len(srcHex) - 1))
    
    'Parse the string to calculate actual numeric values; we can use VB's Val() function for this.
    Dim tmpQuad As RGBQuad
    Const HEX_PREFIX As String = "&H"
    tmpQuad.Red = Val(HEX_PREFIX & Left$(srcHex, 2))
    tmpQuad.Green = Val(HEX_PREFIX & Mid$(srcHex, 3, 2))
    tmpQuad.Blue = Val(HEX_PREFIX & Mid$(srcHex, 5, 2))
    tmpQuad.Alpha = Val(HEX_PREFIX & Right$(srcHex, 2))
    
    'Return the RGB Long
    GetMem4_Ptr VarPtr(tmpQuad), VarPtr(GetRGBALongFromHex)
    
End Function

'Given a hex color representation, return a matching RGBA quad (4 bytes in BGRA order).
' NOTE: this function does not handle alpha, so incoming hex values must be 1, 3, or 6 chars long
' NOTE: that this function DOES NOT validate the incoming string; as a purely internal function, it's assumed you
'       won't send it gibberish!
Public Function GetRGBQuadFromHex(ByVal srcHex As String) As RGBQuad
    
    Dim tmpLong As Long
    tmpLong = Colors.GetRGBLongFromHex(srcHex)
    
    GetRGBQuadFromHex.Red = Colors.ExtractRed(tmpLong)
    GetRGBQuadFromHex.Green = Colors.ExtractGreen(tmpLong)
    GetRGBQuadFromHex.Blue = Colors.ExtractBlue(tmpLong)
    GetRGBQuadFromHex.Alpha = 255
    
End Function

'Given an RGB triplet (Long-type), return a matching hex representation.
Public Function GetHexStringFromRGB(ByVal srcRGB As Long) As String
    srcRGB = Colors.ConvertSystemColor(srcRGB)
    GetHexStringFromRGB = GetTwoCharHexFromByte(Colors.ExtractRed(srcRGB)) & GetTwoCharHexFromByte(Colors.ExtractGreen(srcRGB)) & GetTwoCharHexFromByte(Colors.ExtractBlue(srcRGB))
End Function

'HTML hex requires each RGB entry to be two characters wide, but the VB Hex$ function doesn't add a leading 0.
' We can handle this case manually.
Public Function GetTwoCharHexFromByte(ByVal srcByte As Byte) As String
    If (srcByte < 16) Then
        GetTwoCharHexFromByte = "0" & LCase$(Hex$(srcByte))
    Else
        GetTwoCharHexFromByte = LCase$(Hex$(srcByte))
    End If
End Function

'Given some string value, attempt to wring color information out of it.  The goal is to eventually support all valid CSS
' color descriptors (e.g. http://www.w3schools.com/cssref/css_colors_legal.asp), but for now PD primarily uses hex representations.
Public Function IsStringAColor(ByRef srcString As String, Optional ByRef dstColorType As PD_ColorStringType = ColorUnknown, Optional ByVal validateActualColorValue As Boolean = True) As Boolean
    
    dstColorType = ColorUnknown
    
    'Hex validation is fairly easy: is the string prepended with a hash?
    If (Left$(srcString, 1) = "#") Then
        
        dstColorType = ColorHex
        
        'Only perform additional validation as requested
        If validateActualColorValue Then
            
            'Trim out the non-hash characters
            Dim testString As String
            testString = Right$(srcString, Len(srcString) - 1)
            
            'Is the string 1/3/6 chars long?
            Dim lenStr As Long
            lenStr = Len(testString)
            If (lenStr = 1) Or (lenStr = 3) Or (lenStr = 6) Then
    
                'Does the string only consist of the chars 0-9 and A-F?
                If TextSupport.ValidateHexChars(testString) Then
                    dstColorType = ColorHex
                Else
                    dstColorType = ColorUnknown
                End If
    
            Else
                dstColorType = ColorUnknown
            End If
            
        End If
            
    'Next, look for an RGB prefix
    ElseIf Strings.StringsEqual(Left$(srcString, 4), "rgb(", True) Then
        dstColorType = ColorRGB
        
    'Next, look for HSV prefixes
    ElseIf Strings.StringsEqual(Left$(srcString, 5), "rgba(", True) Then
        'TODO
        
    'Lastly, look for SVG color keywords
    ElseIf IsStringAColorName(srcString) Then
        dstColorType = ColorNamed
    End If
    
    'If we've attempted to match all existing color types without success, return failure
    IsStringAColor = (dstColorType <> ColorUnknown)
    If (Not IsStringAColor) Then dstColorType = ColorInvalid
    
End Function

'Given a string representation of a color and the type of representation (optionally; this function will look it up if
' it's missing), return an RGB value and a matching opacity.
' NOTE: at present, opacity is not actually retrieved; it always returns 100.0.  Also, per comments elsewhere in this module,
' not all color representations have been implemented.  Stick to hex for now.
' RETURNS: TRUE if successful; FALSE otherwise
Public Function GetColorFromString(ByRef srcString As String, ByRef dstRGBLong As Long, Optional ByVal srcColorType As PD_ColorStringType = ColorUnknown) As Boolean

    'If the color type is unknown, attempt to identify it now.
    If (srcColorType = ColorInvalid) Or (srcColorType = ColorUnknown) Then GetColorFromString = IsStringAColor(srcString, srcColorType)
    
    'If the color type is STILL unknown and/or invalid, there's nothing we can do.  Exit immediately.
    If (srcColorType = ColorInvalid) Or (srcColorType = ColorUnknown) Then
        If PDMain.IsProgramRunning() Then PDDebug.LogAction "WARNING!  Colors.GetColorFromString was unable to resolve the color string " & srcString & "."
        GetColorFromString = False
        Exit Function
    End If
    
    'If we made it here safely, the chances of returning a valid color are very good.  Assume a success state.
    GetColorFromString = True
    
    'Depending on the color type, return a matching RGB long now (with optional opacity, depending on the color description)
    Select Case srcColorType
    
        Case ColorHex
            dstRGBLong = GetRGBLongFromHex(srcString)
            
        Case ColorRGB
        
        Case ColorRGBA
        
        Case ColorHSL
        
        Case ColorHSLA
        
        Case ColorNamed
            dstRGBLong = GetRGBLongFromNamedColor(srcString)
    
    End Select
    
End Function

'Attempt to match a string against a known list of SVG color names
Private Function IsStringAColorName(ByRef srcString As String) As Boolean
    'Debug.Print "srcstring", srcString
    If (m_SVGColors Is Nothing) Then BuildColorNameList
    IsStringAColorName = m_SVGColors.DoesKeyExist(srcString)
End Function

'Make sure you call IsStringAColorName, above, BEFORE calling this function, to verify that your
' name actually exists in the collection.
Private Function GetRGBLongFromNamedColor(ByRef srcColorName As String) As Long
    If (m_SVGColors Is Nothing) Then BuildColorNameList
    GetRGBLongFromNamedColor = m_SVGColors.GetEntry_Long(srcColorName)
End Function

Private Sub BuildColorNameList()
    
    Set m_SVGColors = New pdDictionary
    
    'Retrieve the list of color names from PD's resource segment
    Dim colorList As String
    If g_Resources.LoadTextResource("named_colors", colorList) Then
    
        'Split the list by line-ending
        Dim sList As pdStringStack
        Set sList = New pdStringStack
        sList.CreateFromMultilineString colorList
        
        'Each line uses a set format, with the name on the left and the color (hex) on the right, e.g.:
        ' mintcream:#f5fffa
        
        ' Parse each line in turn, splitting out the name and color segments as we go.
        Dim i As Long, tmpString As String
        For i = 0 To sList.GetNumOfStrings - 1
            tmpString = sList.GetString(i)
            m_SVGColors.AddEntry Left$(tmpString, InStr(1, tmpString, ":") - 1), Colors.GetRGBLongFromHex(Right$(tmpString, Len(tmpString) - InStr(1, tmpString, ":")))
        Next i
    
    End If
    
End Sub

'Given a PD alpha mode enum, return a corresponding string representation.  PD alpha mode
' strings are ALWAYS 4-chars long; append spaces as necessary.
Public Function GetAlphaModeIDFromString(ByRef srcString As String) As PD_AlphaMode
    Select Case srcString
        Case "norm"
            GetAlphaModeIDFromString = AM_Normal
        Case "lock"
            GetAlphaModeIDFromString = AM_Locked
        Case "inhr"
            GetAlphaModeIDFromString = AM_Inherit
        Case Else
            GetAlphaModeIDFromString = AM_Normal
            PDDebug.LogAction "WARNING! Colors.GetAlphaModeIDFromString received a bad value: " & srcString
    End Select
End Function

Public Function GetAlphaModeStringFromID(ByVal srcID As PD_AlphaMode) As String
    Select Case srcID
        Case AM_Normal
            GetAlphaModeStringFromID = "norm"
        Case AM_Locked
            GetAlphaModeStringFromID = "lock"
        Case AM_Inherit
            GetAlphaModeStringFromID = "inhr"
        Case Else
            GetAlphaModeStringFromID = "norm"
            PDDebug.LogAction "WARNING! Colors.GetAlphaModeStringFromID received a bad value: " & srcID
    End Select
End Function

'Given a PD blend mode enum, return a corresponding string representation.  PD blend mode
' strings are ALWAYS 4-chars long; append spaces as necessary.
Public Function GetBlendModeIDFromString(ByRef srcString As String) As PD_BlendMode

    Select Case srcString
        Case "norm"
            GetBlendModeIDFromString = BM_Normal
        Case "dark"
            GetBlendModeIDFromString = BM_Darken
        Case "mult"
            GetBlendModeIDFromString = BM_Multiply
        Case "cbrn"
            GetBlendModeIDFromString = BM_ColorBurn
        Case "lbrn"
            GetBlendModeIDFromString = BM_LinearBurn
        Case "lght"
            GetBlendModeIDFromString = BM_Lighten
        Case "scrn"
            GetBlendModeIDFromString = BM_Screen
        Case "cddg"
            GetBlendModeIDFromString = BM_ColorDodge
        Case "lddg"
            GetBlendModeIDFromString = BM_LinearDodge
        Case "ovrl"
            GetBlendModeIDFromString = BM_Overlay
        Case "sftl"
            GetBlendModeIDFromString = BM_SoftLight
        Case "hrdl"
            GetBlendModeIDFromString = BM_HardLight
        Case "vvdl"
            GetBlendModeIDFromString = BM_VividLight
        Case "lnrl"
            GetBlendModeIDFromString = BM_LinearLight
        Case "pinl"
            GetBlendModeIDFromString = BM_PinLight
        Case "hrdm"
            GetBlendModeIDFromString = BM_HardMix
        Case "diff"
            GetBlendModeIDFromString = BM_Difference
        Case "excl"
            GetBlendModeIDFromString = BM_Exclusion
        Case "subt"
            GetBlendModeIDFromString = BM_Subtract
        Case "divd"
            GetBlendModeIDFromString = BM_Divide
        Case "hue "
            GetBlendModeIDFromString = BM_Hue
        Case "satr"
            GetBlendModeIDFromString = BM_Saturation
        Case "clr "
            GetBlendModeIDFromString = BM_Color
        Case "lumn"
            GetBlendModeIDFromString = BM_Luminosity
        Case "gext"
            GetBlendModeIDFromString = BM_GrainExtract
        Case "gmrg"
            GetBlendModeIDFromString = BM_GrainMerge
        Case "eras"
            GetBlendModeIDFromString = BM_Erase
        Case "bhnd"
            GetBlendModeIDFromString = BM_Behind
        Case "copy"
            GetBlendModeIDFromString = BM_Overwrite
        Case Else
            GetBlendModeIDFromString = BM_Normal
            PDDebug.LogAction "WARNING! Colors.GetBlendModeStringFromID received a bad value: " & srcString
    End Select
    
End Function

'Given a PD blend mode string representation, return a corresponding enum.  PD blend mode
' strings are ALWAYS 4-chars long.
Public Function GetBlendModeStringFromID(ByVal srcMode As PD_BlendMode) As String
    
    Select Case srcMode
        Case BM_Normal
            GetBlendModeStringFromID = "norm"
        Case BM_Darken
            GetBlendModeStringFromID = "dark"
        Case BM_Multiply
            GetBlendModeStringFromID = "mult"
        Case BM_ColorBurn
            GetBlendModeStringFromID = "cbrn"
        Case BM_LinearBurn
            GetBlendModeStringFromID = "lbrn"
        Case BM_Lighten
            GetBlendModeStringFromID = "lght"
        Case BM_Screen
            GetBlendModeStringFromID = "scrn"
        Case BM_ColorDodge
            GetBlendModeStringFromID = "cddg"
        Case BM_LinearDodge
            GetBlendModeStringFromID = "lddg"
        Case BM_Overlay
            GetBlendModeStringFromID = "ovrl"
        Case BM_SoftLight
            GetBlendModeStringFromID = "sftl"
        Case BM_HardLight
            GetBlendModeStringFromID = "hrdl"
        Case BM_VividLight
            GetBlendModeStringFromID = "vvdl"
        Case BM_LinearLight
            GetBlendModeStringFromID = "lnrl"
        Case BM_PinLight
            GetBlendModeStringFromID = "pinl"
        Case BM_HardMix
            GetBlendModeStringFromID = "hrdm"
        Case BM_Difference
            GetBlendModeStringFromID = "diff"
        Case BM_Exclusion
            GetBlendModeStringFromID = "excl"
        Case BM_Subtract
            GetBlendModeStringFromID = "subt"
        Case BM_Divide
            GetBlendModeStringFromID = "divd"
        Case BM_Hue
            GetBlendModeStringFromID = "hue "
        Case BM_Saturation
            GetBlendModeStringFromID = "satr"
        Case BM_Color
            GetBlendModeStringFromID = "clr "
        Case BM_Luminosity
            GetBlendModeStringFromID = "lumn"
        Case BM_GrainExtract
            GetBlendModeStringFromID = "gext"
        Case BM_GrainMerge
            GetBlendModeStringFromID = "gmrg"
        Case BM_Erase
            GetBlendModeStringFromID = "eras"
        Case BM_Behind
            GetBlendModeStringFromID = "bhnd"
        Case BM_Overwrite
            GetBlendModeStringFromID = "copy"
        Case Else
            GetBlendModeStringFromID = "    "
            PDDebug.LogAction "WARNING! Colors.GetBlendModeStringFromID received a bad value: " & srcMode
    End Select
    
End Function
