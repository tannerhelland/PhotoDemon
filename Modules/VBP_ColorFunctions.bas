Attribute VB_Name = "Color_Functions"
'***************************************************************************
'Miscellaneous Color Functions
'Copyright ©2012-2013 by Tanner Helland
'Created: 13/June/13
'Last updated: 13/June/13
'Last update: created a dedicated module for color processing functions
'
'Many of these functions are older than the create date above, but I did not organize them into a consistent module
' until June 2013.  This module is now used to store all the random bits of specialized color processing code
' required by the program.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://www.tannerhelland.com/photodemon/#license
'
'***************************************************************************

Option Explicit

'Convert a system color (such as "button face" or "inactive window") to a literal RGB value
Private Declare Function OleTranslateColor Lib "olepro32" (ByVal oColor As OLE_COLOR, ByVal HPALETTE As Long, ByRef cColorRef As Long) As Long

'Present the user with a color selection dialog.  At present, this is just a thin wrapper to the stock Windows color
' selector, but in the future it will link to a custom PhotoDemon one.
' INPUTS:  1) a Long-type variable that will receive the new color
'          2) an optional intial color
'
' OUTPUTS: 1) TRUE if OK was pressed, FALSE for Cancel
Public Function showColorDialog(ByRef colorReceive As Long, ByRef dialogOwner As Form, Optional ByVal initialColor As Long = vbWhite) As Boolean

    'For now, use a standard Windows color selector
    Dim retColor As Long
    Dim CD1 As cCommonDialog
    Set CD1 = New cCommonDialog
    retColor = initialColor
    
    If CD1.VBChooseColor(retColor, True, True, False, dialogOwner.hWnd) Then
        colorReceive = retColor
        showColorDialog = True
    Else
        showColorDialog = False
    End If

End Function

'Given the number of colors in an image (as supplied by getQuickColorCount, below), return the highest color depth
' that includes all those colors and is supported by PhotoDemon (1/4/8/24/32)
Public Function getColorDepthFromColorCount(ByVal srcColors As Long, ByRef refLayer As pdLayer) As Long
    
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
        If refLayer.getLayerColorDepth = 24 Then
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
    
    'Create a local array and point it at the pixel data we want to operate on
    Dim ImageData() As Byte
    Dim tmpSA As SAFEARRAY2D
    prepSafeArray tmpSA, srcImage.mainLayer
    CopyMemory ByVal VarPtrArray(ImageData()), VarPtr(tmpSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, finalX As Long, finalY As Long
    finalX = srcImage.Width - 1
    finalY = srcImage.Height - 1
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long
    qvDepth = (srcImage.mainLayer.getLayerColorDepth) \ 8
    
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
Public Function ExtractR(ByVal CurrentColor As Long) As Integer
    ExtractR = CurrentColor Mod 256
End Function

Public Function ExtractG(ByVal CurrentColor As Long) As Integer
    ExtractG = (CurrentColor \ 256) And 255
End Function

Public Function ExtractB(ByVal CurrentColor As Long) As Integer
    ExtractB = (CurrentColor \ 65536) And 255
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
    Dim delta As Double
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
        
        delta = Max - Min
        
        'Calculate saturation
        If l <= 0.5 Then
            s = delta / (Max + Min)
        Else
            s = delta / (2 - Max - Min)
        End If
        
        'Calculate hue
        
        If rR = Max Then
            h = (rG - rB) / delta    '{Resulting color is between yellow and magenta}
        ElseIf rG = Max Then
            h = 2 + (rB - rR) / delta '{Resulting color is between cyan and yellow}
        ElseIf rB = Max Then
            h = 4 + (rR - rG) / delta '{Resulting color is between magenta and cyan}
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
