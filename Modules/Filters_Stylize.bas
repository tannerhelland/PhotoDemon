Attribute VB_Name = "Filters_Stylize"
'***************************************************************************
'Stylize Filter Collection
'Copyright 2002-2026 by Tanner Helland
'Created: 8/April/02
'Last updated: 19/October/17
'Last update: add upgraded "Antique" filter effect
'
'Container module for PD's stylize filter collection.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

Private m_tmpDIB As pdDIB

'Given two DIBs, fill one with a stylized "color halftone" version of the other.
' Per PhotoDemon convention, this function will return a non-zero value if successful, and 0 if canceled.
'
'As described in FormColorHalftone, this tool's algorithm was developed with help from a similar function
' originally written by Jerry Huxtable of JH Labs. Jerry's original code is licensed under an Apache 2.0 license
' (http://www.apache.org/licenses/LICENSE-2.0).  You may download his original version from the following link
' (good as of March '15): http://www.jhlabs.com/ip/filters/index.html
'
'Please note that there are many differences between his version and mine, and I'd definitely recommend his version
' for beginners, as PD includes many complicated optimizations and other modifications.
Public Function CreateColorHalftoneDIB(ByVal pxRadius As Double, ByVal cyanAngle As Double, ByVal magentaAngle As Double, ByVal yellowAngle As Double, ByVal dotDensity As Double, ByRef srcDIB As pdDIB, ByRef dstDIB As pdDIB, Optional ByVal suppressMessages As Boolean = False, Optional ByVal modifyProgBarMax As Long = -1, Optional ByVal modifyProgBarOffset As Long = 0) As Long
    
    'Create a local array and point it at the pixel data of the destination image
    Dim dstImageData() As Byte, dstSA As SafeArray2D
    dstDIB.WrapArrayAroundDIB dstImageData, dstSA
    
    'Do the same for the source iamge
    Dim srcImageData() As Byte, srcSA As SafeArray2D
    srcDIB.WrapArrayAroundDIB srcImageData, srcSA
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = srcDIB.GetDIBWidth - 1
    finalY = srcDIB.GetDIBHeight - 1
    
    Dim xStride As Long
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    If Not suppressMessages Then
        If (modifyProgBarMax = -1) Then
            SetProgBarMax finalX * 3
        Else
            SetProgBarMax modifyProgBarMax
        End If
        progBarCheck = ProgressBars.FindBestProgBarValue()
    End If
        
    'Because we want each halftone point centered around a grid intersection, we'll precalculate a half-radius value as well
    Dim pxRadiusHalf As Double
    pxRadiusHalf = pxRadius * 0.5
    
    'Density is a [0, 1] scale, but we report it to the user as [0, 100]; transform it now
    dotDensity = dotDensity * 0.01
    
    'At maximum density, a dot of max luminance will extend from the center of a grid point to the diagonal edge
    ' of the grid "block".  This is a distance of Sqr(2) * (grid block size / 2).  We multiply our density value
    ' - in advance - by this value, which simplifies dot calculations in the inner loop.
    dotDensity = dotDensity * Sqr(2#) * pxRadiusHalf
        
    'Convert the various input rotation angles to radians
    cyanAngle = cyanAngle * (PI / 180#)
    yellowAngle = yellowAngle * (PI / 180#)
    magentaAngle = magentaAngle * (PI / 180#)
    
    'Prep a bunch of calculation values.  (Yes, there are many.)
    Dim cosTheta As Double, sinTheta As Double
    Dim rotateAngle As Double
    Dim srcX As Double, srcY As Double, srcXInner As Double, srcYInner As Double
    Dim dstX As Double, dstY As Double
    Dim clampX As Long, clampY As Long
    Dim target As Long, newTarget As Long, fTarget As Double
    Dim dx As Double, dy As Double
    Dim tmpRadius As Double, f2 As Double, F3 As Double
    Dim overlapCheck As Long
    
    'Because dots can overlap (see details in the inner loop comments), we will occasionally need to check neighboring grid
    ' blocks to determine proper overlap colors.  To simplify calculations in the performance-sensitive inner loop, we cache
    ' all neighboring grid offsets in advance.  (Note that additional heuristics are used inside the loop, so these tables
    ' are not needed for all pixels.)
    Dim xCheck(0 To 3) As Double, yCheck(0 To 3) As Double
    xCheck(0) = -1 * pxRadius
    yCheck(0) = 0
    xCheck(1) = 0
    yCheck(1) = -1 * pxRadius
    xCheck(2) = pxRadius
    yCheck(2) = 0
    xCheck(3) = 0
    yCheck(3) = pxRadius
    
    'We can also pre-calculate a lookup table between pixel values and density, since density is hard-coded according to
    ' luminance and dot size.  This spares additional calculations on the inner loop.
    Dim densityLookup(0 To 255) As Single
    
    For x = 0 To 255
        
        'Convert the color value to floating-point CMY, then square it (which yields better luminance control)
        fTarget = x / 255#
        fTarget = 1# - (fTarget * fTarget)
            
        'Modify the radius to match the density requested by the user
        fTarget = fTarget * dotDensity
        
        densityLookup(x) = fTarget
        
    Next x
    
    'We are now ready to loop through each pixel in the image, converting values as we go.
    
    'Unique to this filter is separate handling for each channel.  Because each channel can be independently rotated, this proves
    ' more efficient, as we can reuse all calculations for a given channel, rather than recalculating them 3x on each pixel.
    ' (Also note that this function does not modify alpha.)
    Dim curChannel As Long
    For curChannel = 0 To 2
        
        'Populate new lookup values for this channel
        Select Case curChannel
        
            Case 0
                rotateAngle = yellowAngle
            
            Case 1
                rotateAngle = magentaAngle
            
            Case 2
                rotateAngle = cyanAngle
            
        End Select
        
        cosTheta = Cos(rotateAngle)
        sinTheta = Sin(rotateAngle)
                
        'With all lookup values cached, start mapping
        For x = initX To finalX
            xStride = x * 4
        For y = initY To finalY
            
            'Calculate a source position for this pixel, considering the user-supplied angle
            srcX = x * cosTheta + y * sinTheta
            srcY = -x * sinTheta + y * cosTheta
            
            'Lock those source values to a predetermined grid, using the supplied radius value as grid size
            srcX = srcX - PDMath.Modulo(srcX - pxRadiusHalf, pxRadius) + pxRadiusHalf
            srcY = srcY - PDMath.Modulo(srcY - pxRadiusHalf, pxRadius) + pxRadiusHalf
            
            'We now have a literal "round peg in square hole" problem.  The halftone dots we're drawing are circles, but the underlying
            ' calculation grid is comprised of squares.  This presents an ugly problem: dots can extend outside the underlying grid
            ' block, overlapping the dots of neighboring grids.
            
            'One solution - which I think PhotoShop uses - is to simply restrict each circle size to the size of the underlying grid
            ' block.  This is computationally efficient, but it means that a processed image will always have a ton of dead space,
            ' because even solid black regions can only be partially drawn (as the space between the circle and its containing grid
            ' block must always be white).
            
            'We use a more comprehsnive solution, which is checking neighboring grid locations for overlap.  Because this is
            ' very performance-intensive, heuristics are used to determine if a pixel needs this level of calculation, so we can
            ' skip it if at all possible.
            
            'Anyway, we always start by calculating the default value first.
            
            'Convert the newly grid-aligned points *back* into image space
            dstX = srcX * cosTheta - srcY * sinTheta
            dstY = srcX * sinTheta + srcY * cosTheta
            
            'Clamp the grid-aligned pixel to image boundaries.  (For performance reasons, this is locked to integer values.)
            clampX = dstX
            If (clampX < 0) Then clampX = 0
            If (clampX > finalX) Then clampX = finalX
            
            clampY = dstY
            If (clampY < 0) Then clampY = 0
            If (clampY > finalY) Then clampY = finalY
            
            'Retrieve the relevant channel color at this position
            target = srcImageData(clampX * 4 + curChannel, clampY)
            
            'Calculate a dot size, relative to the underlying grid control point
            dx = x - dstX
            dy = y - dstY
            tmpRadius = Sqr(dx * dx + dy * dy) + 1#
            
            'With a circle radius calculated for this intensity value, apply some basic antialiasing if the pixel
            ' lies along the circle edge.
            f2 = 1# - BasicAA(tmpRadius - 1#, tmpRadius, densityLookup(target))
            
            'If this dot's calculated radius density is greater than a grid block's half-width, this "dot" extends outside
            ' its underlying grid block.  This means it overlaps a neighboring grid, which may have a *different* maximum
            ' density for this channel.  To ensure proper calculations, we must check the neighboring grid locations,
            ' and find the smallest possible value within the overlapping area.  (This strategy makes the function
            ' properly deterministic, so the darkest dot is always guaranteed to be on "top", regardless of channel
            ' processing order.)
            If (tmpRadius >= pxRadiusHalf) Then
                
                'Check four neighboring pixels (left, up, right, down)
                For overlapCheck = 0 To 3
                
                    'We can shortcut some calculations, because they are relative to values we've already calculated above.
                    
                    'Start by calculating modified source (x, y) coordinates for this neighboring grid point.
                    srcXInner = srcX + xCheck(overlapCheck)
                    srcYInner = srcY + yCheck(overlapCheck)
                    
                    'Repeat the transform back into image space, including clamping
                    dstX = srcXInner * cosTheta - srcYInner * sinTheta
                    dstY = srcXInner * sinTheta + srcYInner * cosTheta
                    
                    clampX = dstX
                    If (clampX < 0) Then clampX = 0
                    If (clampX > finalX) Then clampX = finalX
                    
                    clampY = dstY
                    If (clampY < 0) Then clampY = 0
                    If (clampY > finalY) Then clampY = finalY
                    
                    'Calculate an intensity and radius for this overlapped point
                    newTarget = srcImageData(clampX * 4 + curChannel, clampY)
                    dx = x - dstX
                    dy = y - dstY
                    tmpRadius = Sqr(dx * dx + dy * dy)
                    F3 = 1# - BasicAA(tmpRadius, tmpRadius + 1#, densityLookup(newTarget))
                    
                    'Store the *minimum* calculated value (e.g. the darkest color in this area of overlap)
                    If (F3 < f2) Then
                        f2 = F3
                        target = newTarget
                    End If
                
                'Proceed to the next overlapping pixel
                Next overlapCheck
            
            End If
            
            'Convert the final calculated intensity back to byte range, and set the corresponding color in the
            ' destination array.
            dstImageData(xStride + curChannel, y) = Int(255# * f2)
            
        Next y
            If Not suppressMessages Then
                If (x And progBarCheck) = 0 Then
                    If Interface.UserPressedESC() Then Exit For
                    SetProgBarVal x + (finalX * curChannel) + modifyProgBarOffset
                End If
            End If
        Next x
        
    Next curChannel
    
    'Safely deallocate all image arrays
    srcDIB.UnwrapArrayFromDIB srcImageData
    dstDIB.UnwrapArrayFromDIB dstImageData
    
    If g_cancelCurrentAction Then CreateColorHalftoneDIB = 0 Else CreateColorHalftoneDIB = 1

End Function

'This function - courtesy of Jerry Huxtable and jhlabs.com - provides nice, cheap antialiasing along a 1px border
' between two double-type values.
Private Function BasicAA(ByVal a As Double, ByVal b As Double, ByVal x As Single) As Double

    If (x < a) Then
        BasicAA = 0#
    ElseIf (x >= b) Then
        BasicAA = 1#
    Else
        BasicAA = (x - a) / (b - a)
        
        'In his original code, Jerry used a more complicated AA approach, but it seems overkill for a function like this
        ' (especially where the quality trade-off is so minimal):
        'BasicAA = x * x * (3# - 2# * x)
        
    End If

End Function

'Render an "antique" effect to an arbitrary DIB.  The DIB must already exist and be sized to whatever dimensions
' the caller requires.
Public Function ApplyAntiqueEffect(ByRef dstDIB As pdDIB, ByVal colorStrength As Double, ByVal colorSoftness As Double, Optional ByVal colorSoftnessOpacity As Double = 100#, Optional ByVal grainAmt As Double = 0#, Optional ByVal vignetteAmt As Double = 0#, Optional ByVal suppressMessages As Boolean = False, Optional ByVal modifyProgBarMax As Long = -1, Optional ByVal modifyProgBarOffset As Long = 0) As Boolean
    
    'Unpremultiply alpha
    If (dstDIB Is Nothing) Then Exit Function
    If dstDIB.GetAlphaPremultiplication() Then dstDIB.SetAlphaPremultiplication False
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = dstDIB.GetDIBWidth - 1
    finalY = dstDIB.GetDIBHeight - 1
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    If (Not suppressMessages) Then
        If (modifyProgBarMax = -1) Then SetProgBarMax finalY * 2 Else SetProgBarMax modifyProgBarMax
        progBarCheck = ProgressBars.FindBestProgBarValue()
    End If
    
    Dim highlightColor As Long, shadowColor As Long, balance As Double
    highlightColor = vbBlue
    shadowColor = vbRed
    balance = 0.5
    
    'From the incoming colors, determine corresponding hue and saturation values
    Dim highlightHue As Double, highlightSaturation As Double, shadowHue As Double, shadowSaturation As Double
    Dim ignoreLuminance As Double
    PreciseRGBtoHSL Colors.ExtractRed(highlightColor) / 255#, Colors.ExtractGreen(highlightColor) / 255#, Colors.ExtractBlue(highlightColor) / 255#, highlightHue, highlightSaturation, ignoreLuminance
    PreciseRGBtoHSL Colors.ExtractRed(shadowColor) / 255#, Colors.ExtractGreen(shadowColor) / 255#, Colors.ExtractBlue(shadowColor) / 255#, shadowHue, shadowSaturation, ignoreLuminance
    
    'Convert balance mix value from an incoming range of [-100, 100] to a new range of [0,1].  We use this value
    ' to map colors between the shadow tone, neutral gray, and the highlight tone.
    Dim balGradient As Double, invBalGradient As Double
    invBalGradient = (balance + 100#) / 200#
    balGradient = 1# - invBalGradient
    
    'Prevent divide-by-zero errors, below
    If (invBalGradient <= 0.0000001) Then invBalGradient = 0.0000001
    If (balGradient <= 0.0000001) Then balGradient = 0.0000001
    
    'To avoid the need for many divisions on the inner loop, calculate inverse values now
    Dim multBalGradient As Double, multInvBalGradient As Double
    multInvBalGradient = 1# / invBalGradient
    multBalGradient = 1# / balGradient
    
    'Strength controls the ratio at which the split-toned pixels are merged with the original pixels.
    ' Convert it from a [0, 100] to [0, 1] scale.
    colorStrength = colorStrength * 0.01
    
    'Color variables
    Dim r As Long, g As Long, b As Long
    Dim newR As Long, newG As Long, newB As Long
    
    'Create a local array and point it at the pixel data of the current image
    Dim dstImageData() As Byte, dstSA As SafeArray1D
    dstDIB.WrapArrayAroundScanline dstImageData, dstSA, 0
    
    Dim dibPtr As Long, dibStride As Long
    dibPtr = dstSA.pvData
    dibStride = dstSA.cElements
    
    finalX = finalX * 4
    
    'Loop through each pixel in the image, converting values as we go
    For y = initY To finalY
        dstSA.pvData = dibPtr + dibStride * y
    For x = initX To finalX Step 4
    
        b = dstImageData(x)
        g = dstImageData(x + 1)
        r = dstImageData(x + 2)
        
        'Use w3c sepia settings (https://www.w3.org/TR/filter-effects/#sepiaEquivalent)
        newR = (r * 0.393) + (g * 0.769) + (b * 0.189)
        newG = (r * 0.349) + (g * 0.686) + (b * 0.168)
        newB = (r * 0.272) + (g * 0.534) + (b * 0.131)
        If (newR > 255) Then newR = 255
        If (newG > 255) Then newG = 255
        If (newB > 255) Then newB = 255
                
        'Finally, apply the new RGB values to the image by blending them with their original color at the user's requested strength.
        dstImageData(x) = newB * colorStrength + b * (1# - colorStrength)
        dstImageData(x + 1) = newG * colorStrength + g * (1# - colorStrength)
        dstImageData(x + 2) = newR * colorStrength + r * (1# - colorStrength)
        
    Next x
        If (Not suppressMessages) Then
            If (y And progBarCheck) = 0 Then
                If Interface.UserPressedESC() Then Exit For
                SetProgBarVal modifyProgBarOffset + y
            End If
        End If
    Next y
    
    'Re-premultiply alpha
    dstDIB.UnwrapArrayFromDIB dstImageData
    dstDIB.SetAlphaPremultiplication True
    
    Dim cCompositor As pdCompositor
    Set cCompositor = New pdCompositor
        
    'dstDIB now contains a sepia-toned version of the image.  We now want to create a duplicate copy,
    ' which we will blur according to the user's "softness" parameter.
    If (colorSoftness > 0#) Then
        
        If (m_tmpDIB Is Nothing) Then Set m_tmpDIB = New pdDIB
        m_tmpDIB.CreateFromExistingDIB dstDIB
        Filters_Layers.QuickBlurDIB dstDIB, colorSoftness, False
        
        'We now want to merge the resulting, blurred DIB onto our original copy, using a non-standard
        ' blend mode
        cCompositor.QuickMergeTwoDibsOfEqualSize dstDIB, m_tmpDIB, BM_SoftLight, colorSoftnessOpacity
        
    End If
    
    'We now have a properly softened and light-enhanced image.  Time to add film grain, if any.
    grainAmt = grainAmt * 0.25
    If (grainAmt > 0#) Then
        
        Dim cRandom As pdRandomize
        Set cRandom = New pdRandomize
        cRandom.SetSeed_AutomaticAndRandom
        
        If dstDIB.GetAlphaPremultiplication Then dstDIB.SetAlphaPremultiplication False
        
        'Reset our DIB pointer and stride; they may have changed after rotation, above
        dstDIB.WrapArrayAroundScanline dstImageData, dstSA, 0
        dibPtr = dstSA.pvData
        dibStride = dstSA.cElements
        
        Dim noiseVal As Long
        
        'Loop through each pixel in the image, converting values as we go
        For y = initY To finalY
            dstSA.pvData = dibPtr + dibStride * y
        For x = initX To finalX Step 4
        
            b = dstImageData(x)
            g = dstImageData(x + 1)
            r = dstImageData(x + 2)
            
            'Add monochrome noise to each color
            noiseVal = grainAmt * cRandom.GetGaussianFloat_WH()
            
            r = r + noiseVal
            g = g + noiseVal
            b = b + noiseVal
            
            If (r > 255) Then r = 255
            If (r < 0) Then r = 0
            If (g > 255) Then g = 255
            If (g < 0) Then g = 0
            If (b > 255) Then b = 255
            If (b < 0) Then b = 0
                    
            'Finally, apply the new RGB values to the image by blending them with their original color at the user's requested strength.
            dstImageData(x) = b
            dstImageData(x + 1) = g
            dstImageData(x + 2) = r
            
        Next x
            If (Not suppressMessages) Then
                If (y And progBarCheck) = 0 Then
                    If Interface.UserPressedESC() Then Exit For
                    SetProgBarVal modifyProgBarOffset + finalY + y
                End If
            End If
        Next y
        
        dstDIB.UnwrapArrayFromDIB dstImageData
        dstDIB.SetAlphaPremultiplication True
    
    End If
    
    'Finally, we want to apply a vignette (if any).  Note that our vignette strategy is roughly identical to PD's
    ' separate Effects > Stylize > Vignetting tool; in the future, it would be nice to simply call that function directly.
    If (vignetteAmt > 0#) Then
        
        'Vignette input parameters are hard-coded
        Dim maxRadius As Double, vFeathering As Double
        maxRadius = 65#
        vFeathering = 70#
        
        Dim blendVal As Double
        
        finalX = dstDIB.GetDIBWidth - 1
        
        'Calculate the center of the image, in absolute pixels
        Dim midX As Double, midY As Double
        midX = CDbl(finalX - initX) * 0.5
        midX = midX + initX
        midY = CDbl(finalY - initY) * 0.5
        midY = midY + initY
                
        'X and Y values, remapped around a center point of (0, 0)
        Dim nX As Double, nY As Double
        Dim nX2 As Double, nY2 As Double
                
        'Radius is based off the smaller of the two dimensions - width or height.  (This is used in the "circle" mode.)
        Dim tWidth As Long, tHeight As Long
        tWidth = dstDIB.GetDIBWidth
        tHeight = dstDIB.GetDIBHeight
        
        Dim sRadiusW As Double, sRadiusH As Double
        Dim sRadiusW2 As Double, sRadiusH2 As Double
        Dim sRadiusMin As Double, sRadiusMax As Double
        
        sRadiusW = tWidth * (maxRadius / 100#)
        sRadiusH = tHeight * (maxRadius / 100#)
        sRadiusW2 = sRadiusW * sRadiusW
        sRadiusH2 = sRadiusH * sRadiusH
        
        'Adjust the vignetting to be a proportion of the image's maximum radius.  This ensures accurate correlations
        ' between the preview and the final result.
        Dim vFeathering2 As Double
        vFeathering2 = (vFeathering / 100#) * (sRadiusW * sRadiusH)
        
        'Build a lookup table of vignette values.  Because we're just applying the vignette to a standalone layer,
        ' we can treat the vignette as a constant color scaled from transparent to opaque.  This makes it *very*
        ' fast to apply.
        Dim vLookup(0 To 255) As Long
        Dim tmpQuad As RGBQuad
        
        'Extract the RGB values of the vignetting color
        Dim vColor As Long
        vColor = RGB(255, 255, 255)
        
        newR = Colors.ExtractRed(vColor)
        newG = Colors.ExtractGreen(vColor)
        newB = Colors.ExtractBlue(vColor)
        
        For x = 0 To 255
            With tmpQuad
                .Alpha = x
                blendVal = CSng(x / 255)
                .Red = Int(blendVal * CSng(newR))
                .Green = Int(blendVal * CSng(newG))
                .Blue = Int(blendVal * CSng(newB))
            End With
            GetMem4 VarPtr(tmpQuad), vLookup(x)
        Next x
        
        'We're going to use the temporary DIB for this; that lets us process the vignette much faster, and we can blend
        ' it in a single final swoop using a pdCompositor instance.
        If (m_tmpDIB Is Nothing) Then Set m_tmpDIB = New pdDIB
        m_tmpDIB.CreateBlank dstDIB.GetDIBWidth, dstDIB.GetDIBHeight, 32, 0, 0
        
        Dim dstImageDataL() As Long, tmpSA2D As SafeArray2D
        m_tmpDIB.WrapLongArrayAroundDIB dstImageDataL, tmpSA2D
        
        'And that's it!  Loop through each pixel in the image, converting values as we go.
        For y = initY To finalY
        For x = initX To finalX
        
            'Remap the coordinates around a center point of (0, 0)
            nX = x - midX
            nY = y - midY
            nX2 = nX * nX
            nY2 = nY * nY
            
            sRadiusMax = sRadiusH2 - ((sRadiusH2 * nX2) / sRadiusW2)
            
            'Outside
            If (nY2 > sRadiusMax) Then
                dstImageDataL(x, y) = vLookup(255)
            
            'Inside
            Else
                
                sRadiusMin = sRadiusMax - vFeathering2
                
                'Feathered
                If (nY2 >= sRadiusMin) Then
                    blendVal = (nY2 - sRadiusMin) / vFeathering2
                    dstImageDataL(x, y) = vLookup(blendVal * 255)
                End If
                
            End If
            
        Next x
        Next y
        
        m_tmpDIB.UnwrapLongArrayFromDIB dstImageDataL
        m_tmpDIB.SetInitialAlphaPremultiplicationState True
        
        cCompositor.QuickMergeTwoDibsOfEqualSize dstDIB, m_tmpDIB, BM_Normal, vignetteAmt
        
    End If
    
    ApplyAntiqueEffect = True
    
End Function

