Attribute VB_Name = "Filters_Stylize"
'***************************************************************************
'Stylize Filter Collection
'Copyright 2002-2015 by Tanner Helland
'Created: 8/April/02
'Last updated: 02/April/15
'Last update: finish optimizing new Color Halftone filter
'
'Container module for PD's stylize filter collection.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

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
    Dim dstImageData() As Byte
    Dim dstSA As SAFEARRAY2D
    prepSafeArray dstSA, dstDIB
    CopyMemory ByVal VarPtrArray(dstImageData()), VarPtr(dstSA), 4
    
    'Do the same for the source iamge
    Dim srcImageData() As Byte
    Dim srcSA As SAFEARRAY2D
    prepSafeArray srcSA, srcDIB
    CopyMemory ByVal VarPtrArray(srcImageData()), VarPtr(srcSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = srcDIB.getDIBWidth - 1
    finalY = srcDIB.getDIBHeight - 1
        
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long
    qvDepth = srcDIB.getDIBColorDepth \ 8
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    If Not suppressMessages Then
        If modifyProgBarMax = -1 Then
            SetProgBarMax finalX * 3
        Else
            SetProgBarMax modifyProgBarMax
        End If
        progBarCheck = findBestProgBarValue()
    End If
        
    'Because we want each halftone point centered around a grid intersection, we'll precalculate a half-radius value as well
    Dim pxRadiusHalf As Double
    pxRadiusHalf = pxRadius / 2
    
    'Density is a [0, 1] scale, but we report it to the user as [0, 100]; transform it now
    dotDensity = dotDensity / 100
    
    'At maximum density, a dot of max luminance will extend from the center of a grid point to the diagonal edge
    ' of the grid "block".  This is a distance of Sqr(2) * (grid block size / 2).  We multiply our density value
    ' - in advance - by this value, which simplifies dot calculations in the inner loop.
    dotDensity = dotDensity * Sqr(2) * pxRadiusHalf
        
    'Convert the various input rotation angles to radians
    cyanAngle = cyanAngle * (PI / 180)
    yellowAngle = yellowAngle * (PI / 180)
    magentaAngle = magentaAngle * (PI / 180)
    
    'Prep a bunch of calculation values.  (Yes, there are many.)
    Dim cosTheta As Double, sinTheta As Double
    Dim rotateAngle As Double
    Dim srcX As Double, srcY As Double, srcXInner As Double, srcYInner As Double
    Dim dstX As Double, dstY As Double
    Dim clampX As Long, clampY As Long
    Dim r As Long, g As Long, b As Long, a As Long
    Dim target As Long, newTarget As Long, fTarget As Double
    Dim dX As Double, dY As Double
    Dim tmpRadius As Double, f2 As Double, f3 As Double
    Dim overlapCheck As Long
    
    'Because dots can overlap (see details in the inner loop comments), we will occasionally need to check neighboring grid
    ' blocks to determine proper overlap colors.  To simplify calculations in the performance-sensitive inner loop, we cache
    ' all neighboring grid offsets in advance.  (Note that additional heuristics are used inside the loop, so these tables
    ' are not needed for all pixels.)
    Dim xCheck() As Double, yCheck() As Double
    ReDim xCheck(0 To 3) As Double
    ReDim yCheck(0 To 3) As Double
    
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
    Dim densityLookup() As Single
    ReDim densityLookup(0 To 255) As Single
    
    For x = 0 To 255
        
        'Convert the color value to floating-point CMY, then square it (which yields better luminance control)
        fTarget = x / 255
        fTarget = 1 - (fTarget * fTarget)
            
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
            QuickVal = x * qvDepth
        For y = initY To finalY
            
            'Calculate a source position for this pixel, considering the user-supplied angle
            srcX = x * cosTheta + y * sinTheta
            srcY = -x * sinTheta + y * cosTheta
            
            'Lock those source values to a predetermined grid, using the supplied radius value as grid size
            srcX = srcX - Modulo(srcX - pxRadiusHalf, pxRadius) + pxRadiusHalf
            srcY = srcY - Modulo(srcY - pxRadiusHalf, pxRadius) + pxRadiusHalf
            
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
            If (clampX < 0) Then
                clampX = 0
            ElseIf (clampX > finalX) Then
                clampX = finalX
            End If
            
            clampY = dstY
            If (clampY < 0) Then
                clampY = 0
            ElseIf (clampY > finalY) Then
                clampY = finalY
            End If
            
            'Retrieve the relevant channel color at this position
            target = srcImageData(clampX * qvDepth + curChannel, clampY)
                        
            'Calculate a dot size, relative to the underlying grid control point
            dX = x - dstX
            dY = y - dstY
            tmpRadius = Sqr(dX * dX + dY * dY) + 1
            
            'With a circle radius calculated for this intensity value, apply some basic antialiasing if the pixel
            ' lies along the circle edge.
            f2 = 1 - basicAA(tmpRadius - 1, tmpRadius, densityLookup(target))
            
            'If this dot's calculated radius density is greater than a grid block's half-width, this "dot" extends outside
            ' its underlying grid block.  This means it overlaps a neighboring grid, which may have a *different* maximum
            ' density for this channel.  To ensure proper calculations, we must check the neighboring grid locations,
            ' and find the smallest possible value within the overlapping area.  (This strategy makes the function
            ' properly deterministic, so the darkest dot is always guaranteed to be on "top", regardless of channel
            ' processing order.)
            If tmpRadius >= pxRadiusHalf Then
                
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
                    If (clampX < 0) Then
                        clampX = 0
                    ElseIf (clampX > finalX) Then
                        clampX = finalX
                    End If
                    
                    clampY = dstY
                    If (clampY < 0) Then
                        clampY = 0
                    ElseIf (clampY > finalY) Then
                        clampY = finalY
                    End If
                    
                    'Calculate an intensity and radius for this overlapped point
                    newTarget = srcImageData(clampX * qvDepth + curChannel, clampY)
                    dX = x - dstX
                    dY = y - dstY
                    tmpRadius = Sqr(dX * dX + dY * dY)
                    f3 = 1 - basicAA(tmpRadius, tmpRadius + 1, densityLookup(newTarget))
                    
                    'Store the *minimum* calculated value (e.g. the darkest color in this area of overlap)
                    If f3 < f2 Then
                        f2 = f3
                        target = newTarget
                    End If
                
                'Proceed to the next overlapping pixel
                Next overlapCheck
            
            End If
            
            'Convert the final calculated intensity back to byte range, and set the corresponding color in the
            ' destination array.
            target = 255 * f2
            dstImageData(QuickVal + curChannel, y) = target
            
        Next y
            If Not suppressMessages Then
                If (x And progBarCheck) = 0 Then
                    If userPressedESC() Then Exit For
                    SetProgBarVal x + (finalX * curChannel) + modifyProgBarOffset
                End If
            End If
        Next x
        
    Next curChannel
    
    'With our work complete, point both ImageData() arrays away from their DIBs and deallocate them
    CopyMemory ByVal VarPtrArray(srcImageData), 0&, 4
    Erase srcImageData
    
    CopyMemory ByVal VarPtrArray(dstImageData), 0&, 4
    Erase dstImageData
    
    If cancelCurrentAction Then CreateColorHalftoneDIB = 0 Else CreateColorHalftoneDIB = 1

End Function

'This function - courtesy of Jerry Huxtable and jhlabs.com - provides nice, cheap antialiasing along a 1px border
' between two double-type values.
Private Function basicAA(ByRef a As Double, ByRef b As Double, ByRef x As Single) As Double

    If (x < a) Then
        basicAA = 0
    ElseIf (x >= b) Then
        basicAA = 1
    Else
        basicAA = (x - a) / (b - a)
        
        'In his original code, Jerry used a more complicated AA approach, but it seems overkill for a function like this
        ' (especially where the quality trade-off is so minimal):
        'basicAA = x * x * (3 - 2 * x)
    End If

End Function
