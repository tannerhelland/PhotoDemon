Attribute VB_Name = "Palettes"
'***************************************************************************
'PhotoDemon's Master Palette Interface
'Copyright 2017-2017 by Tanner Helland
'Created: 12/January/17
'Last updated: 15/January/17
'Last update: add the WAPI palette matching function, which is quite a bit faster than our naive version.
'
'This module contains a bunch of helper algorithms for generating optimal palettes from arbitrary source images,
' and also applying arbitrary palettes to images.  In the future, I expect it to include a lot more interesting
' palette code, including swatch imports from a variety of external sources.
'
'In the meantime, please note that this module has quite a few dependencies.  In particular, it performs
' no quantization (and relatively little palette-matching) on its own.  This is primarily delegated to helper classes.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Used for more accurate color distance comparisons (using human eye sensitivity as a rough guide, while staying in
' the sRGB space for performance reasons)
Private Const CUSTOM_WEIGHT_RED As Single = 0.299
Private Const CUSTOM_WEIGHT_GREEN As Single = 0.587
Private Const CUSTOM_WEIGHT_BLUE As Single = 0.114

'WAPI provides palette matching functions that run quite a bit faster than an equivalent VB function; we use this
' if "perfect" palette matching is desired (where an exhaustive search is applied against each pixel in the image,
' and each entry in a palette).
Private Type GDI_PALETTEENTRY
    peR     As Byte
    peG     As Byte
    peB     As Byte
    peFlags As Byte
End Type

Private Type GDI_LOGPALETTE256
    palVersion       As Integer
    palNumEntries    As Integer
    palEntry(0 To 255) As GDI_PALETTEENTRY
End Type

Private Declare Function CreatePalette Lib "gdi32" (ByVal lpLogPalette As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetNearestPaletteIndex Lib "gdi32" (ByVal hPalette As Long, ByVal crColor As Long) As Long

'A specially designed QuickSort algorithm is used to sort the original palette.  This allows us to be
' flexible with sort criteria, and to also cache our sort criteria values so we can reuse them during
' the actual palette matching step.
Private Type PaletteSort
    pSortCriteria As Single
    pOrigIndex As Byte
End Type

'Given a source image, an (empty) destination palette array, and a color count, return an optimized palette using
' the source image as the reference.  A modified median-cut system is used, and it achieves a very nice
' combination of performance, low memory usage, and high-quality output.
'
'Because palette generation is a time-consuming task, the source DIB should generally be shrunk to a much smaller
' version of itself.  I built a function specifically for this: DIB_Support.ResizeDIBByPixelCount().  That function
' resizes an image to a target pixel count, and I wouldn't recommend a net size any larger than ~50,000 pixels.
Public Function GetOptimizedPalette(ByRef srcDIB As pdDIB, ByRef dstPalette() As RGBQUAD, Optional ByVal numOfColors As Long = 256) As Boolean
    
    'Do not request less than two colors in the final palette!
    If (numOfColors < 2) Then numOfColors = 2
    
    Dim srcPixels() As Byte, tmpSA As SAFEARRAY2D
    srcDIB.WrapArrayAroundDIB srcPixels, tmpSA
    
    Dim pxSize As Long
    pxSize = srcDIB.GetDIBColorDepth \ 8
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = srcDIB.GetDIBStride - 1
    finalY = srcDIB.GetDIBHeight - 1
    
    'Add all pixels from the source image to a base color stack
    Dim pxStack() As pdMedianCut
    ReDim pxStack(0 To numOfColors - 1) As pdMedianCut
    Set pxStack(0) = New pdMedianCut
    
    For y = 0 To finalY
    For x = 0 To finalX Step pxSize
        pxStack(0).AddColor_RGB srcPixels(x + 2, y), srcPixels(x + 1, y), srcPixels(x, y)
    Next x
    Next y
    
    srcDIB.UnwrapArrayFromDIB srcPixels
    
    'Next, make sure there are more than [numOfColors] colors in the image (otherwise, our work is already done!)
    If (pxStack(0).GetNumOfColors > numOfColors) Then
        
        Dim stackCount As Long
        stackCount = 1
        
        Dim maxVariance As Long, mvIndex As Long
        Dim i As Long
        
        'With the initial stack constructed, we can now start partitioning it into smaller stacks based on variance
        Do
        
            maxVariance = 0
            mvIndex = 0
            
            Dim rVariance As Single, gVariance As Single, bVariance As Single, netVariance As Single
            
            'Find the largest total variance in the current stack collection
            For i = 0 To stackCount - 1
                pxStack(i).GetVariance rVariance, gVariance, bVariance
                netVariance = rVariance + gVariance + bVariance
                If (netVariance > maxVariance) Then
                    mvIndex = i
                    maxVariance = netVariance
                End If
            Next i
            
            'Ask the stack with the largest variance to split itself in half.  (Note that the stack object
            ' itself will figure out which axis is most beneficial for splitting.)
            pxStack(mvIndex).Split pxStack(stackCount)
            stackCount = stackCount + 1
        
        'Continue splitting stacks until we arrive at the desired number of colors.  (Each stack represents
        ' one color in the final palette.)
        Loop While (stackCount < numOfColors)
        
        'We now have [numOfColors] unique color stacks.  Each of these represents a set of similar colors.
        ' Generate a final palette by requesting the weighted average of each stack.
        Dim newR As Long, newG As Long, newB As Long
        
        ReDim dstPalette(0 To numOfColors - 1) As RGBQUAD
        For i = 0 To numOfColors - 1
            pxStack(i).GetAverageColor newR, newG, newB
            dstPalette(i).Red = newR
            dstPalette(i).Green = newG
            dstPalette(i).Blue = newB
        Next i
        
        GetOptimizedPalette = True
        
    'If there are less than [numOfColors] unique colors in the image, simply copy the existing stack into a palette
    Else
        pxStack(0).CopyStackToRGBQuad dstPalette
        GetOptimizedPalette = True
    End If
    
End Function

'Given a source palette (ideally created by GetOptimizedPalette(), above), apply said palette to the target image.
' Dithering is *not* used.  Colors are matched exhaustively, meaning this function is slow but produces the smallest
' possible RMSD result for this palette (when matching in the RGB color space, anyway).
Public Function ApplyPaletteToImage(ByRef dstDIB As pdDIB, ByRef srcPalette() As RGBQUAD) As Boolean

    Dim srcPixels() As Byte, tmpSA As SAFEARRAY2D
    dstDIB.WrapArrayAroundDIB srcPixels, tmpSA
    
    Dim pxSize As Long
    pxSize = dstDIB.GetDIBColorDepth \ 8
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = dstDIB.GetDIBStride - 1
    finalY = dstDIB.GetDIBHeight - 1
    
    'We'll use basic RLE acceleration to try and skip palette matching for long runs of contiguous colors
    Dim lastColor As Long: lastColor = -1
    Dim lastPaletteColor As Long
    Dim r As Long, g As Long, b As Long
    Dim i As Long
    Dim minDistance As Single, calcDistance As Single, minIndex As Long
    Dim rDist As Long, gDist As Long, bDist As Long
    Dim numOfColors As Long
    numOfColors = UBound(srcPalette)
    
    For y = 0 To finalY
    For x = 0 To finalX Step pxSize
        b = srcPixels(x, y)
        g = srcPixels(x + 1, y)
        r = srcPixels(x + 2, y)
        
        'If this color matches the last color we tested, reuse our previous palette match
        If (RGB(r, g, b) <> lastColor) Then
            
            'Find the closest color in the current list, using basic Euclidean distance to compare colors
            minIndex = 0
            minDistance = 9.999999E+16
            
            For i = 0 To numOfColors
                With srcPalette(i)
                    rDist = r - .Red
                    gDist = g - .Green
                    bDist = b - .Blue
                End With
                calcDistance = (rDist * rDist) * CUSTOM_WEIGHT_RED + (gDist * gDist) * CUSTOM_WEIGHT_GREEN + (bDist * bDist) * CUSTOM_WEIGHT_BLUE
                If (calcDistance < minDistance) Then
                    minDistance = calcDistance
                    minIndex = i
                End If
            Next i
            
            lastColor = RGB(r, g, b)
            lastPaletteColor = minIndex
            
        Else
            minIndex = lastPaletteColor
        End If
        
        'Apply this color to the target image
        srcPixels(x, y) = srcPalette(minIndex).Blue
        srcPixels(x + 1, y) = srcPalette(minIndex).Green
        srcPixels(x + 2, y) = srcPalette(minIndex).Red
        
    Next x
    Next y
    
    dstDIB.UnwrapArrayFromDIB srcPixels
    
    ApplyPaletteToImage = True
    
End Function

'Given a source palette (ideally created by GetOptimizedPalette(), above), apply said palette to the target image.
' Dithering is *not* used.  Colors are matched using an optimized bucket-search strategy (where the palette is
' pre-sorted by distance from black, and color-matching only occurs across a small range of neighboring buckets).
' Increasing the bucket count improves performance at some trade-off to quality, while the opposite occurs when
' decreasing bucket count.  If the source palette is small (e.g. 32 colors or less), you'd be better off just
' calling the lossless ApplyPaletteToImage() or ApplyPaletteToImage_SysAPI() functions, as this function won't
' provide much of a performance gain, and you risk potentially mismatched colors because the optimization range
' for the hash table is so small.
Public Function ApplyPaletteToImage_LossyHashTable(ByRef dstDIB As pdDIB, ByRef srcPalette() As RGBQUAD, Optional ByVal numOfBuckets As Long = 16) As Boolean

    Dim srcPixels() As Byte, tmpSA As SAFEARRAY2D
    dstDIB.WrapArrayAroundDIB srcPixels, tmpSA
    
    Dim pxSize As Long
    pxSize = dstDIB.GetDIBColorDepth \ 8
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = dstDIB.GetDIBStride - 1
    finalY = dstDIB.GetDIBHeight - 1
    
    'As with normal palette matching, we'll use basic RLE acceleration to try and skip palette
    ' searching for contiguous matching colors.
    Dim lastColor As Long: lastColor = -1
    Dim lastPaletteColor As Long
    Dim r As Long, g As Long, b As Long
    Dim i As Long
    Dim minDistance As Single, calcDistance As Single, lastDistance As Single, minIndex As Long
    Dim rDist As Long, gDist As Long, bDist As Long
    Dim numOfColors As Long
    numOfColors = UBound(srcPalette)
    Dim missCount As Long
    
    'Start by sorting the palette by each color's distance from black.  A specially designed QuickSort
    ' function is used for the sort.
    Dim pSort() As PaletteSort
    ReDim pSort(0 To numOfColors) As PaletteSort
    
    For x = 0 To numOfColors
        pSort(x).pOrigIndex = x
        r = srcPalette(x).Red
        g = srcPalette(x).Green
        b = srcPalette(x).Blue
        pSort(x).pSortCriteria = r * r + g * g + b * b
    Next x
    
    SortPalette pSort
    
    'pSort now represents the final, sorted palette.  Instead of applying the sort results to the palette
    ' entries themselves, we will instead use the pSort data directly, as it maintains its mapping into
    ' the original palette indices.
    
    'From the pSort list, calculate [numOfBuckets] unique "buckets".  Each bucket will contain a series of
    ' similarly-distanced colors.  During the palette matching step, we will only match colors in a small
    ' range of related buckets, which reduces our search space significantly (thus improving performance).
    Dim bucketSize As Long
    numOfBuckets = (numOfColors + 1) \ numOfBuckets - 1
    bucketSize = (numOfColors + 1) \ (numOfBuckets + 1)
    Dim bucketList() As Single, bucketCount() As Long
    ReDim bucketList(0 To numOfBuckets) As Single
    ReDim bucketCount(0 To numOfBuckets) As Long
    For x = 0 To numOfColors - 1
        r = x \ bucketSize
        bucketList(r) = bucketList(r) + pSort(x).pSortCriteria
        bucketCount(r) = bucketCount(r) + 1
    Next x
    
    'Normalize bucket range values
    For x = 0 To numOfBuckets
        bucketList(x) = bucketList(x) \ bucketCount(x)
    Next x
    
    Erase bucketCount
    
    'Start matching pixels
    Dim startSearch As Long, endSearch As Long
    
    For y = 0 To finalY
    For x = 0 To finalX Step pxSize
    
        b = srcPixels(x, y)
        g = srcPixels(x + 1, y)
        r = srcPixels(x + 2, y)
        
        'If this pixel matches the last pixel we tested, reuse our previous match results
        If (RGB(r, g, b) <> lastColor) Then
            
            'Find the bucket with the average distance from black closest to this color's distance
            ' from black.
            calcDistance = r * r + g * g + b * b
            minDistance = Abs(bucketList(0) - calcDistance)
            minIndex = 0
            
            For i = 1 To numOfBuckets
                lastDistance = Abs(bucketList(i) - calcDistance)
                If (lastDistance < minDistance) Then
                    minDistance = lastDistance
                    minIndex = i
                End If
            Next i
            
            'From the identified bucket, determine a range of palette entries to search.
            ' (Note that the range is necessarily larger than just a single bucket; this is necessary
            '  because the RGB color space is not perceptually uniform.)
            If (minIndex = numOfBuckets) Then
                startSearch = numOfColors - bucketSize * 2
                endSearch = numOfColors
            ElseIf (minIndex = 0) Then
                startSearch = 0
                endSearch = bucketSize * 2
            Else
                startSearch = (minIndex * (bucketSize - 2))
                endSearch = (minIndex * (bucketSize + 2))
            End If
            
            If (startSearch < 0) Then startSearch = 0
            If (endSearch > numOfColors) Then endSearch = numOfColors
            
            'Now that we have a range of possible colors to search, look for the closest color match
            ' inside that range of palette entries.
            minIndex = 0
            minDistance = 9.99999E+15
            
            For i = startSearch To endSearch
                With srcPalette(pSort(i).pOrigIndex)
                    rDist = r - .Red
                    gDist = g - .Green
                    bDist = b - .Blue
                End With
                calcDistance = (rDist * rDist) * CUSTOM_WEIGHT_RED + (gDist * gDist) * CUSTOM_WEIGHT_GREEN + (bDist * bDist) * CUSTOM_WEIGHT_BLUE
                If (calcDistance < minDistance) Then
                    minDistance = calcDistance
                    minIndex = pSort(i).pOrigIndex
                End If
            Next i
            
            lastColor = RGB(r, g, b)
            lastPaletteColor = minIndex
            
        Else
            minIndex = lastPaletteColor
        End If
        
        'Apply the closest discovered color to this pixel.
        srcPixels(x, y) = srcPalette(minIndex).Blue
        srcPixels(x + 1, y) = srcPalette(minIndex).Green
        srcPixels(x + 2, y) = srcPalette(minIndex).Red
        
    Next x
    Next y
    
    dstDIB.UnwrapArrayFromDIB srcPixels
    
    ApplyPaletteToImage_LossyHashTable = True
    
End Function

'Use QuickSort to sort a palette.  The srcPaletteSort must be assembled by the caller, with the .pSortCriteria
' filled with a Single that represents "color order".  (Not defining this strictly allows for many different types
' of palette sorts, based on the caller's needs.)
Private Sub SortPalette(ByRef srcPaletteSort() As PaletteSort)
    SortInner srcPaletteSort, 0, UBound(srcPaletteSort)
End Sub

'Basic QuickSort function.  Recursive calls will sort the palette on the range [lowVal, highVal].  The first
' call must be on the range [0, UBound(srcPaletteSort)].
Private Sub SortInner(ByRef srcPaletteSort() As PaletteSort, ByVal lowVal As Long, ByVal highVal As Long)
    
    'Ignore the search request if the bounds are mismatched
    If (lowVal < highVal) Then
        
        'Sort some sub-portion of the list, use the returned pivot to repeat the sort process
        Dim j As Long
        j = SortPartition(srcPaletteSort, lowVal, highVal)
        SortInner srcPaletteSort, lowVal, j - 1
        SortInner srcPaletteSort, j + 1, highVal
    End If
    
End Sub

'Basic QuickSort partition function.  All values in the range [lowVal, highVal] are sorted against a pivot value, j.
' The final pivot position is returned, and our caller can use that to request two new sorts on either side of the pivot.
Private Function SortPartition(ByRef srcPaletteSort() As PaletteSort, ByVal lowVal As Long, ByVal highVal As Long) As Long
    
    Dim i As Long, j As Long
    i = lowVal
    j = highVal + 1
    
    Dim v As Single
    v = srcPaletteSort(lowVal).pSortCriteria
    
    Dim tmpSort As PaletteSort
    
    Do
        
        'Compare the pivot against points beneath it
        Do
            i = i + 1
            If (i = highVal) Then Exit Do
        Loop While (srcPaletteSort(i).pSortCriteria < v)
        
        'Compare the pivot against points above it
        Do
            j = j - 1
            
            'A failsafe exit check here would be redundant, since we already check this state above
            'If (j = lowVal) Then Exit Do
        Loop While (v < srcPaletteSort(j).pSortCriteria)
        
        'If the pivot has arrived at its final location, exit
        If (i >= j) Then Exit Do
        
        'Swap the values at indexes i and j
        tmpSort = srcPaletteSort(j)
        srcPaletteSort(j) = srcPaletteSort(i)
        srcPaletteSort(i) = tmpSort
        
    Loop
    
    'Move the pivot value into its final location
    tmpSort = srcPaletteSort(j)
    srcPaletteSort(j) = srcPaletteSort(lowVal)
    srcPaletteSort(lowVal) = tmpSort
    
    'Return the pivot's final position
    SortPartition = j
    
End Function

'Given a source palette (ideally created by GetOptimizedPalette(), above), apply said palette to the target image.
' Dithering is *not* used.  Colors are matched using an octree-search strategy (where the palette is pre-loaded
' into an octree, and colors are matched via that tree).  If the palette is known to be small (e.g. 32 colors or less),
' you'd be better off just calling the normal ApplyPaletteToImage function, as this function won't provide much of a
' performance gain.
Public Function ApplyPaletteToImage_Octree(ByRef dstDIB As pdDIB, ByRef srcPalette() As RGBQUAD) As Boolean

    Dim srcPixels() As Byte, tmpSA As SAFEARRAY2D
    dstDIB.WrapArrayAroundDIB srcPixels, tmpSA
    
    Dim pxSize As Long
    pxSize = dstDIB.GetDIBColorDepth \ 8
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = dstDIB.GetDIBStride - 1
    finalY = dstDIB.GetDIBHeight - 1
    
    'As with normal palette matching, we'll use basic RLE acceleration to try and skip palette
    ' searching for contiguous matching colors.
    Dim lastColor As Long: lastColor = -1
    Dim minIndex As Long, lastPaletteColor As Long
    Dim r As Long, g As Long, b As Long
    
    Dim tmpQuad As RGBQUAD
        
    'Build the initial tree
    Dim cOctree As pdColorSearch
    Set cOctree = New pdColorSearch
    cOctree.CreateColorTree srcPalette
    
    'Octrees tend to make colors darker, because they match colors bit-by-bit, meaning dark colors are
    ' preferentially matched over light ones.  (e.g. &h0111 would match to &h0000 rather than &h1000,
    ' because bits are matched in most-significant to least-significant order).
    
    'To reduce the impact this has on the final image, I've considered artifically brightening colors
    ' before matching them.  The problem is that we really only need to do this around power-of-two
    ' values, and mathematically, I'm not sure how to do this most efficiently (e.g. without just making
    ' colors biased against brighter matches instead).
    
    'As such, I've marked this as "TODO" for now.
    'Dim octHelper() As Byte
    'ReDim octHelper(0 To 255) As Byte
    'For x = 0 To 255
    '    r = x + 10
    '    If (r > 255) Then octHelper(x) = 255 Else octHelper(x) = r
    'Next x
    '(Obviously, for this to work, you'd need to updated the tmpQuad assignments in the inner loop, below.)
    
    'Start matching pixels
    For y = 0 To finalY
    For x = 0 To finalX Step pxSize
    
        b = srcPixels(x, y)
        g = srcPixels(x + 1, y)
        r = srcPixels(x + 2, y)
        
        'If this pixel matches the last pixel we tested, reuse our previous match results
        If (RGB(r, g, b) <> lastColor) Then
            
            tmpQuad.Red = r
            tmpQuad.Green = g
            tmpQuad.Blue = b
            
            'Ask the octree to find the best match
            minIndex = cOctree.GetNearestPaletteIndex(tmpQuad)
            
            lastColor = RGB(r, g, b)
            lastPaletteColor = minIndex
            
        Else
            minIndex = lastPaletteColor
        End If
        
        'Apply the closest discovered color to this pixel.
        srcPixels(x, y) = srcPalette(minIndex).Blue
        srcPixels(x + 1, y) = srcPalette(minIndex).Green
        srcPixels(x + 2, y) = srcPalette(minIndex).Red
        
    Next x
    Next y
    
    dstDIB.UnwrapArrayFromDIB srcPixels
    
    ApplyPaletteToImage_Octree = True
    
End Function

'Given a source palette (ideally created by GetOptimizedPalette(), above), apply said palette to the target image.
' Dithering is *not* used.  Colors are matched using System APIs.
Public Function ApplyPaletteToImage_SysAPI(ByRef dstDIB As pdDIB, ByRef srcPalette() As RGBQUAD) As Boolean

    Dim srcPixels() As Byte, tmpSA As SAFEARRAY2D
    dstDIB.WrapArrayAroundDIB srcPixels, tmpSA
    
    Dim pxSize As Long
    pxSize = dstDIB.GetDIBColorDepth \ 8
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = dstDIB.GetDIBStride - 1
    finalY = dstDIB.GetDIBHeight - 1
    
    'As with normal palette matching, we'll use basic RLE acceleration to try and skip palette
    ' searching for contiguous matching colors.
    Dim lastColor As Long: lastColor = -1
    Dim minIndex As Long, lastPaletteColor As Long
    Dim r As Long, g As Long, b As Long
    
    Dim tmpPalette As GDI_LOGPALETTE256
    tmpPalette.palNumEntries = UBound(srcPalette) + 1
    tmpPalette.palVersion = &H300
    Dim i As Long
    For i = 0 To UBound(srcPalette)
        tmpPalette.palEntry(i).peR = srcPalette(i).Red
        tmpPalette.palEntry(i).peG = srcPalette(i).Green
        tmpPalette.palEntry(i).peB = srcPalette(i).Blue
    Next i
    
    Dim hPal As Long
    hPal = CreatePalette(VarPtr(tmpPalette))
    
    'Start matching pixels
    For y = 0 To finalY
    For x = 0 To finalX Step pxSize
    
        b = srcPixels(x, y)
        g = srcPixels(x + 1, y)
        r = srcPixels(x + 2, y)
        
        'If this pixel matches the last pixel we tested, reuse our previous match results
        If (RGB(r, g, b) <> lastColor) Then
            
            'Ask the system to find the nearest color
            minIndex = GetNearestPaletteIndex(hPal, RGB(r, g, b))
            
            lastColor = RGB(r, g, b)
            lastPaletteColor = minIndex
            
        Else
            minIndex = lastPaletteColor
        End If
        
        'Apply the closest discovered color to this pixel.
        srcPixels(x, y) = srcPalette(minIndex).Blue
        srcPixels(x + 1, y) = srcPalette(minIndex).Green
        srcPixels(x + 2, y) = srcPalette(minIndex).Red
        
    Next x
    Next y
    
    dstDIB.UnwrapArrayFromDIB srcPixels
    
    If (hPal <> 0) Then DeleteObject hPal
    
    ApplyPaletteToImage_SysAPI = True
    
End Function
