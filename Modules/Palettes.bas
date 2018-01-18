Attribute VB_Name = "Palettes"
'***************************************************************************
'PhotoDemon's Master Palette Interface
'Copyright 2017-2018 by Tanner Helland
'Created: 12/January/17
'Last updated: 17/January/18
'Last update: add new "Palette import dialog" helper function, which will be useful as we add support for more
'             palette file formats
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
Private Declare Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" (ByVal dstPointer As Long, ByVal Length As Long, ByVal Fill As Byte)

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
' version of itself.  I built a function specifically for this: DIBs.ResizeDIBByPixelCount().  That function
' resizes an image to a target pixel count, and I wouldn't recommend a net size any larger than ~50,000 pixels.
Public Function GetOptimizedPalette(ByRef srcDIB As pdDIB, ByRef dstPalette() As RGBQuad, Optional ByVal numOfColors As Long = 256) As Boolean
    
    'Do not request less than two colors in the final palette!
    If (numOfColors < 2) Then numOfColors = 2
    
    Dim srcPixels() As Byte, tmpSA As SafeArray2D
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
    
    'Note that PD actually supports quite a few different quantization methods.  At present, only the
    ' highest-quality "variance + median split" algorithm is used, however.
    pxStack(0).SetQuantizeMode pdqs_VarPlusMedian
    
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
        
        Dim maxVariance As Single, mvIndex As Long
        Dim i As Long
        
        'With the initial stack constructed, we can now start partitioning it into smaller stacks based on variance
        Do
        
            'Reset maximum variance (because we need to calculate it anew)
            maxVariance = 0#
            
            Dim rVariance As Single, gVariance As Single, bVariance As Single, netVariance As Single
            
            'Find the largest total variance in the current stack collection
            For i = 0 To stackCount - 1
            
                pxStack(i).GetVariance rVariance, gVariance, bVariance
                
                'There are actually two ways to handle this problem.  We can find the net variance (e.g. the
                ' block with the most varied set of colors), or we can find the block with the most varied
                ' *channel* (e.g. if two channels are identical, but the third is hugely varied, treat that
                ' block as the highest variance).
                '
                'I don't have a theoretical framework for determining the better of these two solutions,
                ' but a large amount of trial-and-error leads me to believe that it's best to split according
                ' to net variance.  (Note that the block will still be split along its single channel
                ' with highest variance, regardless.)
                netVariance = rVariance + gVariance + bVariance
                If (netVariance > maxVariance) Then
                    mvIndex = i
                    maxVariance = netVariance
                End If
                
                'Per-channel formula follows:
                'If (rVariance > maxVariance) Then
                '    maxVariance = rVariance
                '    mvIndex = i
                'End If
                'If (gVariance > maxVariance) Then
                '    maxVariance = gVariance
                '    mvIndex = i
                'End If
                'If (bVariance > maxVariance) Then
                '    maxVariance = bVariance
                '    mvIndex = i
                'End If
                
            Next i
            
            'Ask the stack with the largest variance to split itself in half.  (Note that the stack object
            ' itself will figure out which axis is most appropriate for splitting.)
            'Debug.Print "Largest variance was " & maxVariance & ", found in stack #" & mvIndex & " (total stack count is " & stackCount & ")"
            pxStack(mvIndex).Split pxStack(stackCount)
            stackCount = stackCount + 1
        
        'Continue splitting stacks until we arrive at the desired number of colors.  (Each stack represents
        ' one color in the final palette.)
        Loop While (stackCount < numOfColors)
        
        'We now have [numOfColors] unique color stacks.  Each of these represents a set of similar colors.
        ' Generate a final palette by requesting the weighted average of each stack.
        Dim newR As Long, newG As Long, newB As Long
        
        ReDim dstPalette(0 To numOfColors - 1) As RGBQuad
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

'Given a palette, make sure black and white exist.  This function scans the palette and replaces the darkest
' entry with black, and the brightest entry with white.  (We use this approach so that we can accept palettes
' from any source, even ones that have already contain 256 entries.)  No difference is made to the palette
' if it already contains black and white.
Public Function EnsureBlackAndWhiteInPalette(ByRef srcPalette() As RGBQuad, Optional ByRef srcDIB As pdDIB = Nothing) As Boolean
    
    Dim minLuminance As Long, minLuminanceIndex As Long
    Dim maxLuminance As Long, maxLuminanceIndex As Long
    
    Dim pBoundL As Long, pBoundU As Long
    pBoundL = LBound(srcPalette)
    pBoundU = UBound(srcPalette)
    
    If (pBoundL <> pBoundU) Then
    
        With srcPalette(pBoundL)
            minLuminance = Colors.GetHQLuminance(.Red, .Green, .Blue)
            minLuminanceIndex = pBoundL
            maxLuminance = Colors.GetHQLuminance(.Red, .Green, .Blue)
            maxLuminanceIndex = pBoundL
        End With
        
        Dim testLuminance As Long
        
        Dim i As Long
        For i = pBoundL + 1 To pBoundU
        
            With srcPalette(i)
                testLuminance = Colors.GetHQLuminance(.Red, .Green, .Blue)
            End With
            
            If (testLuminance > maxLuminance) Then
                maxLuminance = testLuminance
                maxLuminanceIndex = i
            ElseIf (testLuminance < minLuminance) Then
                minLuminance = testLuminance
                minLuminanceIndex = i
            End If
            
        Next i
        
        Dim preserveWhite As Boolean, preserveBlack As Boolean
        preserveWhite = True
        preserveBlack = True
        
        'If the caller passed us an image, see if the image contains black and/or white.  If it does *not*,
        ' we won't worry about preserving that particular color
        If (Not srcDIB Is Nothing) Then
        
            Dim srcPixels() As Byte, tmpSA As SafeArray2D
            srcDIB.WrapArrayAroundDIB srcPixels, tmpSA
            
            Dim pxSize As Long
            pxSize = srcDIB.GetDIBColorDepth \ 8
            
            Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
            initX = 0
            initY = 0
            finalX = srcDIB.GetDIBStride - 1
            finalY = srcDIB.GetDIBHeight - 1
            
            Dim r As Long, g As Long, b As Long
            Dim blackFound As Boolean, whiteFound As Boolean
            
            For y = 0 To finalY
            For x = 0 To finalX Step pxSize
                b = srcPixels(x, y)
                g = srcPixels(x + 1, y)
                r = srcPixels(x + 2, y)
                
                If (Not blackFound) Then
                    If (r = 0) And (g = 0) And (b = 0) Then blackFound = True
                End If
                
                If (Not whiteFound) Then
                    If (r = 255) And (g = 255) And (b = 255) Then whiteFound = True
                End If
                
                If (blackFound And whiteFound) Then Exit For
            Next x
                If (blackFound And whiteFound) Then Exit For
            Next y
            
            srcDIB.UnwrapArrayFromDIB srcPixels
            
            preserveBlack = blackFound
            preserveWhite = whiteFound
    
        End If
        
        If preserveBlack Then
            With srcPalette(minLuminanceIndex)
                .Red = 0
                .Green = 0
                .Blue = 0
            End With
        End If
        
        If preserveWhite Then
            With srcPalette(maxLuminanceIndex)
                .Red = 255
                .Green = 255
                .Blue = 255
            End With
        End If
        
        EnsureBlackAndWhiteInPalette = True
        
    Else
        EnsureBlackAndWhiteInPalette = False
    End If

End Function

'Given a source palette (ideally created by GetOptimizedPalette(), above), apply said palette to the target image.
' Dithering is *not* used.  Colors are matched exhaustively, meaning this function is slow but produces the smallest
' possible RMSD result for this palette (when matching in the RGB color space, anyway).
Public Function ApplyPaletteToImage(ByRef dstDIB As pdDIB, ByRef srcPalette() As RGBQuad) As Boolean

    Dim srcPixels() As Byte, tmpSA As SafeArray2D
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
Public Function ApplyPaletteToImage_LossyHashTable(ByRef dstDIB As pdDIB, ByRef srcPalette() As RGBQuad, Optional ByVal numOfBuckets As Long = 16) As Boolean

    Dim srcPixels() As Byte, tmpSA As SafeArray2D
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
        
        'Sort some sub-portion of the list, and use the returned pivot to repeat the sort process
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
Public Function ApplyPaletteToImage_Octree(ByRef dstDIB As pdDIB, ByRef srcPalette() As RGBQuad) As Boolean

    Dim srcPixels() As Byte, tmpSA As SafeArray2D
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
    
    Dim tmpQuad As RGBQuad
        
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
Public Function ApplyPaletteToImage_SysAPI(ByRef dstDIB As pdDIB, ByRef srcPalette() As RGBQuad) As Boolean

    Dim srcPixels() As Byte, tmpSA As SafeArray2D
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

'Given a source palette (ideally created by GetOptimizedPalette(), above), apply said palette to the target image.
' Dithering *is* used.  Colors are matched using System APIs.
Public Function ApplyPaletteToImage_Dithered(ByRef dstDIB As pdDIB, ByRef srcPalette() As RGBQuad, Optional ByVal DitherMethod As PD_DITHER_METHOD = PDDM_FloydSteinberg, Optional ByVal reduceBleed As Boolean = False) As Boolean

    Dim srcPixels() As Byte, tmpSA As SafeArray2D
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
    Dim minIndex As Long
    Dim r As Long, g As Long, b As Long
    
    Dim tmpPalette As GDI_LOGPALETTE256
    tmpPalette.palNumEntries = UBound(srcPalette) + 1
    tmpPalette.palVersion = &H300
    Dim i As Long, j As Long
    For i = 0 To UBound(srcPalette)
        tmpPalette.palEntry(i).peR = srcPalette(i).Red
        tmpPalette.palEntry(i).peG = srcPalette(i).Green
        tmpPalette.palEntry(i).peB = srcPalette(i).Blue
    Next i
    
    Dim hPal As Long
    hPal = CreatePalette(VarPtr(tmpPalette))
    
    'Prep a dither table that matches the requested setting.  Note that ordered dithers are handled separately.
    Dim DitherTable() As Long
    Dim orderedDitherInUse As Boolean
    orderedDitherInUse = CBool(DitherMethod = PDDM_Ordered_Bayer4x4) Or CBool(DitherMethod = PDDM_Ordered_Bayer8x8)
    
    If orderedDitherInUse Then
    
        'Ordered dithers are handled specially, because we don't need to track running errors (e.g. no dithering
        ' information is carried to neighboring pixels).  Instead, we simply use the dither tables to adjust our
        ' threshold values on-the-fly.
        Dim ditherRows As Long, ditherColumns As Long
        
        If (DitherMethod = PDDM_Ordered_Bayer4x4) Then
            
            'First, prepare a Bayer dither table
            ditherRows = 3
            ditherColumns = 3
            ReDim DitherTable(0 To ditherRows, 0 To ditherColumns) As Long
            
            DitherTable(0, 0) = 1
            DitherTable(0, 1) = 9
            DitherTable(0, 2) = 3
            DitherTable(0, 3) = 11
            
            DitherTable(1, 0) = 13
            DitherTable(1, 1) = 5
            DitherTable(1, 2) = 15
            DitherTable(1, 3) = 7
            
            DitherTable(2, 0) = 4
            DitherTable(2, 1) = 12
            DitherTable(2, 2) = 2
            DitherTable(2, 3) = 10
            
            DitherTable(3, 0) = 16
            DitherTable(3, 1) = 8
            DitherTable(3, 2) = 14
            DitherTable(3, 3) = 6
    
            'Convert the dither entries to absolute offsets (meaning half are positive, half are negative)
            For x = 0 To 3
            For y = 0 To 3
                DitherTable(x, y) = DitherTable(x, y) * 2 - 16
            Next y
            Next x
            
        ElseIf (DitherMethod = PDDM_Ordered_Bayer8x8) Then
        
            'First, prepare a Bayer dither table
            ditherRows = 7
            ditherColumns = 7
            ReDim DitherTable(0 To ditherRows, 0 To ditherColumns) As Long
            
            DitherTable(0, 0) = 1
            DitherTable(0, 1) = 49
            DitherTable(0, 2) = 13
            DitherTable(0, 3) = 61
            DitherTable(0, 4) = 4
            DitherTable(0, 5) = 52
            DitherTable(0, 6) = 16
            DitherTable(0, 7) = 64
            
            DitherTable(1, 0) = 33
            DitherTable(1, 1) = 17
            DitherTable(1, 2) = 45
            DitherTable(1, 3) = 29
            DitherTable(1, 4) = 36
            DitherTable(1, 5) = 20
            DitherTable(1, 6) = 48
            DitherTable(1, 7) = 32
            
            DitherTable(2, 0) = 9
            DitherTable(2, 1) = 57
            DitherTable(2, 2) = 5
            DitherTable(2, 3) = 53
            DitherTable(2, 4) = 12
            DitherTable(2, 5) = 60
            DitherTable(2, 6) = 8
            DitherTable(2, 7) = 56
            
            DitherTable(3, 0) = 41
            DitherTable(3, 1) = 25
            DitherTable(3, 2) = 37
            DitherTable(3, 3) = 21
            DitherTable(3, 4) = 44
            DitherTable(3, 5) = 28
            DitherTable(3, 6) = 40
            DitherTable(3, 7) = 24
            
            DitherTable(4, 0) = 3
            DitherTable(4, 1) = 51
            DitherTable(4, 2) = 15
            DitherTable(4, 3) = 63
            DitherTable(4, 4) = 2
            DitherTable(4, 5) = 50
            DitherTable(4, 6) = 14
            DitherTable(4, 7) = 62
            
            DitherTable(5, 0) = 35
            DitherTable(5, 1) = 19
            DitherTable(5, 2) = 47
            DitherTable(5, 3) = 31
            DitherTable(5, 4) = 34
            DitherTable(5, 5) = 18
            DitherTable(5, 6) = 46
            DitherTable(5, 7) = 30
    
            DitherTable(6, 0) = 11
            DitherTable(6, 1) = 59
            DitherTable(6, 2) = 7
            DitherTable(6, 3) = 55
            DitherTable(6, 4) = 10
            DitherTable(6, 5) = 58
            DitherTable(6, 6) = 6
            DitherTable(6, 7) = 54
            
            DitherTable(7, 0) = 43
            DitherTable(7, 1) = 27
            DitherTable(7, 2) = 39
            DitherTable(7, 3) = 23
            DitherTable(7, 4) = 42
            DitherTable(7, 5) = 26
            DitherTable(7, 6) = 38
            DitherTable(7, 7) = 22
            
            'Convert the dither entries to 255-based values
            For x = 0 To 7
            For y = 0 To 7
                DitherTable(x, y) = DitherTable(x, y) - 32
            Next y
            Next x
        
        End If
        
        'Apply the finished dither table to the image
        Dim ditherAmt As Long
        
        For y = 0 To finalY
        For x = 0 To finalX Step pxSize
        
            b = srcPixels(x, y)
            g = srcPixels(x + 1, y)
            r = srcPixels(x + 2, y)
            
            'Add dither to each component
            ditherAmt = DitherTable((x \ 4) And ditherRows, y And ditherColumns)
            If reduceBleed Then ditherAmt = ditherAmt * 0.33
            
            r = r + ditherAmt
            If (r > 255) Then
                r = 255
            ElseIf (r < 0) Then
                r = 0
            End If
            
            g = g + ditherAmt
            If (g > 255) Then
                g = 255
            ElseIf (g < 0) Then
                g = 0
            End If
            
            b = b + ditherAmt
            If (b > 255) Then
                b = 255
            ElseIf (b < 0) Then
                b = 0
            End If
            
            'Ask the system to find the nearest color
            minIndex = GetNearestPaletteIndex(hPal, RGB(r, g, b))
            
            srcPixels(x, y) = srcPalette(minIndex).Blue
            srcPixels(x + 1, y) = srcPalette(minIndex).Green
            srcPixels(x + 2, y) = srcPalette(minIndex).Red
            
        Next x
        Next y
    
    'All error-diffusion dither methods are handled similarly
    Else
        
        Dim ditherTableI() As Byte
        Dim xLeft As Long, xRight As Long, yDown As Long
        Dim rError As Long, gError As Long, bError As Long
        Dim errorMult As Single
        Dim ditherDivisor As Single
        
        'Retrieve a hard-coded dithering table matching the requested dither type
        Palettes.GetDitherTable DitherMethod, ditherTableI, ditherDivisor, xLeft, xRight, yDown
        
        'Next, build an error tracking array.  Some diffusion methods require three rows worth of others;
        ' others require two.  Note that errors must be tracked separately for each color component.
        Dim xWidth As Long
        xWidth = workingDIB.GetDIBWidth - 1
        Dim rErrors() As Single, gErrors() As Single, bErrors() As Single
        ReDim rErrors(0 To xWidth, 0 To yDown) As Single
        ReDim gErrors(0 To xWidth, 0 To yDown) As Single
        ReDim bErrors(0 To xWidth, 0 To yDown) As Single
        
        Dim xNonStride As Long, xQuickInner As Long
        Dim newR As Long, newG As Long, newB As Long
        
        'Start calculating pixels.
        For y = 0 To finalY
        For x = 0 To finalX Step pxSize
        
            b = srcPixels(x, y)
            g = srcPixels(x + 1, y)
            r = srcPixels(x + 2, y)
            
            'Add our running errors to the original colors
            xNonStride = x * 0.25
            newR = r + rErrors(xNonStride, 0)
            newG = g + gErrors(xNonStride, 0)
            newB = b + bErrors(xNonStride, 0)
            
            If (newR > 255) Then
                newR = 255
            ElseIf (newR < 0) Then
                newR = 0
            End If
            
            If (newG > 255) Then
                newG = 255
            ElseIf (newG < 0) Then
                newG = 0
            End If
            
            If (newB > 255) Then
                newB = 255
            ElseIf (newB < 0) Then
                newB = 0
            End If
            
            'Find the best palette match
            minIndex = GetNearestPaletteIndex(hPal, RGB(newR, newG, newB))
            
            With srcPalette(minIndex)
            
                'Apply the closest discovered color to this pixel.
                srcPixels(x, y) = .Blue
                srcPixels(x + 1, y) = .Green
                srcPixels(x + 2, y) = .Red
            
                'Calculate new errors
                rError = r - CLng(.Red)
                gError = g - CLng(.Green)
                bError = b - CLng(.Blue)
                
            End With
            
            'Reduce color bleed, if specified
            If reduceBleed Then
                rError = rError * 0.33
                gError = gError * 0.33
                bError = bError * 0.33
            End If
            
            'Spread any remaining error to neighboring pixels, using the precalculated dither table as our guide
            For i = xLeft To xRight
            For j = 0 To yDown
                
                'First, ignore already processed pixels
                If (j = 0) And (i <= 0) Then GoTo NextDitheredPixel
                    
                'Second, ignore pixels that have a zero in the dither table
                If (ditherTableI(i, j) = 0) Then GoTo NextDitheredPixel
                    
                xQuickInner = xNonStride + i
                
                'Next, ignore target pixels that are off the image boundary
                If (xQuickInner < initX) Then
                    GoTo NextDitheredPixel
                ElseIf (xQuickInner > xWidth) Then
                    GoTo NextDitheredPixel
                End If
                
                'If we've made it all the way here, we are able to actually spread the error to this location
                errorMult = CSng(ditherTableI(i, j)) / ditherDivisor
                rErrors(xQuickInner, j) = rErrors(xQuickInner, j) + (rError * errorMult)
                gErrors(xQuickInner, j) = gErrors(xQuickInner, j) + (gError * errorMult)
                bErrors(xQuickInner, j) = bErrors(xQuickInner, j) + (bError * errorMult)
                
NextDitheredPixel:
            Next j
            Next i
            
        Next x
        
            'When moving to the next line, we need to "shift" all accumulated errors upward.
            ' (Basically, what was previously the "next" line, is now the "current" line.
            ' The last line of errors must also be zeroed-out.
            CopyMemory ByVal VarPtr(rErrors(0, 0)), ByVal VarPtr(rErrors(0, 1)), (xWidth + 1) * 4
            CopyMemory ByVal VarPtr(gErrors(0, 0)), ByVal VarPtr(gErrors(0, 1)), (xWidth + 1) * 4
            CopyMemory ByVal VarPtr(bErrors(0, 0)), ByVal VarPtr(bErrors(0, 1)), (xWidth + 1) * 4
            
            If (yDown = 1) Then
                FillMemory VarPtr(rErrors(0, 1)), (xWidth + 1) * 4, 0
                FillMemory VarPtr(gErrors(0, 1)), (xWidth + 1) * 4, 0
                FillMemory VarPtr(bErrors(0, 1)), (xWidth + 1) * 4, 0
            Else
                CopyMemory ByVal VarPtr(rErrors(0, 1)), ByVal VarPtr(rErrors(0, 2)), (xWidth + 1) * 4
                CopyMemory ByVal VarPtr(gErrors(0, 1)), ByVal VarPtr(gErrors(0, 2)), (xWidth + 1) * 4
                CopyMemory ByVal VarPtr(bErrors(0, 1)), ByVal VarPtr(bErrors(0, 2)), (xWidth + 1) * 4
                
                FillMemory VarPtr(rErrors(0, 2)), (xWidth + 1) * 4, 0
                FillMemory VarPtr(gErrors(0, 2)), (xWidth + 1) * 4, 0
                FillMemory VarPtr(bErrors(0, 2)), (xWidth + 1) * 4, 0
            End If
            
        Next y
    
    End If
    
    dstDIB.UnwrapArrayFromDIB srcPixels
    
    If (hPal <> 0) Then DeleteObject hPal
    
    ApplyPaletteToImage_Dithered = True
    
End Function

'Populate a dithering table and relevant markers based on a specific dithering type.
' Returns: TRUE if successful; FALSE otherwise.  Note that some dither types (e.g. ordered dithers) do not
' use this function; they are handled specially.
Public Function GetDitherTable(ByVal ditherType As PD_DITHER_METHOD, ByRef dstDitherTable() As Byte, ByRef ditherDivisor As Single, ByRef xLeft As Long, ByRef xRight As Long, ByRef yDown As Long) As Boolean
    
    GetDitherTable = True
    
    Select Case ditherType
    
        Case PDDM_FalseFloydSteinberg
        
            ReDim dstDitherTable(0 To 1, 0 To 1) As Byte
            
            dstDitherTable(1, 0) = 3
            dstDitherTable(0, 1) = 3
            dstDitherTable(1, 1) = 2
            
            ditherDivisor = 8
            
            xLeft = 0
            xRight = 1
            yDown = 1
            
        Case PDDM_FloydSteinberg
        
            ReDim dstDitherTable(-1 To 1, 0 To 1) As Byte
            
            dstDitherTable(1, 0) = 7
            dstDitherTable(-1, 1) = 3
            dstDitherTable(0, 1) = 5
            dstDitherTable(1, 1) = 1
            
            ditherDivisor = 16
        
            xLeft = -1
            xRight = 1
            yDown = 1
            
        Case PDDM_JarvisJudiceNinke
        
            ReDim dstDitherTable(-2 To 2, 0 To 2) As Byte
            
            dstDitherTable(1, 0) = 7
            dstDitherTable(2, 0) = 5
            dstDitherTable(-2, 1) = 3
            dstDitherTable(-1, 1) = 5
            dstDitherTable(0, 1) = 7
            dstDitherTable(1, 1) = 5
            dstDitherTable(2, 1) = 3
            dstDitherTable(-2, 2) = 1
            dstDitherTable(-1, 2) = 3
            dstDitherTable(0, 2) = 5
            dstDitherTable(1, 2) = 3
            dstDitherTable(2, 2) = 1
            
            ditherDivisor = 48
            
            xLeft = -2
            xRight = 2
            yDown = 2
            
        Case PDDM_Stucki
        
            ReDim dstDitherTable(-2 To 2, 0 To 2) As Byte
            
            dstDitherTable(1, 0) = 8
            dstDitherTable(2, 0) = 4
            dstDitherTable(-2, 1) = 2
            dstDitherTable(-1, 1) = 4
            dstDitherTable(0, 1) = 8
            dstDitherTable(1, 1) = 4
            dstDitherTable(2, 1) = 2
            dstDitherTable(-2, 2) = 1
            dstDitherTable(-1, 2) = 2
            dstDitherTable(0, 2) = 4
            dstDitherTable(1, 2) = 2
            dstDitherTable(2, 2) = 1
            
            ditherDivisor = 42
            
            xLeft = -2
            xRight = 2
            yDown = 2
            
        Case PDDM_Burkes
        
            ReDim dstDitherTable(-2 To 2, 0 To 1) As Byte
            
            dstDitherTable(1, 0) = 8
            dstDitherTable(2, 0) = 4
            dstDitherTable(-2, 1) = 2
            dstDitherTable(-1, 1) = 4
            dstDitherTable(0, 1) = 8
            dstDitherTable(1, 1) = 4
            dstDitherTable(2, 1) = 2
            
            ditherDivisor = 32
            
            xLeft = -2
            xRight = 2
            yDown = 1
            
        Case PDDM_Sierra3
        
            ReDim dstDitherTable(-2 To 2, 0 To 2) As Byte
            
            dstDitherTable(1, 0) = 5
            dstDitherTable(2, 0) = 3
            dstDitherTable(-2, 1) = 2
            dstDitherTable(-1, 1) = 4
            dstDitherTable(0, 1) = 5
            dstDitherTable(1, 1) = 4
            dstDitherTable(2, 1) = 2
            dstDitherTable(-2, 2) = 0
            dstDitherTable(-1, 2) = 2
            dstDitherTable(0, 2) = 3
            dstDitherTable(1, 2) = 2
            dstDitherTable(2, 2) = 0
            
            ditherDivisor = 32
            
            xLeft = -2
            xRight = 2
            yDown = 2
            
        Case PDDM_SierraTwoRow
            
            ReDim dstDitherTable(-2 To 2, 0 To 1) As Byte
            
            dstDitherTable(1, 0) = 4
            dstDitherTable(2, 0) = 3
            dstDitherTable(-2, 1) = 1
            dstDitherTable(-1, 1) = 2
            dstDitherTable(0, 1) = 3
            dstDitherTable(1, 1) = 2
            dstDitherTable(2, 1) = 1
            
            ditherDivisor = 16
            
            xLeft = -2
            xRight = 2
            yDown = 1
        
        Case PDDM_SierraLite
        
            ReDim dstDitherTable(-1 To 1, 0 To 1) As Byte
            
            dstDitherTable(1, 0) = 2
            dstDitherTable(-1, 1) = 1
            dstDitherTable(0, 1) = 1
            
            ditherDivisor = 4
            
            xLeft = -1
            xRight = 1
            yDown = 1
            
        Case PDDM_Atkinson
            
            ReDim dstDitherTable(-1 To 2, 0 To 2) As Byte
            
            dstDitherTable(1, 0) = 1
            dstDitherTable(2, 0) = 1
            dstDitherTable(-1, 1) = 1
            dstDitherTable(0, 1) = 1
            dstDitherTable(1, 1) = 1
            dstDitherTable(0, 2) = 1
            
            ditherDivisor = 8
            
            xLeft = -1
            xRight = 2
            yDown = 2
            
        Case Else
            GetDitherTable = False
    
    End Select
    
End Function

'Display PD's generic palette load dialog.  All supported palette filetypes will be available to the user.
Public Function DisplayPaletteLoadDialog(ByRef srcFilename As String, ByRef dstFilename As String) As Boolean
    
    DisplayPaletteLoadDialog = False
    
    'Disable user input until the dialog closes
    Interface.DisableUserInput
    
    Dim cdFilter As String
    cdFilter = g_Language.TranslateMessage("All supported palettes") & "|*.act;*.gpl|"
    cdFilter = cdFilter & g_Language.TranslateMessage("Adobe Color Table") & " (.act)|*.act|"
    cdFilter = cdFilter & g_Language.TranslateMessage("GIMP Palette") & " (.gpl)|*.gpl|"
    cdFilter = cdFilter & g_Language.TranslateMessage("All files") & "|*.*"
    
    Dim cdTitle As String
    cdTitle = g_Language.TranslateMessage("Select a palette")
            
    'Prep a common dialog interface
    Dim openDialog As pdOpenSaveDialog
    Set openDialog = New pdOpenSaveDialog
            
    Dim sFile As String
    sFile = srcFilename
    
    If openDialog.GetOpenFileName(sFile, , True, False, cdFilter, 1, g_UserPreferences.GetPalettePath, cdTitle, , GetModalOwner().hWnd) Then
    
        'By design, we don't perform any validation here.  Let the caller validate the file as much (or as little)
        ' as they require.
        DisplayPaletteLoadDialog = (Len(sFile) <> 0)
        
        'The dialog was successful.  Return the path, and save this path for future usage.
        If DisplayPaletteLoadDialog Then
            g_UserPreferences.SetPalettePath sFile
            dstFilename = sFile
        Else
            dstFilename = vbNullString
        End If
        
    End If
    
    'Re-enable user input
    Interface.EnableUserInput
    
End Function
