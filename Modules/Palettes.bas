Attribute VB_Name = "Palettes"
'***************************************************************************
'PhotoDemon's Central Palette Interface
'Copyright 2017-2026 by Tanner Helland
'Created: 12/January/17
'Last updated: 08/March/22
'Last update: use new pdHistogramHash class to greatly accelerate median cut palette generation
'
'This module contains a bunch of helper algorithms for generating optimal palettes from arbitrary
' source images, and also applying arbitrary palettes to images.
'
'Please note that this module has quite a few dependencies.  In particular, it performs no quantization
' (and relatively little palette-matching) on its own.  This is primarily delegated to helper classes.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Do *not* change the order of this enum unless you also change the order of common dialog filters in the
' Load/Save palette functions.  Those indices need to match 1:1 to this enum.
Public Enum PD_PaletteFormat
    pdpf_AdobeColorSwatch = 0
    pdpf_AdobeColorTable = 1
    pdpf_AdobeSwatchExchange = 2
    pdpf_GIMP = 3
    pdpf_PaintDotNet = 4
    pdpf_PSP = 5
    pdpf_PhotoDemon = 6
End Enum

#If False Then
    Private Const pdpf_AdobeColorSwatch = 0, pdpf_AdobeColorTable = 1, pdpf_AdobeSwatchExchange = 2, pdpf_GIMP = 3, pdpf_PaintDotNet = 4, pdpf_PSP = 5, pdpf_PhotoDemon = 6
#End If

Public Enum PD_StockPalette
    pdsp_EGA = 0
    pdsp_PSLegacy = 1
End Enum

#If False Then
    Private Const pdsp_EGA = 0, pdsp_PSLegacy = 1
#End If

'Used for more accurate color distance comparisons (using human eye sensitivity as a rough guide, while staying in
' the sRGB space for performance reasons)
Private Const CUSTOM_WEIGHT_RED As Single = 0.299!
Private Const CUSTOM_WEIGHT_GREEN As Single = 0.587!
Private Const CUSTOM_WEIGHT_BLUE As Single = 0.114!

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

'When interacting with a pdPalette class instance, additional options are available.  In particular,
' pdPalette-based palettes support the following per-color features:
' - RGBA color descriptors (including alpha, although in many places PD does *not* guarantee that
'   alpha values for a given palette entry will be respected during matching).
' - Color name.  Some palette formats provide per-color names; some do not.  This value may be null.
Public Type PDPaletteEntry
    ColorValue As RGBQuad
    ColorName As String
End Type

Public Type PDPaletteCache
    ColorValue As RGBQuad
    OrigIndex As Long
End Type

'Return the number of unique colors in a given DIB.  Often helpful for making subsequent palette decisions.
Public Function GetDIBColorCount(ByRef srcDIB As pdDIB, Optional ByVal includeAlpha As Boolean = True) As Long

    'Create a local array and point it at the pixel data we want to operate on
    Dim imageData() As Byte, tmpSA As SafeArray1D
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = srcDIB.GetDIBStride - 1
    finalY = srcDIB.GetDIBHeight - 1
    
    Dim r As Long, g As Long, b As Long, a As Long
    
    'A special color counting class is used to count unique RGB and RGBA values
    Dim cTree As pdColorCount
    Set cTree = New pdColorCount
    cTree.SetAlphaTracking includeAlpha
    
    'Iterate through all pixels, counting unique values as we go.
    For y = initY To finalY
        srcDIB.WrapArrayAroundScanline imageData, tmpSA, y
    For x = initX To finalX Step 4
    
        b = imageData(x)
        g = imageData(x + 1)
        r = imageData(x + 2)
        a = imageData(x + 3)
        cTree.AddColor r, g, b, a
        
    Next x
    Next y
    
    'Safely deallocate imageData()
    srcDIB.UnwrapArrayFromDIB imageData
    
    If includeAlpha Then
        GetDIBColorCount = cTree.GetUniqueRGBACount()
    Else
        GetDIBColorCount = cTree.GetUniqueRGBCount()
    End If
    
End Function

'Return the number of unique colors in a DIB, *UP TO 256 RGBA QUADS*.  Once 257 unique quads are identified,
' the function immediately aborts.  This function is designed primarily for image export to 8-bit formats,
' so you can quickly determine if you need to manually palettize a given image or save it as-is.
'
'Note that - by design - this function includes alpha in its calculations (e.g. [0, 0, 0, 0] and [0, 0, 0, 255]
' will be treated as *different* entities).  If you don't want this behavior, composite against a backcolor
' prior to calling this function.
Public Function GetDIBColorCount_FastAbort(ByRef srcDIB As pdDIB, ByRef dstPalette() As RGBQuad) As Long

    'Create a local array and point it at the pixel data we want to operate on
    Dim imageData() As Byte, tmpSA As SafeArray1D
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = srcDIB.GetDIBStride - 1
    finalY = srcDIB.GetDIBHeight - 1
    
    Dim r As Long, g As Long, b As Long, a As Long, numColors As Long
    numColors = 0
    
    'A special color counting class is used to count unique RGB and RGBA values
    Dim cTree As pdColorCount
    Set cTree = New pdColorCount
    cTree.SetAlphaTracking True
    
    'Iterate through all pixels, counting unique values as we go.
    For y = initY To finalY
        srcDIB.WrapArrayAroundScanline imageData, tmpSA, y
    For x = initX To finalX Step 4
    
        b = imageData(x)
        g = imageData(x + 1)
        r = imageData(x + 2)
        a = imageData(x + 3)
        If cTree.AddColor(r, g, b, a) Then
            numColors = numColors + 1
            If (numColors > 256) Then
                GetDIBColorCount_FastAbort = 257
                GoTo ExitImmediately
            End If
        End If
        
    Next x
    Next y
    
ExitImmediately:
    
    'Safely deallocate imageData(), then return the final color count
    srcDIB.UnwrapArrayFromDIB imageData
    GetDIBColorCount_FastAbort = cTree.GetUniqueRGBACount()
    
    'If the final color count is <= 256, the palette will most likely be used as-is.
    ' Return it, so the caller can skip an additional palette generation call after this.
    If (numColors <= 256) Then cTree.GetPalette dstPalette
    
End Function

'Does an arbitrary parent palette contain all colors in an arbitrary child palette?  This is used when exporting
' animated GIF files, as a global palette that already contains all colors in a child palette can be used in place
' of a (redundant) local palette.

'Returns: TRUE if parentPalette() contains all colors in childPalette().
Public Function DoesPaletteContainPalette(ByRef parentPalette() As RGBQuad, ByVal numColorsInParent As Long, ByRef srcChildPalette() As RGBQuad, ByVal numColorsInChild As Long) As Boolean
    
    'The inner test will set this to FALSE if/when a missing color is found
    DoesPaletteContainPalette = True
    
    Dim i As Long, j As Long
    Dim matchFound As Boolean
    Dim chkColor1 As Long, chkColor2 As Long
    
    For i = 0 To numColorsInChild - 1
        
        matchFound = False
        
        For j = 0 To numColorsInParent - 1
            GetMem4 VarPtr(parentPalette(j)), chkColor1
            GetMem4 VarPtr(srcChildPalette(i)), chkColor2
            If (chkColor1 = chkColor2) Then
                matchFound = True
                Exit For
            End If
        Next j
        
        If (Not matchFound) Then
            DoesPaletteContainPalette = False
            Exit For
        End If
        
    Next i
    
End Function

'Ensure alpha values in a palette are limited to 0 or 255, nothing in-between.  (This is useful
' when exporting GIFs, for example.)
Public Sub EnsureBinaryAlphaPalette(ByRef srcPalette() As RGBQuad)
    Dim idxEntry As Long
    For idxEntry = LBound(srcPalette) To UBound(srcPalette)
        If (srcPalette(idxEntry).Alpha < 127) Then
            srcPalette(idxEntry).Alpha = 0
        Else
            srcPalette(idxEntry).Alpha = 255
        End If
    Next idxEntry
End Sub

'Given a palette, make sure black and white exist.  This function scans the palette and replaces the darkest
' entry with black, and the brightest entry with white.  (We use this approach so that we can accept palettes
' from any source, even ones that have already contain 256+ entries.)  No changes are made to palettes that
' already contain black and white.
Public Function EnsureBlackAndWhiteInPalette(ByRef srcPalette() As RGBQuad, Optional ByRef srcDIB As pdDIB = Nothing, Optional ByVal mustHaveBlack As Boolean = True, Optional ByVal mustHaveWhite As Boolean = True) As Boolean
    
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
        
        If (preserveBlack And mustHaveBlack) Then
            With srcPalette(minLuminanceIndex)
                .Red = 0
                .Green = 0
                .Blue = 0
            End With
        End If
        
        If (preserveWhite And mustHaveWhite) Then
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

'Given a source palette and an arbitrary RGB value, return the best-matching palette index.
' This function is intended for one-off use only; for best performance, you should integrate
' pdKDTree directly into your function.
Public Function GetNearestIndexRGB(ByRef srcPalette() As RGBQuad, ByVal srcColor As Long, Optional ByVal numOfColors As Long = -1) As Long

    If (numOfColors <= 0) Then numOfColors = UBound(srcPalette) + 1
    
    Dim cTree As pdKDTree
    Set cTree = New pdKDTree
    cTree.BuildTree srcPalette, numOfColors
    
    Dim tmpColor As RGBQuad
    tmpColor.Red = Colors.ExtractRed(srcColor)
    tmpColor.Green = Colors.ExtractGreen(srcColor)
    tmpColor.Blue = Colors.ExtractBlue(srcColor)
    tmpColor.Alpha = 255
    GetNearestIndexRGB = cTree.GetNearestPaletteIndex(tmpColor)
    
End Function

'Given a source image, an (empty) destination palette array, and a color count, return an optimized palette using
' the source image as the reference.  A modified median-cut system is used, and it achieves a very nice
' combination of performance, low memory usage, and high-quality output.
'
'Because palette generation is a time-consuming task, the source DIB should generally be shrunk to a much smaller
' version of itself.  I built a function specifically for this: DIBs.ResizeDIBByPixelCount().  That function
' resizes an image to a target pixel count, and I wouldn't recommend a net size any larger than ~500,000 pixels.
Public Function GetNeuquantPalette_RGBA(ByRef srcDIB As pdDIB, ByRef dstPalette() As RGBQuad, Optional ByVal numOfColors As Long = 256, Optional ByVal suppressMessages As Boolean = True, Optional ByVal modifyProgBarMax As Long = -1, Optional ByVal modifyProgBarOffset As Long = 0) As Boolean
    
    'Do not request less than two colors in the final palette!
    If (numOfColors < 2) Then numOfColors = 2
    
    'Neuquant has some quirks compared to other palette-generation algorithms.  One quirk is that it
    ' works very very well on smaller pixel counts, because too many pixels counts just contribute
    ' noise (which the algorithm will successfully filter out... albeit very slowly).
    '
    'As such, image sizes above like a megapixel do not produce better results.  If the source image
    ' is very large, downsample it.
    Const ARBITRARY_MAX_PIXELS As Long = 1500000
    Const ARBITRARY_MIN_PIXELS As Long = 10000
    
    Dim numPixelsSrc As Long
    numPixelsSrc = srcDIB.GetDIBWidth * srcDIB.GetDIBHeight
    
    Dim actualDIBToSample As pdDIB, samplingQuality As Long
    If (numPixelsSrc > ARBITRARY_MAX_PIXELS) Then
        
        'Downsample the image, and instruct neuquant to only sample half the image (as noise may still
        ' be present)
        DIBs.ResizeDIBByPixelCount srcDIB, actualDIBToSample, ARBITRARY_MAX_PIXELS, GP_IM_HighQualityBilinear
        samplingQuality = 1
        
    'Similarly, if the image is tiny, neuquant won't have enough time+data to converge on a great palette.
    ElseIf (numPixelsSrc < ARBITRARY_MIN_PIXELS) Then
        
        'Upsample the image, and instruct neuquant to sample every pixel (to avoid unwanted weighting
        ' against a particular region of the upsampled image)
        DIBs.ResizeDIBByPixelCount srcDIB, actualDIBToSample, ARBITRARY_MIN_PIXELS, GP_IM_NearestNeighbor, True
        samplingQuality = 1
    
    'For other sizes, we can simple sample the source image as-is
    Else
        Set actualDIBToSample = srcDIB
        samplingQuality = 1
    End If
    
    'On previews, cut sampling quality further to improve responsiveness
    If suppressMessages Then samplingQuality = samplingQuality + 2
    
    'Instantiate a neural network class and notify it of the desired color count.
    Dim cNeuquant As pdNeuquant
    Set cNeuquant = New pdNeuquant
    cNeuquant.SetColorCount numOfColors
    
    'Initialize the network against the source image, and pass the sampling quality factor
    ' (1 = perfect sampling, 30 = 1/30th of pixels in image sampled).  The initialization function
    ' will return the net number of pixels the function expects to sample based on the input settings.
    Dim maxProgress As Long
    maxProgress = cNeuquant.InitializeNeuralNetwork(actualDIBToSample, samplingQuality)
    
    'Determine progress bar increments (and note that these can be modified by the caller, if this function
    ' is called as part of a broader operation)
    Dim progBarCheck As Long
    If (Not suppressMessages) Then
        If (modifyProgBarMax < 0) Then ProgressBars.SetProgBarMax maxProgress Else ProgressBars.SetProgBarMax modifyProgBarMax
        progBarCheck = ProgressBars.FindBestProgBarValue()
    End If
    
    'Train the network against the image
    cNeuquant.TrainNeuralNetwork suppressMessages, modifyProgBarMax, modifyProgBarOffset
    
    'Temporary image is no longer required; free it immediately
    Set actualDIBToSample = Nothing
    
    'Retrieve the final palette
    cNeuquant.GetFinalPalette dstPalette
    GetNeuquantPalette_RGBA = True
    
    'If the palette retrieval process was successful, sort the palette from "least alpha" to "most alpha";
    ' this typically produces a slightly smaller final file size (as PNG, for example, only requires you to
    ' write non-255 values to its transparency segment; any unspecified values are assumed to be opaque).
    ' (Note that this step is not required; a "default order" palette is perfectly valid.)
    Dim cPaletteSorter As pdPaletteChild
    Set cPaletteSorter = New pdPaletteChild
    cPaletteSorter.CreateFromRGBQuads dstPalette
    cPaletteSorter.SortByChannel 3                  '0 = red, 1 = green, 2 = blue, 3 = alpha
    cPaletteSorter.CopyRGBQuadsToArray dstPalette
    
End Function

'Given a source image, an (empty) destination palette array, and a color count, return an optimized palette using
' the source image as the reference.  A modified median-cut system is used, and it achieves a very nice
' combination of performance, low memory usage, and high-quality output.
Public Function GetOptimizedPalette(ByRef srcDIB As pdDIB, ByRef dstPalette() As RGBQuad, Optional ByVal numOfColors As Long = 256, Optional ByVal quantMode As PD_QuantizeMode = pdqs_Variance, Optional ByVal suppressMessages As Boolean = True, Optional ByVal modifyProgBarMax As Long = -1, Optional ByVal modifyProgBarOffset As Long = 0) As Boolean
    
    'Do not request less than two colors in the final palette!
    If (numOfColors < 2) Then numOfColors = 2
    
    Dim pxSize As Long
    pxSize = srcDIB.GetDIBColorDepth \ 8
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = srcDIB.GetDIBWidth - 1
    finalY = srcDIB.GetDIBHeight - 1
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates a
    ' refresh interval based on the size of the area to be processed.
    Dim progBarCheck As Long
    If (Not suppressMessages) Then
        If (modifyProgBarMax = -1) Then SetProgBarMax finalY Else SetProgBarMax modifyProgBarMax
        progBarCheck = ProgressBars.FindBestProgBarValue()
    End If
    
    'Add all pixels from the source image to a base color stack
    Dim pxStack() As pdMedianCut
    ReDim pxStack(0 To numOfColors - 1) As pdMedianCut
    Set pxStack(0) = New pdMedianCut
    
    'Note that PD actually supports quite a few different quantization methods.  At present, we use a technique
    ' that's a good compromise between performance and quality.
    pxStack(0).SetQuantizeMode quantMode
    
    'To improve performance further, we start by assembling an RGBA histogram of each color in the image.
    ' Most photos contain only a small subset of colors (typical color count is < 10k colors per megapixel,
    ' so < 200k colors for a 20mp photo), so we are likely to see many repeat color entries.  By merging
    ' repeat color instances into a single "color + occurrence count" value, we greatly trim the size of
    ' the final color tree, which makes pruning it *significantly* faster.  For example, on a 20-megapixel
    ' RGBA image with variable transparency, this technique cuts running time of this function by ~60-70%.
    Dim pxHist As pdHistogramHash
    Set pxHist = New pdHistogramHash
    
    'Add all colors to the histogram
    Dim srcPixels() As Long, tmpSA As SafeArray1D
    
    For y = 0 To finalY
        srcDIB.WrapLongArrayAroundScanline srcPixels, tmpSA, y
    For x = 0 To finalX
        pxHist.AddColor srcPixels(x) Or &HFF000000
    Next x
        If (Not suppressMessages) Then
            If (y And progBarCheck) = 0 Then
                If Interface.UserPressedESC() Then Exit For
                SetProgBarVal y + modifyProgBarOffset
            End If
        End If
    Next y
    
    'Direct pixel access is no longer required; safely free the unsafe pixel wrapper
    srcDIB.UnwrapLongArrayFromDIB srcPixels
    
    'Next, retrieve the list of unique colors (and their counts), and bulk-add all of those colors
    ' to the median cut object.
    Dim listOfPixels() As RGBQuad, pixelCounts() As Long, numPixels As Long
    numPixels = pxHist.GetUniqueColors(listOfPixels, pixelCounts)
    Set pxHist = Nothing
    pxStack(0).BulkAddColors_RGBA listOfPixels, pixelCounts, numPixels
    
    'Next, make sure there are more than [numOfColors] colors in the image (otherwise, our work is already done!)
    If (pxStack(0).GetNumOfColors > numOfColors) Then
        
        Dim stackCount As Long
        stackCount = 1
        
        Dim maxVariance As Single, mvIndex As Long
        Dim i As Long
        
        Dim rVariance As Single, gVariance As Single, bVariance As Single, netVariance As Single
        
        'With the initial stack constructed, we can now start partitioning it into smaller stacks based on variance
        Do
        
            'Reset maximum variance (because we need to calculate it anew)
            maxVariance = 0!
            
            'Find the largest total variance in the current stack collection
            For i = 0 To stackCount - 1
            
                pxStack(i).GetVariance rVariance, gVariance, bVariance
                
                netVariance = rVariance + gVariance + bVariance
                If (netVariance > maxVariance) Then
                    mvIndex = i
                    maxVariance = netVariance
                End If
                
            Next i
            
            'Ask the stack with the largest net variance to split itself in half.  (Note that the stack object
            ' itself decides which axis is most appropriate for splitting; typically this is the axis -
            ' e.g. channel - with the largest variance.)
            'Debug.Print "Largest variance was " & maxVariance & ", found in stack #" & mvIndex & " (total stack count is " & stackCount & ")"
            If (maxVariance > 0!) Then
                pxStack(mvIndex).Split pxStack(stackCount)
                stackCount = stackCount + 1
            
            'All current stacks only contain a single color, meaning this image contains fewer unique colors
            ' than the target number of colors the user requested.  That's okay!  Exit now, and use the colors
            ' we've discovered as the optimal palette.
            Else
                numOfColors = stackCount
                Exit Do
            End If
        
        'Continue splitting stacks until we arrive at the desired number of colors.  (Each stack represents
        ' one color in the final palette.)
        Loop While (stackCount < numOfColors)
        
        'We now have [numOfColors] unique color stacks.  Each of these represents a set of similar colors.
        ' Generate a final palette by requesting the weighted average of each stack.  (As an alternate solution,
        ' you could also request the most "populous" color; this would preserve precise tones from the image,
        ' but rarely-appearing colors would never influence the final output.  Trade-offs!
        Dim newR As Long, newG As Long, newB As Long
        ReDim dstPalette(0 To numOfColors - 1) As RGBQuad
        For i = 0 To numOfColors - 1
            pxStack(i).GetAverageColor newR, newG, newB
            dstPalette(i).Red = newR
            dstPalette(i).Green = newG
            dstPalette(i).Blue = newB
            dstPalette(i).Alpha = 255
        Next i
        
        GetOptimizedPalette = True
        
    'If there are less than [numOfColors] unique colors in the image, simply copy the existing stack into a palette
    Else
        pxStack(0).CopyStackToRGBQuad dstPalette
        For i = 0 To UBound(dstPalette)
            dstPalette(i).Alpha = 255
        Next i
        GetOptimizedPalette = True
    End If
    
End Function

'Given a source image, an (empty) destination palette array, and a color count, return an optimized palette using
' the source image as the reference.  A modified median-cut system is used, and it achieves a very nice
' combination of performance, low memory usage, and high-quality output.
'
'Because palette generation is a time-consuming task, you can save some time by shrinking the source image
' prior to quantizing.  I built a function specifically for this: DIBs.ResizeDIBByPixelCount().  That function
' resizes an image to a target pixel count, and generally speaking, resizing down to ~500,000 pixels rarely
' affects image quality (but can greatly increase quantization perf).
Public Function GetOptimizedPaletteIncAlpha(ByRef srcDIB As pdDIB, ByRef dstPalette() As RGBQuad, Optional ByVal numOfColors As Long = 256, Optional ByVal quantMode As PD_QuantizeMode = pdqs_Variance, Optional ByVal suppressMessages As Boolean = True, Optional ByVal modifyProgBarMax As Long = -1, Optional ByVal modifyProgBarOffset As Long = 0) As Boolean
    
    'Do not request less than two colors in the final palette!
    If (numOfColors < 2) Then numOfColors = 2
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = srcDIB.GetDIBWidth - 1
    finalY = srcDIB.GetDIBHeight - 1
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates a
    ' refresh interval based on the size of the area to be processed.
    Dim progBarCheck As Long
    If (Not suppressMessages) Then
        If (modifyProgBarMax = -1) Then SetProgBarMax finalY Else SetProgBarMax modifyProgBarMax
        progBarCheck = ProgressBars.FindBestProgBarValue()
    End If
    
    'Add all pixels from the source image to a base color stack
    Dim pxStack() As pdMedianCut
    ReDim pxStack(0 To numOfColors - 1) As pdMedianCut
    Set pxStack(0) = New pdMedianCut
    
    'Note that PD actually supports quite a few different quantization methods.  At present, we use a technique
    ' that's a good compromise between performance and quality.
    pxStack(0).SetQuantizeMode quantMode
    
    'To improve performance further, we start by assembling an RGBA histogram of each color in the image.
    ' Most photos contain only a small subset of colors (typical color count is < 10k colors per megapixel,
    ' so < 200k colors for a 20mp photo), so we are likely to see many repeat color entries.  By merging
    ' repeat color instances into a single "color + occurrence count" value, we greatly trim the size of
    ' the final color tree, which makes pruning it *significantly* faster.  For example, on a 20-megapixel
    ' RGBA image with variable transparency, this technique cuts running time of this function by ~60-70%.
    Dim pxHist As pdHistogramHash
    Set pxHist = New pdHistogramHash
    
    'Add all colors to the histogram
    Dim srcPixels() As Long, tmpSA As SafeArray1D
    
    For y = 0 To finalY
        srcDIB.WrapLongArrayAroundScanline srcPixels, tmpSA, y
    For x = 0 To finalX
        pxHist.AddColor srcPixels(x)
    Next x
        If (Not suppressMessages) Then
            If (y And progBarCheck) = 0 Then
                If Interface.UserPressedESC() Then Exit For
                SetProgBarVal y + modifyProgBarOffset
            End If
        End If
    Next y
    
    'Direct pixel access is no longer required; safely free the unsafe pixel wrapper
    srcDIB.UnwrapLongArrayFromDIB srcPixels
    
    'Next, retrieve the list of unique colors (and their counts), and bulk-add all of those colors
    ' to the median cut object.
    Dim listOfPixels() As RGBQuad, pixelCounts() As Long, numPixels As Long
    numPixels = pxHist.GetUniqueColors(listOfPixels, pixelCounts)
    Set pxHist = Nothing
    pxStack(0).BulkAddColors_RGBA listOfPixels, pixelCounts, numPixels
    
    'Next, make sure there are more than [numOfColors] colors in the image (otherwise, our work is already done!)
    If (pxStack(0).GetNumOfColors > numOfColors) Then
        
        Dim stackCount As Long
        stackCount = 1
        
        Dim maxVariance As Single, mvIndex As Long
        Dim i As Long
        
        Dim rVariance As Single, gVariance As Single, bVariance As Single, aVariance As Single, netVariance As Single
        
        'With the initial stack constructed, we can now start partitioning it into smaller stacks based on variance
        Do
        
            'Reset maximum variance (because we need to calculate it anew)
            maxVariance = 0!
            
            'Find the largest total variance in the current stack collection
            For i = 0 To stackCount - 1
            
                pxStack(i).GetVariance_Alpha rVariance, gVariance, bVariance, aVariance
                
                netVariance = rVariance + gVariance + bVariance + aVariance
                If (netVariance > maxVariance) Then
                    mvIndex = i
                    maxVariance = netVariance
                End If
                
            Next i
            
            'Ask the stack with the largest net variance to split itself in half.  (Note that the stack object
            ' itself decides which axis is most appropriate for splitting; typically this is the axis - channel -
            ' with the largest variance.)
            'Debug.Print "Largest variance was " & maxVariance & ", found in stack #" & mvIndex & " (total stack count is " & stackCount & ")"
            If (maxVariance > 0!) Then
                pxStack(mvIndex).SplitIncludingAlpha pxStack(stackCount)
                stackCount = stackCount + 1
            
            'All current stacks only contain a single color, meaning this image contains fewer unique colors
            ' than the target number of colors the user requested.  That's okay!  Exit now, and use the colors
            ' we've discovered as the optimal palette.
            Else
                numOfColors = stackCount
                Exit Do
            End If
        
        'Continue splitting stacks until we arrive at the desired number of colors.  (Each stack represents
        ' one color in the final palette.)
        Loop While (stackCount < numOfColors)
        
        'We now have [numOfColors] unique color stacks.  Each of these represents a set of similar colors.
        ' Generate a final palette by requesting the weighted average of each stack.  (As an alternate solution,
        ' you could also request the most "populous" color; this would preserve precise tones from the image,
        ' but rarely-appearing colors would never influence the final output.  Trade-offs!
        Dim newR As Long, newG As Long, newB As Long, newA As Long
        ReDim dstPalette(0 To numOfColors - 1) As RGBQuad
        For i = 0 To numOfColors - 1
            pxStack(i).GetAverageColorAndAlpha newR, newG, newB, newA
            dstPalette(i).Red = newR
            dstPalette(i).Green = newG
            dstPalette(i).Blue = newB
            dstPalette(i).Alpha = newA
        Next i
        
        GetOptimizedPaletteIncAlpha = True
        
    'If there are less than [numOfColors] unique colors in the image, simply copy the existing stack into a palette
    Else
        pxStack(0).CopyStackToRGBQuad dstPalette, True
        GetOptimizedPaletteIncAlpha = True
    End If
    
    'If the palette retrieval process was successful, sort the palette from "least alpha" to "most alpha";
    ' this typically produces a slightly smaller final file size (as PNG, for example, only requires you to
    ' write non-255 values to its transparency segment; any unspecified values are assumed to be opaque).
    ' (Note that this step is not required; a "default order" palette is perfectly valid.)
    Dim cPaletteSorter As pdPaletteChild
    Set cPaletteSorter = New pdPaletteChild
    cPaletteSorter.CreateFromRGBQuads dstPalette
    cPaletteSorter.SortByChannel 3                  '0 = red, 1 = green, 2 = blue, 3 = alpha
    
    'Failsafe; it's fast and easy to remove duplicates, just in case weirdness happened during quantization
    cPaletteSorter.FindAndRemoveDuplicates
    cPaletteSorter.CopyRGBQuadsToArray dstPalette
    
End Function

'Given a source pdImage, an (empty) destination palette array, and a color count, return an optimized palette
' using *ALL* source layers from the source pdImage.  A modified median-cut system is used, and it achieves
' a great combination of performance, low memory usage, and high-quality output.
'
'Because palette generation is a time-consuming task, you might consider implementing some element of
' random sampling here (or alternatively, reducing each layer to some subset of its original size).
' The hard thing with this is that there's not exactly a ton of research on the ideal values to use here,
' so any existing code is based on my best guesses rather than an empirical study.
Public Function GetOptimizedPaletteIncAlpha_AllLayers(ByRef srcImage As pdImage, ByRef dstPalette() As RGBQuad, Optional ByVal numOfColors As Long = 256, Optional ByVal quantMode As PD_QuantizeMode = pdqs_Variance, Optional ByVal suppressMessages As Boolean = True, Optional ByVal modifyProgBarMax As Long = -1, Optional ByVal modifyProgBarOffset As Long = 0) As Boolean
    
    'Failsafe checks
    If (srcImage Is Nothing) Then Exit Function
    
    'Do not request less than two colors in the final palette!
    If (numOfColors < 2) Then numOfColors = 2
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates a
    ' refresh interval based on the size of the area to be processed.
    Dim progBarCheck As Long
    If (Not suppressMessages) Then
        If (modifyProgBarMax = -1) Then SetProgBarMax srcImage.GetNumOfLayers Else SetProgBarMax modifyProgBarMax
        progBarCheck = ProgressBars.FindBestProgBarValue()
    End If
    
    Dim tmpDIB As pdDIB
    
    Const MAX_PIXELS_PER_LAYER As Long = 50000
    
    'PD has several different tools for generating palettes.  A median-cut approach is fast, even for this
    ' multi-layer approach.
    Dim pxStack() As pdMedianCut
    ReDim pxStack(0 To numOfColors - 1) As pdMedianCut
    Set pxStack(0) = New pdMedianCut
    
    'Note that PD actually supports quite a few different quantization methods.  At present, we use a technique
    ' that's a good compromise between performance and quality.
    pxStack(0).SetQuantizeMode quantMode
    
    'To improve performance further, we start by assembling an RGBA histogram of each color in the image.
    ' Most photos contain only a small subset of colors (typical color count is < 10k colors per megapixel,
    ' so < 200k colors for a 20mp photo), so we are likely to see many repeat color entries.  By merging
    ' repeat color instances into a single "color + occurrence count" value, we greatly trim the size of
    ' the final color tree, which makes pruning it *significantly* faster.  For example, on a 20-megapixel
    ' RGBA image with variable transparency, this technique cuts running time of this function by ~60-70%.
    Dim pxHist As pdHistogramHash
    Set pxHist = New pdHistogramHash
    
    'Add all colors to the histogram
    Dim srcPixels() As Long, tmpSA As SafeArray1D
    
    'Add all pixels from every layer to a base color stack
    Dim i As Long
    For i = 0 To srcImage.GetNumOfLayers - 1
    
        'Limit the size of the source layer to MAX_PIXELS_PER_LAYER
        If DIBs.ResizeDIBByPixelCount(srcImage.GetLayerByIndex(i).GetLayerDIB, tmpDIB, MAX_PIXELS_PER_LAYER) Then
        
            'Wrap an array around the temporary DIB copy
            tmpDIB.WrapLongArrayAroundDIB_1D srcPixels, tmpSA
            
            Dim numPixels As Long
            numPixels = (tmpDIB.GetDIBWidth * tmpDIB.GetDIBHeight) - 1
            
            'Load all pixels into the median cut object
            Dim x As Long
            For x = 0 To numPixels
                pxHist.AddColor srcPixels(x)
            Next x
            
            'Ensure we safely unwrap the array wrapper from each layer
            tmpDIB.UnwrapLongArrayFromDIB srcPixels
            
        End If
        
        'Update progress bar on each completed layer
        If (Not suppressMessages) Then
            If (i And progBarCheck) = 0 Then
                If Interface.UserPressedESC() Then Exit For
                ProgressBars.SetProgBarVal i + modifyProgBarOffset
            End If
        End If
        
    Next i
    
    'Next, retrieve the list of unique colors (and their counts), and bulk-add all of those colors
    ' to the median cut object.
    Dim listOfPixels() As RGBQuad, pixelCounts() As Long
    numPixels = pxHist.GetUniqueColors(listOfPixels, pixelCounts)
    Set pxHist = Nothing
    pxStack(0).BulkAddColors_RGBA listOfPixels, pixelCounts, numPixels
    
    'Next, make sure there are more than [numOfColors] colors in the image (otherwise, our work is already done!)
    If (pxStack(0).GetNumOfColors > numOfColors) Then
        
        Dim stackCount As Long
        stackCount = 1
        
        Dim maxVariance As Single, mvIndex As Long
        Dim rVariance As Single, gVariance As Single, bVariance As Single, aVariance As Single, netVariance As Single
        Dim r As Long, g As Long, b As Long, a As Long
        
        'With the initial stack constructed, we can now start partitioning it into smaller stacks based on variance
        Do
        
            'Reset maximum variance (because we need to calculate it anew)
            maxVariance = 0!
            
            'Find the largest total variance in the current stack collection
            For i = 0 To stackCount - 1
            
                pxStack(i).GetVariance_Alpha rVariance, gVariance, bVariance, aVariance
                
                netVariance = rVariance + gVariance + bVariance + aVariance
                If (netVariance > maxVariance) Then
                    mvIndex = i
                    maxVariance = netVariance
                End If
                
            Next i
            
            'Ask the stack with the largest net variance to split itself in half.  (Note that the stack object
            ' itself decides which axis is most appropriate for splitting; typically this is the axis - channel -
            ' with the largest variance.)
            'Debug.Print "Largest variance was " & maxVariance & ", found in stack #" & mvIndex & " (total stack count is " & stackCount & ")"
            If (maxVariance > 0!) Then
                pxStack(mvIndex).SplitIncludingAlpha pxStack(stackCount)
                stackCount = stackCount + 1
            
            'All current stacks only contain a single color, meaning this image contains fewer unique colors
            ' than the target number of colors the user requested.  That's okay!  Exit now, and use the colors
            ' we've discovered as the optimal palette.
            Else
                numOfColors = stackCount
                Exit Do
            End If
            
            'If we're still splitting colors, and we're at one less than the target color count,
            ' stop and look for a fully transparent color in the palette.  We need this for pixel-blanking
            ' when optimizing animated images, and if we don't have one yet, it's easiest to just manually
            ' add one now.
            If (stackCount = numOfColors - 1) Then
                
                Dim trnsFound As Boolean: trnsFound = False
                
                For x = 0 To stackCount - 1
                    pxStack(i).GetAverageColorAndAlpha r, g, b, a
                    If (r = 0) And (g = 0) And (b = 0) And (a = 0) Then
                        trnsFound = True
                        Exit For
                    End If
                Next x
                
                'If we didn't find transparency, add it now.
                If (Not trnsFound) Then
                    Set pxStack(stackCount) = New pdMedianCut
                    pxStack(stackCount).AddColor_RGBA 0, 0, 0, 0
                    stackCount = stackCount + 1
                End If
                
            End If
            
        'Continue splitting stacks until we arrive at the desired number of colors.  (Each stack represents
        ' one color in the final palette.)
        Loop While (stackCount < numOfColors)
        
        'We now have [numOfColors] unique color stacks.  Each of these represents a set of similar colors.
        ' Generate a final palette by requesting the weighted average of each stack.  (As an alternate solution,
        ' you could also request the most "populous" color; this would preserve precise tones from the image,
        ' but rarely-appearing colors would never influence the final output.  Trade-offs!
        Dim newR As Long, newG As Long, newB As Long, newA As Long
        ReDim dstPalette(0 To numOfColors - 1) As RGBQuad
        For i = 0 To numOfColors - 1
            pxStack(i).GetAverageColorAndAlpha newR, newG, newB, newA
            dstPalette(i).Red = newR
            dstPalette(i).Green = newG
            dstPalette(i).Blue = newB
            dstPalette(i).Alpha = newA
        Next i
        
        GetOptimizedPaletteIncAlpha_AllLayers = True
        
    'If there are less than [numOfColors] unique colors in the image, simply copy the existing stack into a palette
    Else
        pxStack(0).CopyStackToRGBQuad dstPalette, True
        GetOptimizedPaletteIncAlpha_AllLayers = True
    End If
    
    'If the palette retrieval process was successful, sort the palette from "least alpha" to "most alpha";
    ' this typically produces a slightly smaller final file size (as PNG, for example, only requires you to
    ' write non-255 values to its transparency segment; any unspecified values are assumed to be opaque).
    ' (Note that this step is not required; a "default order" palette is perfectly valid.)
    Dim cPaletteSorter As pdPaletteChild
    Set cPaletteSorter = New pdPaletteChild
    cPaletteSorter.CreateFromRGBQuads dstPalette
    cPaletteSorter.SortByChannel 3                  '0 = red, 1 = green, 2 = blue, 3 = alpha
    
    'Failsafe; it's fast and easy to remove duplicates, just in case weirdness happened during quantization.
    ' (This is really only needed if there are less than [requested numOfColors] colors in the image, because
    ' the "split" step never go ta chance to merge duplicate pixels.)
    cPaletteSorter.FindAndRemoveDuplicates
    cPaletteSorter.CopyRGBQuadsToArray dstPalette
    
End Function

'Given a source image, an (empty) destination palette array, and a color count, return an optimized palette using
' the source image as the reference.  Analysis is done in LAB color space, for slower but potentially better results.
' (That said, internal testing shows that this produces results that don't "look as good" as traditional methods.
' I'm not sure why - it's possibly caused, in part, by not weighting L more aggressively than A/B, but I can't easily
' rectify that without major changes to the underlying median cut engine.)  As such, PD doesn't use this function
' at present, but it may be useful in the future if I have more time to refine it.
Public Function GetOptimizedPaletteIncAlpha_LAB(ByRef srcDIB As pdDIB, ByRef dstPalette() As RGBQuad, Optional ByVal numOfColors As Long = 256, Optional ByVal quantMode As PD_QuantizeMode = pdqs_Variance, Optional ByVal suppressMessages As Boolean = True, Optional ByVal modifyProgBarMax As Long = -1, Optional ByVal modifyProgBarOffset As Long = 0) As Boolean
    
    'Do not request less than two colors in the final palette!
    If (numOfColors < 2) Then numOfColors = 2
    
    Dim srcPixels() As Byte, srcSA As SafeArray1D
    Dim srcPixelsLab() As Byte
    
    'Resize the LAB array to the same size as a scanline of the source image
    ReDim srcPixelsLab(0 To srcDIB.GetDIBStride - 1) As Byte
    
    Dim pxSize As Long
    pxSize = srcDIB.GetDIBColorDepth \ 8
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = srcDIB.GetDIBStride - 1
    finalY = srcDIB.GetDIBHeight - 1
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates a
    ' refresh interval based on the size of the area to be processed.
    Dim progBarCheck As Long
    If (Not suppressMessages) Then
        If (modifyProgBarMax = -1) Then SetProgBarMax finalY Else SetProgBarMax modifyProgBarMax
        progBarCheck = ProgressBars.FindBestProgBarValue()
    End If
    
    'Add all pixels from the source image to a base color stack
    Dim pxStack() As pdMedianCut
    ReDim pxStack(0 To numOfColors - 1) As pdMedianCut
    Set pxStack(0) = New pdMedianCut
    
    'Note that PD actually supports quite a few different quantization methods.  At present, we use a technique
    ' that's a good compromise between performance and quality.
    pxStack(0).SetQuantizeMode quantMode
    
    'Use littleCMS to create an RGB -> Lab transform
    Dim cRGB As pdLCMSProfile
    Set cRGB = New pdLCMSProfile
    cRGB.CreateSRGBProfile True
    
    Dim cLAB As pdLCMSProfile
    Set cLAB = New pdLCMSProfile
    cLAB.CreateLabProfile True
    
    Dim cTransform As pdLCMSTransform
    Set cTransform = New pdLCMSTransform
    cTransform.CreateTwoProfileTransform cRGB, cLAB, TYPE_BGRA_8, TYPE_ALab_8, INTENT_RELATIVE_COLORIMETRIC
    
    For y = 0 To finalY
        
        srcDIB.WrapArrayAroundScanline srcPixels, srcSA, y
        
        'Translate the RGB values into LAB
        cTransform.ApplyTransformToScanline VarPtr(srcPixels(0)), VarPtr(srcPixelsLab(0)), srcDIB.GetDIBWidth
            
        For x = 0 To finalX Step pxSize
            
            'Note that we're adding the values in LAB format, and to ensure order is maintained (BGRA vs ALAB),
            ' we manually swap B and R in this function
            pxStack(0).AddColor_RGBA srcPixelsLab(x + 2), srcPixelsLab(x + 1), srcPixelsLab(x), srcPixelsLab(x + 3)
            
        Next x
        
        If (Not suppressMessages) Then
            If (y And progBarCheck) = 0 Then
                If Interface.UserPressedESC() Then Exit For
                SetProgBarVal y + modifyProgBarOffset
            End If
        End If
    Next y
    
    srcDIB.UnwrapArrayFromDIB srcPixels
    
    'Palette generation works the exact same as RGB data; the splitter doesn't care about what format
    ' the colors are in - it'll just sort them in 3D space and split planes optimally.  (In a perfect
    ' world we might weight L over A/B, but that's not handled at present.)
    If (pxStack(0).GetNumOfColors > numOfColors) Then
        
        Dim stackCount As Long
        stackCount = 1
        
        Dim maxVariance As Single, mvIndex As Long
        Dim i As Long
        
        Dim rVariance As Single, gVariance As Single, bVariance As Single, aVariance As Single, netVariance As Single
        
        Do
        
            maxVariance = 0!
            
            For i = 0 To stackCount - 1
                pxStack(i).GetVariance_Alpha rVariance, gVariance, bVariance, aVariance
                netVariance = rVariance + gVariance + bVariance + aVariance
                If (netVariance > maxVariance) Then
                    mvIndex = i
                    maxVariance = netVariance
                End If
            Next i
            
            If (maxVariance > 0!) Then
                pxStack(mvIndex).SplitIncludingAlpha pxStack(stackCount)
                stackCount = stackCount + 1
            Else
                numOfColors = stackCount
                Exit Do
            End If
        
        Loop While (stackCount < numOfColors)
        
        Dim newR As Long, newG As Long, newB As Long, newA As Long
        ReDim dstPalette(0 To numOfColors - 1) As RGBQuad
        For i = 0 To numOfColors - 1
            pxStack(i).GetAverageColorAndAlpha newR, newG, newB, newA
            dstPalette(i).Blue = newB
            dstPalette(i).Green = newG
            dstPalette(i).Red = newR
            dstPalette(i).Alpha = newA
        Next i
        
        GetOptimizedPaletteIncAlpha_LAB = True
        
    Else
        pxStack(0).CopyStackToRGBQuad dstPalette
        GetOptimizedPaletteIncAlpha_LAB = True
    End If
    
    'We now have a finished palette.  Convert it from ALAB back to BGRA color space
    Dim tmpPalette() As RGBQuad
    ReDim tmpPalette(0 To UBound(dstPalette)) As RGBQuad
    
    cTransform.ReleaseTransform
    cTransform.CreateTwoProfileTransform cLAB, cRGB, TYPE_ALab_8, TYPE_BGRA_8, INTENT_RELATIVE_COLORIMETRIC
    cTransform.ApplyTransformToScanline VarPtr(dstPalette(0)), VarPtr(tmpPalette(0)), UBound(dstPalette) + 1
    
    CopyMemoryStrict VarPtr(dstPalette(0)), VarPtr(tmpPalette(0)), (UBound(dstPalette) + 1) * 4
    
End Function

Public Sub GetPalette_Grayscale(ByRef dstPalette() As RGBQuad)
    ReDim dstPalette(0 To 255) As RGBQuad
    Dim i As Long
    For i = 0 To 255
        With dstPalette(i)
            .Red = i
            .Green = i
            .Blue = i
            .Alpha = 255
        End With
    Next i
End Sub

Public Sub GetPalette_GrayscaleEx(ByRef dstPalette() As RGBQuad, ByVal numShades As Long, Optional ByVal dontSizeArray As Boolean = False)
    
    If (numShades > 256) Then numShades = 256
    
    Dim maxVal As Long
    maxVal = numShades - 1
    
    If (Not dontSizeArray) Then ReDim dstPalette(0 To maxVal) As RGBQuad
    
    Dim i As Long, finalGray As Long
    For i = 0 To maxVal
        
        finalGray = Int((i * 255#) / maxVal + 0.5)
        If (finalGray > 255) Then finalGray = 255
        
        With dstPalette(i)
            .Red = finalGray
            .Green = finalGray
            .Blue = finalGray
            .Alpha = 255
        End With
        
    Next i
    
End Sub

'Given a palette with (potentially) one-or-more non-opaque pixels, return only the opaque colors.
Public Function GetPalette_OpaqueColorsOnly(ByRef srcQuads() As RGBQuad) As Long
    Dim i As Long, numOKColors As Long
    For i = 0 To UBound(srcQuads)
        If (srcQuads(i).Alpha = 255) Then
            If (numOKColors < i) Then srcQuads(numOKColors) = srcQuads(i)
            numOKColors = numOKColors + 1
        End If
    Next i
    GetPalette_OpaqueColorsOnly = numOKColors
End Function

'Given an arbitrary source palette, apply said palette to the target image.  Dithering is *not* used.
' Colors are matched exhaustively, meaning this function slows significantly as palette size increases.
Public Function ApplyPaletteToImage_Naive(ByRef dstDIB As pdDIB, ByRef srcPalette() As RGBQuad) As Boolean

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
    
    ApplyPaletteToImage_Naive = True
    
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

'Given an arbitrary palette (including palettes > 256 colors - they work just fine!), apply said palette to the
' target image.  Dithering is *not* used.  Colors are matched using a KD-tree (where the palette is pre-loaded into
' a tree, and colors are matched via that tree).
Public Function ApplyPaletteToImage_KDTree(ByRef dstDIB As pdDIB, ByRef srcPalette() As RGBQuad, Optional ByVal suppressMessages As Boolean = False, Optional ByVal modifyProgBarMax As Long = -1, Optional ByVal modifyProgBarOffset As Long = 0) As Boolean

    Dim srcPixels() As Byte, tmpSA As SafeArray1D
    
    Dim pxSize As Long
    pxSize = dstDIB.GetDIBColorDepth \ 8
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = dstDIB.GetDIBStride - 1
    finalY = dstDIB.GetDIBHeight - 1
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates a
    ' refresh interval based on the size of the area to be processed.
    Dim progBarCheck As Long
    If (Not suppressMessages) Then
        If (modifyProgBarMax = -1) Then SetProgBarMax finalY Else SetProgBarMax modifyProgBarMax
        progBarCheck = ProgressBars.FindBestProgBarValue()
    End If
    
    'As with normal palette matching, we'll use basic RLE acceleration to try and skip palette
    ' searching for contiguous matching colors.
    Dim lastColor As Long: lastColor = -1
    Dim r As Long, g As Long, b As Long
    
    Dim tmpQuad As RGBQuad, newQuad As RGBQuad, lastQuad As RGBQuad
    
    'Build the initial tree
    Dim kdTree As pdKDTree
    Set kdTree = New pdKDTree
    kdTree.BuildTree srcPalette, UBound(srcPalette) + 1
    
    'To test the array-backed implementation, use this setup:
    'Dim kdTree As pdKDTreeArray
    'Set kdTree = New pdKDTreeArray
    'kdTree.BuildTreeBalanced srcPalette, 0, UBound(srcPalette), False
    
    'Start matching pixels
    For y = 0 To finalY
        dstDIB.WrapArrayAroundScanline srcPixels, tmpSA, y
    For x = 0 To finalX Step pxSize
    
        b = srcPixels(x)
        g = srcPixels(x + 1)
        r = srcPixels(x + 2)
        
        'If this pixel matches the last pixel we tested, reuse our previous match results
        If (RGB(r, g, b) <> lastColor) Then
            
            tmpQuad.Red = r
            tmpQuad.Green = g
            tmpQuad.Blue = b
            
            'Ask the tree for its best match
            newQuad = kdTree.GetNearestColor(tmpQuad)
            
            lastColor = RGB(r, g, b)
            lastQuad = newQuad
            
        Else
            newQuad = lastQuad
        End If
        
        'Apply the closest discovered color to this pixel.
        srcPixels(x) = newQuad.Blue
        srcPixels(x + 1) = newQuad.Green
        srcPixels(x + 2) = newQuad.Red
        
    Next x
        If (Not suppressMessages) Then
            If (y And progBarCheck) = 0 Then
                If Interface.UserPressedESC() Then Exit For
                SetProgBarVal y + modifyProgBarOffset
            End If
        End If
    Next y
    
    dstDIB.UnwrapArrayFromDIB srcPixels
    
    ApplyPaletteToImage_KDTree = True
    
End Function

'Given an arbitrary palette (including palettes > 256 colors - they work just fine!), apply said palette to the
' target image.  Dithering is *not* used.  Alpha is included in palette matching calculations.  Colors are matched
' using a KD-tree (where the palette is pre-loaded into a tree, and colors are matched via that tree).
Public Function ApplyPaletteToImage_IncAlpha_KDTree(ByRef dstDIB As pdDIB, ByRef srcPalette() As RGBQuad, Optional ByVal suppressMessages As Boolean = False, Optional ByVal modifyProgBarMax As Long = -1, Optional ByVal modifyProgBarOffset As Long = 0) As Boolean

    Dim srcPixels() As Byte, tmpSA As SafeArray1D
    
    Dim pxSize As Long
    pxSize = dstDIB.GetDIBColorDepth \ 8
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = dstDIB.GetDIBStride - 1
    finalY = dstDIB.GetDIBHeight - 1
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates a
    ' refresh interval based on the size of the area to be processed.
    Dim progBarCheck As Long
    If (Not suppressMessages) Then
        If (modifyProgBarMax = -1) Then SetProgBarMax finalY Else SetProgBarMax modifyProgBarMax
        progBarCheck = ProgressBars.FindBestProgBarValue()
    End If
    
    'As with normal palette matching, we'll use basic RLE acceleration to try and skip palette
    ' searching for contiguous matching colors.
    Dim lastColor As Long: lastColor = -1
    Dim lastAlpha As Long: lastAlpha = -1
    Dim r As Long, g As Long, b As Long, a As Long
    
    Dim tmpQuad As RGBQuad, newQuad As RGBQuad, lastQuad As RGBQuad
    
    'Build the initial tree
    Dim kdTree As pdKDTree
    Set kdTree = New pdKDTree
    kdTree.BuildTreeIncAlpha srcPalette, UBound(srcPalette) + 1
    
    'Want to see the source palette?  Uncomment this code:
    'For x = 0 To UBound(srcPalette)
    '    Debug.Print srcPalette(x).Red, srcPalette(x).Green, srcPalette(x).Blue, srcPalette(x).Alpha
    'Next x
    
    'Start matching pixels
    For y = 0 To finalY
        dstDIB.WrapArrayAroundScanline srcPixels, tmpSA, y
    For x = 0 To finalX Step pxSize
    
        b = srcPixels(x)
        g = srcPixels(x + 1)
        r = srcPixels(x + 2)
        a = srcPixels(x + 3)
        
        'If this pixel matches the last pixel we tested, reuse our previous match results
        If ((RGB(r, g, b) <> lastColor) Or (a <> lastAlpha)) Then
            
            tmpQuad.Red = r
            tmpQuad.Green = g
            tmpQuad.Blue = b
            tmpQuad.Alpha = a
            
            'Ask the tree for its best match
            newQuad = kdTree.GetNearestColorIncAlpha(tmpQuad)
            
            lastColor = RGB(r, g, b)
            lastAlpha = a
            lastQuad = newQuad
            
        Else
            newQuad = lastQuad
        End If
        
        'Apply the closest discovered color to this pixel.
        With newQuad
            srcPixels(x) = .Blue
            srcPixels(x + 1) = .Green
            srcPixels(x + 2) = .Red
            srcPixels(x + 3) = .Alpha
        End With
        
    Next x
        If (Not suppressMessages) Then
            If (y And progBarCheck) = 0 Then
                If Interface.UserPressedESC() Then Exit For
                SetProgBarVal y + modifyProgBarOffset
            End If
        End If
    Next y
    
    dstDIB.UnwrapArrayFromDIB srcPixels
    
    ApplyPaletteToImage_IncAlpha_KDTree = True
    
End Function

'Given an arbitrary RGBA palette (including palettes > 256 colors - they work just fine!),
' apply said palette to a target image.  Dithering is *not* used.  Alpha is included in
' palette matching calculations, and all matches are performed in 8-bit ALab color space.
' Colors are matched using a KD-tree (where the palette is pre-loaded into a tree, and
' colors are matched via that tree).
Public Function ApplyPaletteToImage_IncAlpha_KDTree_Lab(ByRef dstDIB As pdDIB, ByRef srcPalette() As RGBQuad, Optional ByVal suppressMessages As Boolean = False, Optional ByVal modifyProgBarMax As Long = -1, Optional ByVal modifyProgBarOffset As Long = 0) As Boolean

    Dim srcPixels() As Byte, tmpSA As SafeArray1D
    Dim srcPixelsLab() As Byte
    
    'Resize the LAB array to the same size as a scanline of the source image
    ReDim srcPixelsLab(0 To dstDIB.GetDIBStride - 1) As Byte
    
    Dim pxSize As Long
    pxSize = dstDIB.GetDIBColorDepth \ 8
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = dstDIB.GetDIBStride - 1
    finalY = dstDIB.GetDIBHeight - 1
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates a
    ' refresh interval based on the size of the area to be processed.
    Dim progBarCheck As Long
    If (Not suppressMessages) Then
        If (modifyProgBarMax = -1) Then SetProgBarMax finalY Else SetProgBarMax modifyProgBarMax
        progBarCheck = ProgressBars.FindBestProgBarValue()
    End If
    
    'As with normal palette matching, we'll use basic RLE acceleration to try and skip palette
    ' searching for contiguous matching colors.
    Dim lastColor As Long: lastColor = -1
    Dim lastAlpha As Long: lastAlpha = -1
    Dim r As Long, g As Long, b As Long, a As Long
    
    Dim tmpQuad As RGBQuad, newQuad As RGBQuad, lastQuad As RGBQuad
    
    'Convert the palette to LabA format
    Dim labPalette() As RGBQuad
    ReDim labPalette(0 To UBound(srcPalette)) As RGBQuad
    
    Dim cRGB As pdLCMSProfile
    Set cRGB = New pdLCMSProfile
    cRGB.CreateSRGBProfile True
    
    Dim cLAB As pdLCMSProfile
    Set cLAB = New pdLCMSProfile
    cLAB.CreateLabProfile True
    
    Dim cTransform As pdLCMSTransform
    Set cTransform = New pdLCMSTransform
    cTransform.CreateTwoProfileTransform cRGB, cLAB, TYPE_BGRA_8, TYPE_ALab_8, INTENT_PERCEPTUAL
    
    cTransform.ApplyTransformToScanline VarPtr(srcPalette(0)), VarPtr(labPalette(0)), UBound(srcPalette) + 1
    
    'Build the initial tree
    Dim kdTree As pdKDTree
    Set kdTree = New pdKDTree
    kdTree.BuildTreeIncAlpha labPalette, UBound(srcPalette) + 1
    
    'Want to see the source palette?  Uncomment this code:
    'For x = 0 To UBound(srcPalette)
    '    Debug.Print "RGBA", srcPalette(x).Red, srcPalette(x).Green, srcPalette(x).Blue, srcPalette(x).Alpha
    '    Debug.Print "ALAB", labPalette(x).Red, labPalette(x).Green, labPalette(x).Blue, labPalette(x).Alpha
    'Next x
    
    'Start matching pixels
    For y = 0 To finalY
        dstDIB.WrapArrayAroundScanline srcPixels, tmpSA, y
        cTransform.ApplyTransformToScanline VarPtr(srcPixels(0)), VarPtr(srcPixelsLab(0)), dstDIB.GetDIBWidth
    For x = 0 To finalX Step pxSize
    
        b = srcPixels(x)
        g = srcPixels(x + 1)
        r = srcPixels(x + 2)
        a = srcPixels(x + 3)
        
        'If this pixel matches the last pixel we tested, reuse our previous match results
        If ((RGB(r, g, b) <> lastColor) Or (a <> lastAlpha)) Then
            
            tmpQuad.Blue = srcPixelsLab(x)
            tmpQuad.Green = srcPixelsLab(x + 1)
            tmpQuad.Red = srcPixelsLab(x + 2)
            tmpQuad.Alpha = srcPixelsLab(x + 3)
            
            'Ask the tree for its best match
            newQuad = srcPalette(kdTree.GetNearestPaletteIndexIncAlpha(tmpQuad))
            
            lastColor = RGB(r, g, b)
            lastAlpha = a
            lastQuad = newQuad
            
        Else
            newQuad = lastQuad
        End If
        
        'Apply the closest discovered color to this pixel.
        With newQuad
            srcPixels(x) = .Blue
            srcPixels(x + 1) = .Green
            srcPixels(x + 2) = .Red
            srcPixels(x + 3) = .Alpha
        End With
        
    Next x
        If (Not suppressMessages) Then
            If (y And progBarCheck) = 0 Then
                If Interface.UserPressedESC() Then Exit For
                SetProgBarVal y + modifyProgBarOffset
            End If
        End If
    Next y
    
    dstDIB.UnwrapArrayFromDIB srcPixels
    
    ApplyPaletteToImage_IncAlpha_KDTree_Lab = True
    
End Function

'Given a source palette (ideally created by GetOptimizedPalette(), above), apply said palette to the target image.
' Dithering is *not* used.  Colors are matched using Windows APIs.
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

'Given an arbitrary source palette, apply said palette to the target image.
' Dithering *is* used.  Colors are matched using a KD-tree.  Alpha values are NOT used when matching.
Public Function ApplyPaletteToImage_Dithered(ByRef dstDIB As pdDIB, ByRef srcPalette() As RGBQuad, Optional ByVal ditherMethod As PD_DITHER_METHOD = PDDM_FloydSteinberg, Optional ByVal ditherStrength As Single = 1!, Optional ByVal suppressMessages As Boolean = False, Optional ByVal modifyProgBarMax As Long = -1, Optional ByVal modifyProgBarOffset As Long = 0) As Boolean

    Dim srcPixels() As Byte, tmpSA As SafeArray2D
    dstDIB.WrapArrayAroundDIB srcPixels, tmpSA
    
    Dim srcPixels1D() As Byte, tmpSA1D As SafeArray1D, srcPtr As Long, srcStride As Long
    
    Dim pxSize As Long
    pxSize = dstDIB.GetDIBColorDepth \ 8
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = dstDIB.GetDIBStride - 1
    finalY = dstDIB.GetDIBHeight - 1
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates a
    ' refresh interval based on the size of the area to be processed.
    Dim progBarCheck As Long
    If (Not suppressMessages) Then
        If (modifyProgBarMax = -1) Then SetProgBarMax finalY Else SetProgBarMax modifyProgBarMax
        progBarCheck = ProgressBars.FindBestProgBarValue()
    End If
    
    Dim r As Long, g As Long, b As Long, i As Long, j As Long
    Dim newQuad As RGBQuad, tmpQuad As RGBQuad
    
    'Validate dither strength
    If (ditherStrength < 0!) Then ditherStrength = 0!
    If (ditherStrength > 1!) Then ditherStrength = 1!
    
    'Build A KD-tree for fast palette matching
    Dim kdTree As pdKDTree
    Set kdTree = New pdKDTree
    kdTree.BuildTree srcPalette, UBound(srcPalette) + 1
    
    'Prep a dither table that matches the requested setting.  Note that ordered dithers are handled separately.
    Dim ditherTableI() As Byte, ditherDivisor As Single
    Dim xLeft As Long, xRight As Long, yDown As Long
    
    Dim orderedDitherInUse As Boolean
    orderedDitherInUse = (ditherMethod = PDDM_Ordered_Bayer4x4) Or (ditherMethod = PDDM_Ordered_Bayer8x8)
    
    If orderedDitherInUse Then
    
        'Ordered dithers are handled specially, because we don't need to track running errors (e.g. no dithering
        ' information is carried to neighboring pixels).  Instead, we simply use the dither tables to adjust our
        ' threshold values on-the-fly.
        Dim ditherRows As Long, ditherColumns As Long
        
        'First, prepare a dithering table
        Palettes.GetDitherTable ditherMethod, ditherTableI, ditherDivisor, xLeft, xRight, yDown
        
        If (ditherMethod = PDDM_Ordered_Bayer4x4) Then
            ditherRows = 3
            ditherColumns = 3
        ElseIf (ditherMethod = PDDM_Ordered_Bayer8x8) Then
            ditherRows = 7
            ditherColumns = 7
        End If
        
        'By default, ordered dither trees use a scale of [0, 255].  This works great for thresholding
        ' against pure black/white, but for color data, it leads to extreme shifts.  Reduce the strength
        ' of the table before continuing.
        For x = 0 To ditherRows
        For y = 0 To ditherColumns
            ditherTableI(x, y) = ditherTableI(x, y) \ 2
        Next y
        Next x
        
        'Apply the finished dither table to the image
        Dim ditherAmt As Long
        
        dstDIB.WrapArrayAroundScanline srcPixels1D, tmpSA1D, 0
        srcPtr = tmpSA1D.pvData
        srcStride = tmpSA1D.cElements
        
        For y = 0 To finalY
            tmpSA1D.pvData = srcPtr + (srcStride * y)
        For x = 0 To finalX Step pxSize
        
            b = srcPixels1D(x)
            g = srcPixels1D(x + 1)
            r = srcPixels1D(x + 2)
            
            'Add dither to each component
            ditherAmt = Int(ditherTableI(Int(x \ 4) And ditherRows, y And ditherColumns)) - 63
            ditherAmt = ditherAmt * ditherStrength
            
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
            
            'Retrieve the best-match color
            tmpQuad.Blue = b
            tmpQuad.Green = g
            tmpQuad.Red = r
            newQuad = kdTree.GetNearestColor(tmpQuad)
            
            srcPixels1D(x) = newQuad.Blue
            srcPixels1D(x + 1) = newQuad.Green
            srcPixels1D(x + 2) = newQuad.Red
            
        Next x
            If (Not suppressMessages) Then
                If (y And progBarCheck) = 0 Then
                    If Interface.UserPressedESC() Then Exit For
                    SetProgBarVal y + modifyProgBarOffset
                End If
            End If
        Next y
        
        dstDIB.UnwrapArrayFromDIB srcPixels1D
    
    'All error-diffusion dither methods are handled similarly
    Else
        
        Dim rError As Long, gError As Long, bError As Long
        Dim errorMult As Single
        
        'Retrieve a hard-coded dithering table matching the requested dither type
        Palettes.GetDitherTable ditherMethod, ditherTableI, ditherDivisor, xLeft, xRight, yDown
        If (ditherDivisor <> 0!) Then ditherDivisor = 1! / ditherDivisor
        
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
        
        dstDIB.WrapArrayAroundScanline srcPixels1D, tmpSA1D, 0
        srcPtr = tmpSA1D.pvData
        srcStride = tmpSA1D.cElements
        
        'Start calculating pixels.
        For y = 0 To finalY
            tmpSA1D.pvData = srcPtr + (srcStride * y)
        For x = 0 To finalX Step pxSize
        
            b = srcPixels1D(x)
            g = srcPixels1D(x + 1)
            r = srcPixels1D(x + 2)
            
            'Add our running errors to the original colors
            xNonStride = x \ 4
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
            tmpQuad.Blue = newB
            tmpQuad.Green = newG
            tmpQuad.Red = newR
            newQuad = kdTree.GetNearestColor(tmpQuad)
            
            With newQuad
            
                'Apply the closest discovered color to this pixel.
                srcPixels1D(x) = .Blue
                srcPixels1D(x + 1) = .Green
                srcPixels1D(x + 2) = .Red
            
                'Calculate new errors
                rError = newR - CLng(.Red)
                gError = newG - CLng(.Green)
                bError = newB - CLng(.Blue)
                
            End With
            
            'Reduce color bleed, if specified
            rError = rError * ditherStrength
            gError = gError * ditherStrength
            bError = bError * ditherStrength
            
            'Spread any remaining error to neighboring pixels, using the precalculated dither table as our guide
            For i = xLeft To xRight
            For j = 0 To yDown
                
                If (ditherTableI(i, j) <> 0) Then
                    
                    xQuickInner = xNonStride + i
                    
                    'Next, ignore target pixels that are off the image boundary
                    If (xQuickInner >= initX) Then
                        If (xQuickInner < xWidth) Then
                        
                            'If we've made it all the way here, we are able to actually spread the error to this location
                            errorMult = CSng(ditherTableI(i, j)) * ditherDivisor
                            rErrors(xQuickInner, j) = rErrors(xQuickInner, j) + (rError * errorMult)
                            gErrors(xQuickInner, j) = gErrors(xQuickInner, j) + (gError * errorMult)
                            bErrors(xQuickInner, j) = bErrors(xQuickInner, j) + (bError * errorMult)
                            
                        End If
                    End If
                    
                End If
                
            Next j
            Next i
            
        Next x
        
            'When moving to the next line, we need to "shift" all accumulated errors upward.
            ' (Basically, what was previously the "next" line, is now the "current" line.
            ' The last line of errors must also be zeroed-out.
            If (yDown > 0) Then
            
                CopyMemoryStrict VarPtr(rErrors(0, 0)), VarPtr(rErrors(0, 1)), (xWidth + 1) * 4
                CopyMemoryStrict VarPtr(gErrors(0, 0)), VarPtr(gErrors(0, 1)), (xWidth + 1) * 4
                CopyMemoryStrict VarPtr(bErrors(0, 0)), VarPtr(bErrors(0, 1)), (xWidth + 1) * 4
                
                If (yDown = 1) Then
                    FillMemory VarPtr(rErrors(0, 1)), (xWidth + 1) * 4, 0
                    FillMemory VarPtr(gErrors(0, 1)), (xWidth + 1) * 4, 0
                    FillMemory VarPtr(bErrors(0, 1)), (xWidth + 1) * 4, 0
                Else
                    CopyMemoryStrict VarPtr(rErrors(0, 1)), VarPtr(rErrors(0, 2)), (xWidth + 1) * 4
                    CopyMemoryStrict VarPtr(gErrors(0, 1)), VarPtr(gErrors(0, 2)), (xWidth + 1) * 4
                    CopyMemoryStrict VarPtr(bErrors(0, 1)), VarPtr(bErrors(0, 2)), (xWidth + 1) * 4
                    
                    FillMemory VarPtr(rErrors(0, 2)), (xWidth + 1) * 4, 0
                    FillMemory VarPtr(gErrors(0, 2)), (xWidth + 1) * 4, 0
                    FillMemory VarPtr(bErrors(0, 2)), (xWidth + 1) * 4, 0
                End If
                
            Else
                FillMemory VarPtr(rErrors(0, 0)), (xWidth + 1) * 4, 0
                FillMemory VarPtr(gErrors(0, 0)), (xWidth + 1) * 4, 0
                FillMemory VarPtr(bErrors(0, 0)), (xWidth + 1) * 4, 0
            End If
            
            'Update the progress bar, as necessary
            If (Not suppressMessages) Then
                If (y And progBarCheck) = 0 Then
                    If Interface.UserPressedESC() Then Exit For
                    SetProgBarVal y + modifyProgBarOffset
                End If
            End If
            
        Next y
        
        dstDIB.UnwrapArrayFromDIB srcPixels1D
    
    End If
    
    dstDIB.UnwrapArrayFromDIB srcPixels
    
    ApplyPaletteToImage_Dithered = True
    
End Function

'Given an arbitrary source palette, apply said palette to the target image.
' Dithering *is* used.  Colors are matched using a KD-tree.  Alpha values are used when matching.
Public Function ApplyPaletteToImage_Dithered_IncAlpha(ByRef dstDIB As pdDIB, ByRef srcPalette() As RGBQuad, Optional ByVal ditherMethod As PD_DITHER_METHOD = PDDM_FloydSteinberg, Optional ByVal ditherStrength As Single = 1!, Optional ByVal suppressMessages As Boolean = False, Optional ByVal modifyProgBarMax As Long = -1, Optional ByVal modifyProgBarOffset As Long = 0) As Boolean

    Dim srcPixels() As Byte, tmpSA As SafeArray2D
    dstDIB.WrapArrayAroundDIB srcPixels, tmpSA
    
    Dim srcPixels1D() As Byte, tmpSA1D As SafeArray1D, srcPtr As Long, srcStride As Long
    
    Dim pxSize As Long
    pxSize = dstDIB.GetDIBColorDepth \ 8
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = dstDIB.GetDIBStride - 1
    finalY = dstDIB.GetDIBHeight - 1
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates a
    ' refresh interval based on the size of the area to be processed.
    Dim progBarCheck As Long
    If (Not suppressMessages) Then
        If (modifyProgBarMax = -1) Then SetProgBarMax finalY Else SetProgBarMax modifyProgBarMax
        progBarCheck = ProgressBars.FindBestProgBarValue()
    End If
    
    Dim r As Long, g As Long, b As Long, a As Long, i As Long, j As Long
    Dim newQuad As RGBQuad, tmpQuad As RGBQuad
    
    'Validate dither strength
    If (ditherStrength < 0!) Then ditherStrength = 0!
    If (ditherStrength > 1!) Then ditherStrength = 1!
    
    'Build A KD-tree for fast palette matching
    Dim kdTree As pdKDTree
    Set kdTree = New pdKDTree
    kdTree.BuildTreeIncAlpha srcPalette, UBound(srcPalette) + 1
    
    'Prep a dither table that matches the requested setting.  Note that ordered dithers are handled separately.
    Dim ditherTableI() As Byte, ditherDivisor As Single
    Dim xLeft As Long, xRight As Long, yDown As Long
    
    Dim orderedDitherInUse As Boolean
    orderedDitherInUse = (ditherMethod = PDDM_Ordered_Bayer4x4) Or (ditherMethod = PDDM_Ordered_Bayer8x8)
    
    If orderedDitherInUse Then
    
        'Ordered dithers are handled specially, because we don't need to track running errors (e.g. no dithering
        ' information is carried to neighboring pixels).  Instead, we simply use the dither tables to adjust our
        ' threshold values on-the-fly.
        Dim ditherRows As Long, ditherColumns As Long
        
        'First, prepare a dithering table
        Palettes.GetDitherTable ditherMethod, ditherTableI, ditherDivisor, xLeft, xRight, yDown
        
        If (ditherMethod = PDDM_Ordered_Bayer4x4) Then
            ditherRows = 3
            ditherColumns = 3
        ElseIf (ditherMethod = PDDM_Ordered_Bayer8x8) Then
            ditherRows = 7
            ditherColumns = 7
        End If
        
        'By default, ordered dither trees use a scale of [0, 255].  This works great for thresholding
        ' against pure black/white, but for color data, it leads to extreme shifts.  Reduce the strength
        ' of the table before continuing.
        For x = 0 To ditherRows
        For y = 0 To ditherColumns
            ditherTableI(x, y) = ditherTableI(x, y) \ 2
        Next y
        Next x
        
        'Apply the finished dither table to the image
        Dim ditherAmt As Long
        
        dstDIB.WrapArrayAroundScanline srcPixels1D, tmpSA1D, 0
        srcPtr = tmpSA1D.pvData
        srcStride = tmpSA1D.cElements
        
        For y = 0 To finalY
            tmpSA1D.pvData = srcPtr + (srcStride * y)
        For x = 0 To finalX Step pxSize
        
            b = srcPixels1D(x)
            g = srcPixels1D(x + 1)
            r = srcPixels1D(x + 2)
            a = srcPixels1D(x + 3)
            
            'Add dither to each component
            ditherAmt = Int(ditherTableI(Int(x \ 4) And ditherRows, y And ditherColumns)) - 63
            ditherAmt = ditherAmt * ditherStrength
            
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
            
            a = a + ditherAmt
            If (a > 255) Then
                a = 255
            ElseIf (a < 0) Then
                a = 0
            End If
            
            'Retrieve the best-match color
            tmpQuad.Blue = b
            tmpQuad.Green = g
            tmpQuad.Red = r
            tmpQuad.Alpha = a
            newQuad = kdTree.GetNearestColorIncAlpha(tmpQuad)
            
            srcPixels1D(x) = newQuad.Blue
            srcPixels1D(x + 1) = newQuad.Green
            srcPixels1D(x + 2) = newQuad.Red
            srcPixels1D(x + 3) = newQuad.Alpha
            
        Next x
            If (Not suppressMessages) Then
                If (y And progBarCheck) = 0 Then
                    If Interface.UserPressedESC() Then Exit For
                    SetProgBarVal y + modifyProgBarOffset
                End If
            End If
        Next y
        
        dstDIB.UnwrapArrayFromDIB srcPixels1D
    
    'All error-diffusion dither methods are handled similarly
    Else
        
        Dim rError As Long, gError As Long, bError As Long, aError As Long
        Dim errorMult As Single
        
        'Retrieve a hard-coded dithering table matching the requested dither type
        Palettes.GetDitherTable ditherMethod, ditherTableI, ditherDivisor, xLeft, xRight, yDown
        If (ditherDivisor <> 0!) Then ditherDivisor = 1! / ditherDivisor
        
        'Next, build an error tracking array.  Some diffusion methods require three rows worth of others;
        ' others require two.  Note that errors must be tracked separately for each color component.
        Dim xWidth As Long
        xWidth = dstDIB.GetDIBWidth - 1
        Dim rErrors() As Single, gErrors() As Single, bErrors() As Single, aErrors() As Single
        ReDim rErrors(0 To xWidth, 0 To yDown) As Single
        ReDim gErrors(0 To xWidth, 0 To yDown) As Single
        ReDim bErrors(0 To xWidth, 0 To yDown) As Single
        ReDim aErrors(0 To xWidth, 0 To yDown) As Single
        
        Dim xNonStride As Long, xQuickInner As Long
        Dim newR As Long, newG As Long, newB As Long, newA As Long
        
        dstDIB.WrapArrayAroundScanline srcPixels1D, tmpSA1D, 0
        srcPtr = tmpSA1D.pvData
        srcStride = tmpSA1D.cElements
        
        'Start calculating pixels.
        For y = 0 To finalY
            tmpSA1D.pvData = srcPtr + (srcStride * y)
        For x = 0 To finalX Step pxSize
        
            b = srcPixels1D(x)
            g = srcPixels1D(x + 1)
            r = srcPixels1D(x + 2)
            a = srcPixels1D(x + 3)
            
            'Add our running errors to the original colors
            xNonStride = x \ 4
            newR = r + rErrors(xNonStride, 0)
            newG = g + gErrors(xNonStride, 0)
            newB = b + bErrors(xNonStride, 0)
            newA = a + aErrors(xNonStride, 0)
            
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
            
            If (newA > 255) Then
                newA = 255
            ElseIf (newA < 0) Then
                newA = 0
            End If
            
            'Find the best palette match
            tmpQuad.Blue = newB
            tmpQuad.Green = newG
            tmpQuad.Red = newR
            tmpQuad.Alpha = newA
            newQuad = kdTree.GetNearestColorIncAlpha(tmpQuad)
            
            With newQuad
            
                'Apply the closest discovered color to this pixel.
                srcPixels1D(x) = .Blue
                srcPixels1D(x + 1) = .Green
                srcPixels1D(x + 2) = .Red
                srcPixels1D(x + 3) = .Alpha
            
                'Calculate new errors
                rError = newR - CLng(.Red)
                gError = newG - CLng(.Green)
                bError = newB - CLng(.Blue)
                aError = newA - CLng(.Alpha)
                
            End With
            
            'Reduce color bleed, if specified
            rError = rError * ditherStrength
            gError = gError * ditherStrength
            bError = bError * ditherStrength
            aError = aError * ditherStrength
            
            'Spread any remaining error to neighboring pixels, using the precalculated dither table as our guide
            For i = xLeft To xRight
            For j = 0 To yDown
                
                If (ditherTableI(i, j) <> 0) Then
                    
                    xQuickInner = xNonStride + i
                    
                    'Next, ignore target pixels that are off the image boundary
                    If (xQuickInner >= initX) Then
                        If (xQuickInner < xWidth) Then
                        
                            'If we've made it all the way here, we are able to actually spread the error to this location
                            errorMult = CSng(ditherTableI(i, j)) * ditherDivisor
                            rErrors(xQuickInner, j) = rErrors(xQuickInner, j) + (rError * errorMult)
                            gErrors(xQuickInner, j) = gErrors(xQuickInner, j) + (gError * errorMult)
                            bErrors(xQuickInner, j) = bErrors(xQuickInner, j) + (bError * errorMult)
                            aErrors(xQuickInner, j) = aErrors(xQuickInner, j) + (aError * errorMult)
                            
                        End If
                    End If
                    
                End If
                
            Next j
            Next i
            
        Next x
        
            'When moving to the next line, we need to "shift" all accumulated errors upward.
            ' (Basically, what was previously the "next" line, is now the "current" line.
            ' The last line of errors must also be zeroed-out.
            If (yDown > 0) Then
            
                CopyMemoryStrict VarPtr(rErrors(0, 0)), VarPtr(rErrors(0, 1)), (xWidth + 1) * 4
                CopyMemoryStrict VarPtr(gErrors(0, 0)), VarPtr(gErrors(0, 1)), (xWidth + 1) * 4
                CopyMemoryStrict VarPtr(bErrors(0, 0)), VarPtr(bErrors(0, 1)), (xWidth + 1) * 4
                CopyMemoryStrict VarPtr(aErrors(0, 0)), VarPtr(aErrors(0, 1)), (xWidth + 1) * 4
                
                If (yDown = 1) Then
                    FillMemory VarPtr(rErrors(0, 1)), (xWidth + 1) * 4, 0
                    FillMemory VarPtr(gErrors(0, 1)), (xWidth + 1) * 4, 0
                    FillMemory VarPtr(bErrors(0, 1)), (xWidth + 1) * 4, 0
                    FillMemory VarPtr(aErrors(0, 1)), (xWidth + 1) * 4, 0
                Else
                    CopyMemoryStrict VarPtr(rErrors(0, 1)), VarPtr(rErrors(0, 2)), (xWidth + 1) * 4
                    CopyMemoryStrict VarPtr(gErrors(0, 1)), VarPtr(gErrors(0, 2)), (xWidth + 1) * 4
                    CopyMemoryStrict VarPtr(bErrors(0, 1)), VarPtr(bErrors(0, 2)), (xWidth + 1) * 4
                    CopyMemoryStrict VarPtr(aErrors(0, 1)), VarPtr(aErrors(0, 2)), (xWidth + 1) * 4
                    
                    FillMemory VarPtr(rErrors(0, 2)), (xWidth + 1) * 4, 0
                    FillMemory VarPtr(gErrors(0, 2)), (xWidth + 1) * 4, 0
                    FillMemory VarPtr(bErrors(0, 2)), (xWidth + 1) * 4, 0
                    FillMemory VarPtr(aErrors(0, 2)), (xWidth + 1) * 4, 0
                End If
                
            Else
                FillMemory VarPtr(rErrors(0, 0)), (xWidth + 1) * 4, 0
                FillMemory VarPtr(gErrors(0, 0)), (xWidth + 1) * 4, 0
                FillMemory VarPtr(bErrors(0, 0)), (xWidth + 1) * 4, 0
                FillMemory VarPtr(aErrors(0, 0)), (xWidth + 1) * 4, 0
            End If
            
            'Update the progress bar, as necessary
            If (Not suppressMessages) Then
                If (y And progBarCheck) = 0 Then
                    If Interface.UserPressedESC() Then Exit For
                    SetProgBarVal y + modifyProgBarOffset
                End If
            End If
            
        Next y
        
        dstDIB.UnwrapArrayFromDIB srcPixels1D
    
    End If
    
    dstDIB.UnwrapArrayFromDIB srcPixels
    
    ApplyPaletteToImage_Dithered_IncAlpha = True
    
End Function

'Given an arbitrary source palette, apply said palette to the target image.
' Dithering *is* used.  Colors are matched using a KD-tree.  Alpha values are used when matching.
Public Function ApplyPaletteToImage_Dithered_IncAlpha_Lab(ByRef dstDIB As pdDIB, ByRef srcPalette() As RGBQuad, Optional ByVal ditherMethod As PD_DITHER_METHOD = PDDM_FloydSteinberg, Optional ByVal ditherStrength As Single = 1!, Optional ByVal suppressMessages As Boolean = False, Optional ByVal modifyProgBarMax As Long = -1, Optional ByVal modifyProgBarOffset As Long = 0) As Boolean

    Dim srcPixels() As Byte, tmpSA As SafeArray2D
    dstDIB.WrapArrayAroundDIB srcPixels, tmpSA
    
    Dim srcPixels1D() As Byte, tmpSA1D As SafeArray1D, srcPtr As Long, srcStride As Long
    
    'Comparisons will be done in LAB color space; resize a LAB array to the same size
    ' as a scanline of the source image (we'll convert lines as we go)
    Dim srcPixelsLab() As Byte
    ReDim srcPixelsLab(0 To dstDIB.GetDIBStride - 1) As Byte
    
    Dim pxSize As Long
    pxSize = dstDIB.GetDIBColorDepth \ 8
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = dstDIB.GetDIBStride - 1
    finalY = dstDIB.GetDIBHeight - 1
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates a
    ' refresh interval based on the size of the area to be processed.
    Dim progBarCheck As Long
    If (Not suppressMessages) Then
        If (modifyProgBarMax = -1) Then SetProgBarMax finalY Else SetProgBarMax modifyProgBarMax
        progBarCheck = ProgressBars.FindBestProgBarValue()
    End If
    
    Dim r As Long, g As Long, b As Long, a As Long, i As Long, j As Long
    Dim newQuad As RGBQuad, tmpQuad As RGBQuad
    
    'Validate dither strength
    If (ditherStrength < 0!) Then ditherStrength = 0!
    If (ditherStrength > 1!) Then ditherStrength = 1!
    
    'Convert the palette to LabA format
    Dim labPalette() As RGBQuad
    ReDim labPalette(0 To UBound(srcPalette)) As RGBQuad
    
    Dim cRGB As pdLCMSProfile
    Set cRGB = New pdLCMSProfile
    cRGB.CreateSRGBProfile True
    
    Dim cLAB As pdLCMSProfile
    Set cLAB = New pdLCMSProfile
    cLAB.CreateLabProfile True
    
    Dim cTransform As pdLCMSTransform
    Set cTransform = New pdLCMSTransform
    cTransform.CreateTwoProfileTransform cRGB, cLAB, TYPE_BGRA_8, TYPE_ALab_8, INTENT_PERCEPTUAL
    cTransform.ApplyTransformToScanline VarPtr(srcPalette(0)), VarPtr(labPalette(0)), UBound(srcPalette) + 1
    
    'Build A KD-tree for fast palette matching
    Dim kdTree As pdKDTree
    Set kdTree = New pdKDTree
    kdTree.BuildTreeIncAlpha labPalette, UBound(srcPalette) + 1
    
    'Prep a dither table that matches the requested setting.  Note that ordered dithers are handled separately.
    Dim ditherTableI() As Byte, ditherDivisor As Single
    Dim xLeft As Long, xRight As Long, yDown As Long
    
    Dim orderedDitherInUse As Boolean
    orderedDitherInUse = (ditherMethod = PDDM_Ordered_Bayer4x4) Or (ditherMethod = PDDM_Ordered_Bayer8x8)
    
    If orderedDitherInUse Then
    
        'Ordered dithers are handled specially, because we don't need to track running errors (e.g. no dithering
        ' information is carried to neighboring pixels).  Instead, we simply use the dither tables to adjust our
        ' threshold values on-the-fly.
        Dim ditherRows As Long, ditherColumns As Long
        
        'First, prepare a dithering table
        Palettes.GetDitherTable ditherMethod, ditherTableI, ditherDivisor, xLeft, xRight, yDown
        
        If (ditherMethod = PDDM_Ordered_Bayer4x4) Then
            ditherRows = 3
            ditherColumns = 3
        ElseIf (ditherMethod = PDDM_Ordered_Bayer8x8) Then
            ditherRows = 7
            ditherColumns = 7
        End If
        
        'By default, ordered dither trees use a scale of [0, 255].  This works great for thresholding
        ' against pure black/white, but for color data, it leads to extreme shifts.  Reduce the strength
        ' of the table before continuing.
        For x = 0 To ditherRows
        For y = 0 To ditherColumns
            ditherTableI(x, y) = ditherTableI(x, y) \ 2
        Next y
        Next x
        
        'Apply the finished dither table to the image
        Dim ditherAmt As Long
        
        dstDIB.WrapArrayAroundScanline srcPixels1D, tmpSA1D, 0
        srcPtr = tmpSA1D.pvData
        srcStride = tmpSA1D.cElements
        
        For y = 0 To finalY
            tmpSA1D.pvData = srcPtr + (srcStride * y)
            cTransform.ApplyTransformToScanline VarPtr(srcPixels1D(0)), VarPtr(srcPixelsLab(0)), dstDIB.GetDIBWidth
        For x = 0 To finalX Step pxSize
        
            b = srcPixelsLab(x)
            g = srcPixelsLab(x + 1)
            r = srcPixelsLab(x + 2)
            a = srcPixelsLab(x + 3)
            
            'Add dither to each component
            ditherAmt = Int(ditherTableI(Int(x \ 4) And ditherRows, y And ditherColumns)) - 63
            ditherAmt = ditherAmt * ditherStrength
            
            r = r + ditherAmt
            If (r > 255) Then r = 255
            If (r < 0) Then r = 0
            
            g = g + ditherAmt
            If (g > 255) Then g = 255
            If (g < 0) Then g = 0
            
            b = b + ditherAmt
            If (b > 255) Then b = 255
            If (b < 0) Then b = 0
            
            a = a + ditherAmt
            If (a > 255) Then a = 255
            If (a < 0) Then a = 0
            
            'Retrieve the best-match color
            tmpQuad.Blue = b
            tmpQuad.Green = g
            tmpQuad.Red = r
            tmpQuad.Alpha = a
            newQuad = srcPalette(kdTree.GetNearestPaletteIndexIncAlpha(tmpQuad))
            
            srcPixels1D(x) = newQuad.Blue
            srcPixels1D(x + 1) = newQuad.Green
            srcPixels1D(x + 2) = newQuad.Red
            srcPixels1D(x + 3) = newQuad.Alpha
            
        Next x
            If (Not suppressMessages) Then
                If (y And progBarCheck) = 0 Then
                    If Interface.UserPressedESC() Then Exit For
                    SetProgBarVal y + modifyProgBarOffset
                End If
            End If
        Next y
        
        dstDIB.UnwrapArrayFromDIB srcPixels1D
    
    'All error-diffusion dither methods are handled similarly
    Else
        
        Dim rError As Long, gError As Long, bError As Long, aError As Long
        Dim errorMult As Single
        
        'Retrieve a hard-coded dithering table matching the requested dither type
        Palettes.GetDitherTable ditherMethod, ditherTableI, ditherDivisor, xLeft, xRight, yDown
        If (ditherDivisor <> 0!) Then ditherDivisor = 1! / ditherDivisor
        
        'Next, build an error tracking array.  Some diffusion methods require three rows worth of others;
        ' others require two.  Note that errors must be tracked separately for each color component.
        Dim xWidth As Long
        xWidth = dstDIB.GetDIBWidth - 1
        Dim rErrors() As Single, gErrors() As Single, bErrors() As Single, aErrors() As Single
        ReDim rErrors(0 To xWidth, 0 To yDown) As Single
        ReDim gErrors(0 To xWidth, 0 To yDown) As Single
        ReDim bErrors(0 To xWidth, 0 To yDown) As Single
        ReDim aErrors(0 To xWidth, 0 To yDown) As Single
        
        Dim xNonStride As Long, xQuickInner As Long
        Dim newR As Long, newG As Long, newB As Long, newA As Long, newIndex As Long
        
        dstDIB.WrapArrayAroundScanline srcPixels1D, tmpSA1D, 0
        srcPtr = tmpSA1D.pvData
        srcStride = tmpSA1D.cElements
        
        'Start calculating pixels.
        For y = 0 To finalY
            tmpSA1D.pvData = srcPtr + (srcStride * y)
            cTransform.ApplyTransformToScanline VarPtr(srcPixels1D(0)), VarPtr(srcPixelsLab(0)), dstDIB.GetDIBWidth
        For x = 0 To finalX Step pxSize
        
            b = srcPixelsLab(x)
            g = srcPixelsLab(x + 1)
            r = srcPixelsLab(x + 2)
            a = srcPixelsLab(x + 3)
            
            'Add our running errors to the original colors
            xNonStride = x \ 4
            newR = r + rErrors(xNonStride, 0)
            newG = g + gErrors(xNonStride, 0)
            newB = b + bErrors(xNonStride, 0)
            newA = a + aErrors(xNonStride, 0)
            
            If (newR > 255) Then newR = 255
            If (newR < 0) Then newR = 0
            
            If (newG > 255) Then newG = 255
            If (newG < 0) Then newG = 0
            
            If (newB > 255) Then newB = 255
            If (newB < 0) Then newB = 0
            
            If (newA > 255) Then newA = 255
            If (newA < 0) Then newA = 0
            
            'Find the best palette match
            tmpQuad.Blue = newB
            tmpQuad.Green = newG
            tmpQuad.Red = newR
            tmpQuad.Alpha = newA
            newIndex = kdTree.GetNearestPaletteIndexIncAlpha(tmpQuad)
            
            With srcPalette(newIndex)
            
                'Apply the closest discovered color to this pixel.
                srcPixels1D(x) = .Blue
                srcPixels1D(x + 1) = .Green
                srcPixels1D(x + 2) = .Red
                srcPixels1D(x + 3) = .Alpha
            
            End With
            
            With labPalette(newIndex)
                
                'Calculate new errors
                rError = newR - CLng(.Red)
                gError = newG - CLng(.Green)
                bError = newB - CLng(.Blue)
                aError = newA - CLng(.Alpha)
                
            End With
            
            'Reduce color bleed, if specified
            rError = rError * ditherStrength
            gError = gError * ditherStrength
            bError = bError * ditherStrength
            aError = aError * ditherStrength
            
            'Spread any remaining error to neighboring pixels, using the precalculated dither table as our guide
            For i = xLeft To xRight
            For j = 0 To yDown
                
                If (ditherTableI(i, j) <> 0) Then
                    
                    xQuickInner = xNonStride + i
                    
                    'Next, ignore target pixels that are off the image boundary
                    If (xQuickInner >= initX) Then
                        If (xQuickInner < xWidth) Then
                        
                            'If we've made it all the way here, we are able to actually spread the error to this location
                            errorMult = CSng(ditherTableI(i, j)) * ditherDivisor
                            rErrors(xQuickInner, j) = rErrors(xQuickInner, j) + (rError * errorMult)
                            gErrors(xQuickInner, j) = gErrors(xQuickInner, j) + (gError * errorMult)
                            bErrors(xQuickInner, j) = bErrors(xQuickInner, j) + (bError * errorMult)
                            aErrors(xQuickInner, j) = aErrors(xQuickInner, j) + (aError * errorMult)
                            
                        End If
                    End If
                    
                End If
                
            Next j
            Next i
            
        Next x
        
            'When moving to the next line, we need to "shift" all accumulated errors upward.
            ' (Basically, what was previously the "next" line, is now the "current" line.
            ' The last line of errors must also be zeroed-out.
            If (yDown > 0) Then
            
                CopyMemoryStrict VarPtr(rErrors(0, 0)), VarPtr(rErrors(0, 1)), (xWidth + 1) * 4
                CopyMemoryStrict VarPtr(gErrors(0, 0)), VarPtr(gErrors(0, 1)), (xWidth + 1) * 4
                CopyMemoryStrict VarPtr(bErrors(0, 0)), VarPtr(bErrors(0, 1)), (xWidth + 1) * 4
                CopyMemoryStrict VarPtr(aErrors(0, 0)), VarPtr(aErrors(0, 1)), (xWidth + 1) * 4
                
                If (yDown = 1) Then
                    FillMemory VarPtr(rErrors(0, 1)), (xWidth + 1) * 4, 0
                    FillMemory VarPtr(gErrors(0, 1)), (xWidth + 1) * 4, 0
                    FillMemory VarPtr(bErrors(0, 1)), (xWidth + 1) * 4, 0
                    FillMemory VarPtr(aErrors(0, 1)), (xWidth + 1) * 4, 0
                Else
                    CopyMemoryStrict VarPtr(rErrors(0, 1)), VarPtr(rErrors(0, 2)), (xWidth + 1) * 4
                    CopyMemoryStrict VarPtr(gErrors(0, 1)), VarPtr(gErrors(0, 2)), (xWidth + 1) * 4
                    CopyMemoryStrict VarPtr(bErrors(0, 1)), VarPtr(bErrors(0, 2)), (xWidth + 1) * 4
                    CopyMemoryStrict VarPtr(aErrors(0, 1)), VarPtr(aErrors(0, 2)), (xWidth + 1) * 4
                    
                    FillMemory VarPtr(rErrors(0, 2)), (xWidth + 1) * 4, 0
                    FillMemory VarPtr(gErrors(0, 2)), (xWidth + 1) * 4, 0
                    FillMemory VarPtr(bErrors(0, 2)), (xWidth + 1) * 4, 0
                    FillMemory VarPtr(aErrors(0, 2)), (xWidth + 1) * 4, 0
                End If
                
            Else
                FillMemory VarPtr(rErrors(0, 0)), (xWidth + 1) * 4, 0
                FillMemory VarPtr(gErrors(0, 0)), (xWidth + 1) * 4, 0
                FillMemory VarPtr(bErrors(0, 0)), (xWidth + 1) * 4, 0
                FillMemory VarPtr(aErrors(0, 0)), (xWidth + 1) * 4, 0
            End If
            
            'Update the progress bar, as necessary
            If (Not suppressMessages) Then
                If (y And progBarCheck) = 0 Then
                    If Interface.UserPressedESC() Then Exit For
                    SetProgBarVal y + modifyProgBarOffset
                End If
            End If
            
        Next y
        
        dstDIB.UnwrapArrayFromDIB srcPixels1D
    
    End If
    
    dstDIB.UnwrapArrayFromDIB srcPixels
    
    ApplyPaletteToImage_Dithered_IncAlpha_Lab = True
    
End Function

'Populate a dithering table and relevant markers based on a specific dithering type.
' Returns: TRUE if successful; FALSE otherwise.  Note that some dither types (e.g. ordered dithers) do not
' use this function; they are handled specially.
Public Function GetDitherTable(ByVal ditherType As PD_DITHER_METHOD, ByRef dstDitherTable() As Byte, ByRef ditherDivisor As Single, ByRef xLeft As Long, ByRef xRight As Long, ByRef yDown As Long) As Boolean
    
    GetDitherTable = True
    
    Dim x As Long, y As Long
    
    Select Case ditherType
    
        Case PDDM_Ordered_Bayer4x4
        
            ReDim dstDitherTable(0 To 3, 0 To 3) As Byte
            
            dstDitherTable(0, 0) = 0
            dstDitherTable(0, 1) = 8
            dstDitherTable(0, 2) = 2
            dstDitherTable(0, 3) = 10
            
            dstDitherTable(1, 0) = 12
            dstDitherTable(1, 1) = 4
            dstDitherTable(1, 2) = 14
            dstDitherTable(1, 3) = 6
            
            dstDitherTable(2, 0) = 3
            dstDitherTable(2, 1) = 11
            dstDitherTable(2, 2) = 1
            dstDitherTable(2, 3) = 9
            
            dstDitherTable(3, 0) = 15
            dstDitherTable(3, 1) = 7
            dstDitherTable(3, 2) = 13
            dstDitherTable(3, 3) = 5
    
            'Scale the table to [0, 255] range
            For x = 0 To 3
            For y = 0 To 3
                dstDitherTable(x, y) = dstDitherTable(x, y) * 16
            Next y
            Next x
        
        Case PDDM_Ordered_Bayer8x8
            
            ReDim dstDitherTable(0 To 7, 0 To 7) As Byte
            
            dstDitherTable(0, 0) = 0
            dstDitherTable(0, 1) = 48
            dstDitherTable(0, 2) = 12
            dstDitherTable(0, 3) = 60
            dstDitherTable(0, 4) = 3
            dstDitherTable(0, 5) = 51
            dstDitherTable(0, 6) = 15
            dstDitherTable(0, 7) = 63
            
            dstDitherTable(1, 0) = 32
            dstDitherTable(1, 1) = 16
            dstDitherTable(1, 2) = 44
            dstDitherTable(1, 3) = 28
            dstDitherTable(1, 4) = 35
            dstDitherTable(1, 5) = 19
            dstDitherTable(1, 6) = 47
            dstDitherTable(1, 7) = 31
            
            dstDitherTable(2, 0) = 8
            dstDitherTable(2, 1) = 56
            dstDitherTable(2, 2) = 4
            dstDitherTable(2, 3) = 52
            dstDitherTable(2, 4) = 11
            dstDitherTable(2, 5) = 59
            dstDitherTable(2, 6) = 7
            dstDitherTable(2, 7) = 55
            
            dstDitherTable(3, 0) = 40
            dstDitherTable(3, 1) = 24
            dstDitherTable(3, 2) = 36
            dstDitherTable(3, 3) = 20
            dstDitherTable(3, 4) = 43
            dstDitherTable(3, 5) = 27
            dstDitherTable(3, 6) = 39
            dstDitherTable(3, 7) = 23
            
            dstDitherTable(4, 0) = 2
            dstDitherTable(4, 1) = 50
            dstDitherTable(4, 2) = 14
            dstDitherTable(4, 3) = 62
            dstDitherTable(4, 4) = 1
            dstDitherTable(4, 5) = 49
            dstDitherTable(4, 6) = 13
            dstDitherTable(4, 7) = 61
            
            dstDitherTable(5, 0) = 34
            dstDitherTable(5, 1) = 18
            dstDitherTable(5, 2) = 46
            dstDitherTable(5, 3) = 30
            dstDitherTable(5, 4) = 33
            dstDitherTable(5, 5) = 17
            dstDitherTable(5, 6) = 45
            dstDitherTable(5, 7) = 29
    
            dstDitherTable(6, 0) = 10
            dstDitherTable(6, 1) = 58
            dstDitherTable(6, 2) = 6
            dstDitherTable(6, 3) = 54
            dstDitherTable(6, 4) = 9
            dstDitherTable(6, 5) = 57
            dstDitherTable(6, 6) = 5
            dstDitherTable(6, 7) = 53
            
            dstDitherTable(7, 0) = 42
            dstDitherTable(7, 1) = 26
            dstDitherTable(7, 2) = 38
            dstDitherTable(7, 3) = 22
            dstDitherTable(7, 4) = 41
            dstDitherTable(7, 5) = 25
            dstDitherTable(7, 6) = 37
            dstDitherTable(7, 7) = 21
            
            'Scale the table to [0, 255] range
            For x = 0 To 7
            For y = 0 To 7
                dstDitherTable(x, y) = dstDitherTable(x, y) * 4
            Next y
            Next x
            
        Case PDDM_SingleNeighbor
        
            ReDim dstDitherTable(0 To 1, 0) As Byte
            
            dstDitherTable(1, 0) = 1
            ditherDivisor = 1
            
            xLeft = 0
            xRight = 1
            yDown = 0
            
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
    
    Dim cdFilter As pdString
    Set cdFilter = New pdString
    cdFilter.Append g_Language.TranslateMessage("All supported palettes") & "|*.aco;*.act;*.ase;*.gpl;*.pal;*.pdpalette;*.psppalette;*.txt|"
    
    cdFilter.Append g_Language.TranslateMessage("Adobe Color Swatch") & " (.aco)|*.aco|"
    cdFilter.Append g_Language.TranslateMessage("Adobe Color Table") & " (.act)|*.act|"
    cdFilter.Append g_Language.TranslateMessage("Adobe Swatch Exchange") & " (.ase)|*.ase|"
    cdFilter.Append g_Language.TranslateMessage("GIMP Palette") & " (.gpl)|*.gpl|"
    cdFilter.Append g_Language.TranslateMessage("Paint.NET Palette") & " (.txt)|*.txt|"
    cdFilter.Append g_Language.TranslateMessage("PaintShop Pro Palette") & " (.pal, .psppalette)|*.pal;*.psppalette|"
    cdFilter.Append g_Language.TranslateMessage("PhotoDemon Palette") & " (.pdpalette)|*.pdpalette|"
    cdFilter.Append g_Language.TranslateMessage("All files") & "|*.*"
    
    Dim cdTitle As String
    cdTitle = g_Language.TranslateMessage("Select a palette")
            
    'Prep a common dialog interface
    Dim openDialog As pdOpenSaveDialog
    Set openDialog = New pdOpenSaveDialog
            
    Dim sFile As String
    sFile = srcFilename
    
    If openDialog.GetOpenFileName(sFile, vbNullString, True, False, cdFilter.ToString(), 1, UserPrefs.GetPalettePath, cdTitle, , GetModalOwner().hWnd) Then
    
        'By design, we don't perform any validation here.  Let the caller validate the file as much (or as little)
        ' as they require.
        DisplayPaletteLoadDialog = (LenB(sFile) <> 0)
        
        'The dialog was successful.  Return the path, and save this path for future usage.
        If DisplayPaletteLoadDialog Then
            UserPrefs.SetPalettePath sFile
            dstFilename = sFile
        Else
            dstFilename = vbNullString
        End If
        
    End If
    
    'Re-enable user input
    Interface.EnableUserInput
    
End Function

'Display PD's generic palette export dialog.  All supported palette filetypes will be available to the user.
Public Function DisplayPaletteSaveDialog(ByRef srcImage As pdImage, ByRef dstFilename As String, ByRef dstFormat As PD_PaletteFormat) As Boolean
    
    DisplayPaletteSaveDialog = False
    
    'Disable user input until the dialog closes
    Interface.DisableUserInput
    
    'Prior to showing the "save palette" dialog, we need to determine three things:
    ' 1) An initial folder
    ' 2) What palette format to suggest
    ' 3) What filename to suggest (*without* a file extension)
    ' 4) What filename + extension to suggest, based on the results of 2 and 3
    
    'Each of these will be handled in turn
    
    '1) Determine an initial folder.  This is easy - just grab the last "palette" path from the preferences file.
    '   (The preferences engine will automatically pass us PD's local palette folder if no "last path" entry exists.)
    Dim initialSaveFolder As String
    initialSaveFolder = UserPrefs.GetPalettePath
    
    '2) What palette format to suggest.  After building the export palette list, retrieve the last-used palette
    '   format index from the user prefs file.
    Dim cdFilter As pdString, cdFilterExtensions As pdString
    Set cdFilter = New pdString
    Set cdFilterExtensions = New pdString
    
    cdFilter.Append g_Language.TranslateMessage("Adobe Color Swatch") & " (.aco)|*.aco|"
    cdFilterExtensions.Append ".aco|"
    cdFilter.Append g_Language.TranslateMessage("Adobe Color Table") & " (.act)|*.act|"
    cdFilterExtensions.Append ".act|"
    cdFilter.Append g_Language.TranslateMessage("Adobe Swatch Exchange") & " (.ase)|*.ase|"
    cdFilterExtensions.Append ".ase|"
    cdFilter.Append g_Language.TranslateMessage("GIMP Palette") & " (.gpl)|*.gpl|"
    cdFilterExtensions.Append ".gpl|"
    cdFilter.Append g_Language.TranslateMessage("Paint.NET Palette") & " (.txt)|*.txt|"
    cdFilterExtensions.Append ".txt"
    cdFilter.Append g_Language.TranslateMessage("PaintShop Pro Palette") & " (.pal)|*.pal|"
    cdFilterExtensions.Append ".pal|"
    cdFilter.Append g_Language.TranslateMessage("PhotoDemon Palette") & " (.pdpalette)|*.pdpalette|"
    cdFilterExtensions.Append ".pdpalette|"
    
    Dim cdIndex As PD_PaletteFormat
    cdIndex = UserPrefs.GetPref_Long("Saving", "Palette Format", pdpf_PhotoDemon) + 1
    
    '3) What palette name to suggest.  At present, we just reuse the current image's name.
    Dim palFileName As String
    palFileName = srcImage.ImgStorage.GetEntry_String("OriginalFileName", vbNullString)
    If (LenB(palFileName) = 0) Then palFileName = g_Language.TranslateMessage("New palette")
    palFileName = initialSaveFolder & palFileName
    
    Dim cdTitle As String
    cdTitle = g_Language.TranslateMessage("Export palette")
    
    'Prep a common dialog interface
    Dim saveDialog As pdOpenSaveDialog
    Set saveDialog = New pdOpenSaveDialog
    
    If saveDialog.GetSaveFileName(palFileName, , True, cdFilter.ToString(), cdIndex, UserPrefs.GetPalettePath, cdTitle, cdFilterExtensions.ToString(), GetModalOwner().hWnd) Then
    
        'Update preferences
        UserPrefs.SetPref_Long "Saving", "Palette Format", cdIndex - 1
        UserPrefs.SetPalettePath Files.FileGetPath(palFileName)
        
        'Notify the caller of the new settings
        dstFilename = palFileName
        dstFormat = cdIndex - 1
        DisplayPaletteSaveDialog = True
        
    End If
    
    'Re-enable user input
    Interface.EnableUserInput
    
End Function

Public Function ExportCurrentImagePalette(ByRef srcImage As pdImage, Optional ByVal exportParams As String = vbNullString) As Boolean
    
    'At present, a source image is *required*
    If (srcImage Is Nothing) Then Exit Function
    
    'Start by getting a destination filename and palette format from the user
    Dim dstFilename As String, dstFormat As PD_PaletteFormat
    If Palettes.DisplayPaletteSaveDialog(srcImage, dstFilename, dstFormat) Then
    
        'Before exporting, we need to get export preferences for the current format.  (Some formats support
        ' additional custom features; others do not.)
        
        'Disable user input until the next dialog closes
        Interface.DisableUserInput
        
        Dim exportSettings As String
        If (Dialogs.PromptPaletteSettings(srcImage, dstFormat, dstFilename, exportSettings) = vbOK) Then
            
            Message "Exporting palette..."
            
            'Parse settings and perform the actual export
            Dim cParams As pdSerialize
            Set cParams = New pdSerialize
            cParams.SetParamString exportSettings
            
            Dim cPalette As pdPalette
            Set cPalette = New pdPalette
            
            Dim numColors As Long, optColors As Long
            Dim palName As String
            
            With cParams
                
                'Before retrieving the actual palette, retrieve the number of colors we need to use.
                ' (If we have to generate an optimal palette, we want to know this in advance.)
                numColors = .GetLong("numColors", -1)
                If (numColors <= 0) Then
                    
                    'Paint.NET technically enforces a limit of 96 colors.  We currently ignore this
                    ' and write as many colors as we discover.
                    'If (dstFormat = pdpf_PaintDotNet) Then optColors = 96 Else optColors = 256
                    optColors = 256
                    
                Else
                    optColors = numColors
                End If
                
                If (.GetLong("srcPalette", 0) = 1) And srcImage.HasOriginalPalette Then
                    srcImage.GetOriginalPalette cPalette
                    If (optColors < cPalette.GetPaletteColorCount()) Then cPalette.SetNewPaletteCount optColors
                Else
                    Dim tmpDIB As pdDIB, tmpQuads() As RGBQuad
                    srcImage.GetCompositedImage tmpDIB, True
                    If Palettes.GetOptimizedPaletteIncAlpha(tmpDIB, tmpQuads, optColors, pdqs_Variance) Then
                        Palettes.SetPaletteAlphaPremultiplication False, tmpQuads
                        cPalette.CreateFromPaletteArray tmpQuads, UBound(tmpQuads) + 1
                    End If
                    Set tmpDIB = Nothing
                End If
                
                'Palette name won't always be used, but retrieve and set it anyway
                palName = .GetString("palName", vbNullString)
                cPalette.SetPaletteName palName
                
                'The actual export is handled by the palette object itself!
                If (cPalette.GetPaletteColorCount() > 0) Then
                
                    If (dstFormat = pdpf_AdobeColorSwatch) Then
                        ExportCurrentImagePalette = cPalette.SavePaletteAdobeSwatch(dstFilename)
                    ElseIf (dstFormat = pdpf_AdobeColorTable) Then
                        ExportCurrentImagePalette = cPalette.SavePaletteAdobeColorTable(dstFilename)
                    ElseIf (dstFormat = pdpf_AdobeSwatchExchange) Then
                        
                        'ASE is a unique format is it supports multiple embedded palettes, and we allow the
                        ' user to overwrite *OR* append this palette to an existing file, if any.
                        Dim tmpPalette As pdPalette
                        If (.GetLong("embedPaletteASE", 0) = 1) And Files.FileExists(dstFilename) Then
                        
                            'The user wants us to merge this palette with the existing file.  Yay?
                            
                            'Start by retrieving the current file; if this fails, we'll default to just writing
                            ' the current palette as-is.
                            Set tmpPalette = New pdPalette
                            If tmpPalette.LoadPaletteFromFile(dstFilename, False) Then
                            
                                'The palette appears to have loaded okay.  Append this palette to the end of it,
                                ' and if that works, swap palette references.
                                If tmpPalette.AppendExistingPalette(cPalette) Then Set cPalette = tmpPalette
                                
                            End If
                            
                        Else
                            'Standard behavior: overwrite the target file.  We don't need to do anything here.
                        End If
                        
                        ExportCurrentImagePalette = cPalette.SavePaletteAdobeSwatchExchange(dstFilename)
                        
                    ElseIf (dstFormat = pdpf_GIMP) Then
                        ExportCurrentImagePalette = cPalette.SavePaletteGIMP(dstFilename)
                    ElseIf (dstFormat = pdpf_PaintDotNet) Then
                        ExportCurrentImagePalette = cPalette.SavePalettePaintDotNet(dstFilename, , , False)
                    ElseIf (dstFormat = pdpf_PSP) Then
                        ExportCurrentImagePalette = cPalette.SavePalettePaintShopPro(dstFilename)
                    Else
                        ExportCurrentImagePalette = cPalette.SavePalettePhotoDemon(dstFilename)
                    End If
                    
                End If
                
            End With
        
        End If
        
        'Re-enable user input
        Interface.EnableUserInput
    
    End If
    
    Message "Finished."

End Function

'Merge two palettes into one (basePalette receives the merge).  Colors shared between the two palettes are
' automatically identified and skipped (e.g. the merged palette will only contain one copy of the color).

'Returns: the number of colors in the merged palette.
Public Function MergePalettes(ByRef basePalette() As RGBQuad, ByVal numColorsInBasePalette As Long, ByRef appendPalette() As RGBQuad, ByVal numColorsInAppendPalette As Long) As Long

    Dim i As Long, j As Long
    Dim matchFound As Boolean
    Dim chkColor1 As Long, chkColor2 As Long
    
    For i = 0 To numColorsInAppendPalette - 1
        
        matchFound = False
        
        For j = 0 To numColorsInBasePalette - 1
            GetMem4 VarPtr(basePalette(j)), chkColor1
            GetMem4 VarPtr(appendPalette(i)), chkColor2
            If (chkColor1 = chkColor2) Then
                matchFound = True
                Exit For
            End If
        Next j
        
        'If a match *wasn't* found, append this color to the merged palette
        If (Not matchFound) Then
            If (UBound(basePalette) < numColorsInBasePalette) Then ReDim Preserve basePalette(0 To numColorsInBasePalette * 2 - 1) As RGBQuad
            basePalette(numColorsInBasePalette) = appendPalette(i)
            numColorsInBasePalette = numColorsInBasePalette + 1
        End If
    
    Next i
    
    'Trim the UBound() of the merged palette
    If (numColorsInBasePalette <> UBound(basePalette) + 1) Then ReDim Preserve basePalette(0 To numColorsInBasePalette - 1) As RGBQuad
    MergePalettes = numColorsInBasePalette

End Function

'Bit RGB color reduction (no error diffusion)
Public Sub Palettize_BitRGB(ByRef dstDIB As pdDIB, ByVal rNumShades As Byte, ByVal gNumShades As Byte, ByVal bNumShades As Byte, Optional ByVal smartColors As Boolean = False, Optional ByVal suppressMessages As Boolean = False)
    
    Dim pxSize As Long, imageData() As Byte, tmpSA1D As SafeArray1D
    pxSize = dstDIB.GetDIBColorDepth \ 8
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = (dstDIB.GetDIBWidth - 1) * pxSize
    finalY = dstDIB.GetDIBHeight - 1
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    If (Not suppressMessages) Then
        If smartColors Then ProgressBars.SetProgBarMax finalY * 2 Else ProgressBars.SetProgBarMax finalY
        progBarCheck = ProgressBars.FindBestProgBarValue()
    End If
    
    'Color variables
    Dim r As Long, g As Long, b As Long
    Dim newR As Long, newG As Long, newB As Long
    Dim mR As Double, mG As Double, mB As Double
    
    'New code for so-called "color matching"
    Dim rLookup() As Long
    Dim gLookup() As Long
    Dim bLookup() As Long
    Dim countLookup() As Long
    
    'Validate input params
    If (rNumShades > 256) Then rNumShades = 256
    If (gNumShades > 256) Then gNumShades = 256
    If (bNumShades > 256) Then bNumShades = 256
    If (rNumShades < 2) Then rNumShades = 2
    If (gNumShades < 2) Then gNumShades = 2
    If (bNumShades < 2) Then bNumShades = 2
    
    rNumShades = rNumShades - 1
    gNumShades = gNumShades - 1
    bNumShades = bNumShades - 1
    
    ReDim rLookup(0 To rNumShades, 0 To gNumShades, 0 To bNumShades) As Long
    ReDim gLookup(0 To rNumShades, 0 To gNumShades, 0 To bNumShades) As Long
    ReDim bLookup(0 To rNumShades, 0 To gNumShades, 0 To bNumShades) As Long
    ReDim countLookup(0 To rNumShades, 0 To gNumShades, 0 To bNumShades) As Long
    
    'Prepare conversion look-up tables (which will make the actual color reduction much faster)
    mR = (255 / rNumShades)
    mG = (255 / gNumShades)
    mB = (255 / bNumShades)
    
    Dim rQuick(0 To 255) As Byte, gQuick(0 To 255) As Byte, bQuick(0 To 255) As Byte
    For x = 0 To 255
        rQuick(x) = Int((x / mR) + 0.5)
        gQuick(x) = Int((x / mG) + 0.5)
        bQuick(x) = Int((x / mB) + 0.5)
    Next x
    
    'Loop through each pixel in the image, converting values as we go
    For y = initY To finalY
        dstDIB.WrapArrayAroundScanline imageData, tmpSA1D, y
    For x = initX To finalX Step pxSize
        
        b = imageData(x)
        g = imageData(x + 1)
        r = imageData(x + 2)
        
        'Truncate R, G, and B values (posterize-style) into discreet increments.  0.5 is added for rounding purposes.
        newR = rQuick(r)
        newG = gQuick(g)
        newB = bQuick(b)
        
        'If we're doing color matching, place color values into a look-up table
        If smartColors Then
        
            rLookup(newR, newG, newB) = rLookup(newR, newG, newB) + r
            gLookup(newR, newG, newB) = gLookup(newR, newG, newB) + g
            bLookup(newR, newG, newB) = bLookup(newR, newG, newB) + b
            
            'Also, keep track of how many colors fall into this bucket (so we can later determine an average color)
            countLookup(newR, newG, newB) = countLookup(newR, newG, newB) + 1
            
        End If
        
        'Multiply the now-discretely divided R, G, and B values to (0-255) equivalents
        newR = Int(newR * mR + 0.5)
        newG = Int(newG * mG + 0.5)
        newB = Int(newB * mB + 0.5)
        
        'If we are *not* color-matching, assign color values immediately
        If (Not smartColors) Then
            imageData(x) = newB
            imageData(x + 1) = newG
            imageData(x + 2) = newR
        End If
        
    Next x
        If (Not suppressMessages) Then
            If (y And progBarCheck) = 0 Then
                If Interface.UserPressedESC() Then Exit For
                SetProgBarVal y
            End If
        End If
    Next y
    
    'Color matching requires extra work.  Perform a second loop through the image, replacing values with their
    ' average counterparts.
    If smartColors And (Not g_cancelCurrentAction) Then
    
        'Find average colors based on color counts
        For r = 0 To rNumShades
        For g = 0 To gNumShades
        For b = 0 To bNumShades
            If (countLookup(r, g, b) <> 0) Then
                rLookup(r, g, b) = Int(rLookup(r, g, b) / countLookup(r, g, b) + 0.5)
                gLookup(r, g, b) = Int(gLookup(r, g, b) / countLookup(r, g, b) + 0.5)
                bLookup(r, g, b) = Int(bLookup(r, g, b) / countLookup(r, g, b) + 0.5)
            End If
        Next b
        Next g
        Next r
        
        'Assign average colors back into the picture
        For y = initY To finalY
            dstDIB.WrapArrayAroundScanline imageData, tmpSA1D, y
        For x = initX To finalX Step pxSize
            
            newB = bQuick(imageData(x))
            newG = gQuick(imageData(x + 1))
            newR = rQuick(imageData(x + 2))
            
            imageData(x) = bLookup(newR, newG, newB)
            imageData(x + 1) = gLookup(newR, newG, newB)
            imageData(x + 2) = rLookup(newR, newG, newB)
            
        Next x
            If (Not suppressMessages) Then
                If (y And progBarCheck) = 0 Then
                    If Interface.UserPressedESC() Then Exit For
                    SetProgBarVal finalY + y
                End If
            End If
        Next y
        
    End If
    
    'Safely deallocate imageData()
    dstDIB.UnwrapArrayFromDIB imageData
    
End Sub

'Error Diffusion dithering to x# shades of color per component
Public Sub Palettize_BitRGB_Dither(ByRef dstDIB As pdDIB, ByVal rNumShades As Byte, ByVal gNumShades As Byte, ByVal bNumShades As Byte, ByVal ditherType As PD_DITHER_METHOD, ByVal ditherStrength As Single, Optional ByVal smartColors As Boolean = False, Optional ByVal suppressMessages As Boolean = False)
    
    Dim srcPixels1D() As Byte, tmpSA1D As SafeArray1D, srcPtr As Long, srcStride As Long
    
    Dim pxSize As Long
    pxSize = dstDIB.GetDIBColorDepth \ 8
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = (dstDIB.GetDIBWidth - 1) * pxSize
    finalY = dstDIB.GetDIBHeight - 1
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates a
    ' refresh interval based on the size of the area to be processed.
    Dim progBarCheck As Long
    If (Not suppressMessages) Then
        If smartColors Then ProgressBars.SetProgBarMax finalY * 2 Else ProgressBars.SetProgBarMax finalY
        progBarCheck = ProgressBars.FindBestProgBarValue()
    End If
    
    Dim r As Long, g As Long, b As Long
    Dim i As Long, j As Long
    Dim origR As Long, origG As Long, origB As Long
    Dim newR As Long, newG As Long, newB As Long
    Dim mR As Double, mG As Double, mB As Double
    
    'New code for so-called "color matching"
    Dim rLookup() As Long
    Dim gLookup() As Long
    Dim bLookup() As Long
    Dim countLookup() As Long
    
    'Validate input params
    If (rNumShades > 256) Then rNumShades = 256
    If (gNumShades > 256) Then gNumShades = 256
    If (bNumShades > 256) Then bNumShades = 256
    If (rNumShades < 2) Then rNumShades = 2
    If (gNumShades < 2) Then gNumShades = 2
    If (bNumShades < 2) Then bNumShades = 2
    
    rNumShades = rNumShades - 1
    gNumShades = gNumShades - 1
    bNumShades = bNumShades - 1
    
    ReDim rLookup(0 To rNumShades, 0 To gNumShades, 0 To bNumShades) As Long
    ReDim gLookup(0 To rNumShades, 0 To gNumShades, 0 To bNumShades) As Long
    ReDim bLookup(0 To rNumShades, 0 To gNumShades, 0 To bNumShades) As Long
    ReDim countLookup(0 To rNumShades, 0 To gNumShades, 0 To bNumShades) As Long
    
    'Prepare conversion look-up tables (which will make the actual color reduction much faster)
    mR = (255 / rNumShades)
    mG = (255 / gNumShades)
    mB = (255 / bNumShades)
    
    Dim rQuick(0 To 255) As Byte, gQuick(0 To 255) As Byte, bQuick(0 To 255) As Byte
    For x = 0 To 255
        rQuick(x) = Int((x / mR) + 0.5)
        gQuick(x) = Int((x / mG) + 0.5)
        bQuick(x) = Int((x / mB) + 0.5)
    Next x
    
    'Validate dither strength
    If (ditherStrength < 0!) Then ditherStrength = 0!
    If (ditherStrength > 1!) Then ditherStrength = 1!
    
    'Prep a dither table that matches the requested setting.  Note that ordered dithers are handled separately.
    Dim ditherTableI() As Byte, ditherDivisor As Single
    Dim xLeft As Long, xRight As Long, yDown As Long
    
    Dim orderedDitherInUse As Boolean
    orderedDitherInUse = (ditherType = PDDM_Ordered_Bayer4x4) Or (ditherType = PDDM_Ordered_Bayer8x8)
    
    If orderedDitherInUse Then
    
        'Ordered dithers are handled specially, because we don't need to track running errors (e.g. no dithering
        ' information is carried to neighboring pixels).  Instead, we simply use the dither tables to adjust our
        ' threshold values on-the-fly.
        Dim ditherRows As Long, ditherColumns As Long
        
        'First, prepare a dithering table
        Palettes.GetDitherTable ditherType, ditherTableI, ditherDivisor, xLeft, xRight, yDown
        
        If (ditherType = PDDM_Ordered_Bayer4x4) Then
            ditherRows = 3
            ditherColumns = 3
        ElseIf (ditherType = PDDM_Ordered_Bayer8x8) Then
            ditherRows = 7
            ditherColumns = 7
        End If
        
        'By default, ordered dither trees use a scale of [0, 255].  This works great for thresholding
        ' against pure black/white, but for color data, it leads to extreme shifts.  Reduce the strength
        ' of the table before continuing.
        For x = 0 To ditherRows
        For y = 0 To ditherColumns
            ditherTableI(x, y) = ditherTableI(x, y) \ 2
        Next y
        Next x
        
        'Apply the finished dither table to the image
        Dim ditherAmt As Long
        
        dstDIB.WrapArrayAroundScanline srcPixels1D, tmpSA1D, 0
        srcPtr = tmpSA1D.pvData
        srcStride = tmpSA1D.cElements
        
        For y = 0 To finalY
            tmpSA1D.pvData = srcPtr + (srcStride * y)
        For x = 0 To finalX Step pxSize
        
            b = srcPixels1D(x)
            g = srcPixels1D(x + 1)
            r = srcPixels1D(x + 2)
            origR = r
            origG = g
            origB = b
            
            'Add dither to each component
            ditherAmt = Int(ditherTableI(Int(x \ 4) And ditherRows, y And ditherColumns)) - 63
            ditherAmt = ditherAmt * ditherStrength
            
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
            
            'Posterize
            newR = rQuick(r)
            newG = gQuick(g)
            newB = bQuick(b)
            
            'If we're doing color matching, place color values into a look-up table
            If smartColors Then
            
                rLookup(newR, newG, newB) = rLookup(newR, newG, newB) + origR
                gLookup(newR, newG, newB) = gLookup(newR, newG, newB) + origG
                bLookup(newR, newG, newB) = bLookup(newR, newG, newB) + origB
                
                'Also, keep track of how many colors fall into this bucket (so we can later determine an average color)
                countLookup(newR, newG, newB) = countLookup(newR, newG, newB) + 1
                
            End If
                
            'Multiply the now-discretely divided R, G, and B values to (0-255) equivalents
            newR = Int(newR * mR + 0.5)
            newG = Int(newG * mG + 0.5)
            newB = Int(newB * mB + 0.5)
            
            srcPixels1D(x) = newB
            srcPixels1D(x + 1) = newG
            srcPixels1D(x + 2) = newR
            
        Next x
            If (Not suppressMessages) Then
                If (y And progBarCheck) = 0 Then
                    If Interface.UserPressedESC() Then Exit For
                    SetProgBarVal y
                End If
            End If
        Next y
        
        workingDIB.UnwrapArrayFromDIB srcPixels1D
    
    'All error-diffusion dither methods are handled similarly
    Else
        
        Dim rError As Long, gError As Long, bError As Long
        Dim errorMult As Single
        
        'Retrieve a hard-coded dithering table matching the requested dither type
        Palettes.GetDitherTable ditherType, ditherTableI, ditherDivisor, xLeft, xRight, yDown
        If (ditherDivisor <> 0!) Then ditherDivisor = 1! / ditherDivisor
        
        'Next, build an error tracking array.  Some diffusion methods require three rows worth of others;
        ' others require two.  Note that errors must be tracked separately for each color component.
        Dim xWidth As Long
        xWidth = dstDIB.GetDIBWidth - 1
        
        Dim rErrors() As Single, gErrors() As Single, bErrors() As Single
        ReDim rErrors(0 To xWidth, 0 To yDown) As Single
        ReDim gErrors(0 To xWidth, 0 To yDown) As Single
        ReDim bErrors(0 To xWidth, 0 To yDown) As Single
        
        Dim xNonStride As Long, xQuickInner As Long
        
        dstDIB.WrapArrayAroundScanline srcPixels1D, tmpSA1D, 0
        srcPtr = tmpSA1D.pvData
        srcStride = tmpSA1D.cElements
        
        'Start calculating pixels.
        For y = 0 To finalY
            tmpSA1D.pvData = srcPtr + (srcStride * y)
        For x = 0 To finalX Step pxSize
        
            b = srcPixels1D(x)
            g = srcPixels1D(x + 1)
            r = srcPixels1D(x + 2)
            origR = r
            origG = g
            origB = b
            
            'Add our running errors to the original colors
            xNonStride = x \ 4
            r = origR + rErrors(xNonStride, 0)
            g = origG + gErrors(xNonStride, 0)
            b = origB + bErrors(xNonStride, 0)
            
            If (r > 255) Then
                r = 255
            ElseIf (r < 0) Then
                r = 0
            End If
            
            If (g > 255) Then
                g = 255
            ElseIf (g < 0) Then
                g = 0
            End If
            
            If (b > 255) Then
                b = 255
            ElseIf (b < 0) Then
                b = 0
            End If
            
            'Posterize
            newR = rQuick(r)
            newG = gQuick(g)
            newB = bQuick(b)
            
            'If we're doing color matching, place color values into a look-up table
            If smartColors Then
            
                rLookup(newR, newG, newB) = rLookup(newR, newG, newB) + origR
                gLookup(newR, newG, newB) = gLookup(newR, newG, newB) + origG
                bLookup(newR, newG, newB) = bLookup(newR, newG, newB) + origB
                
                'Also, keep track of how many colors fall into this bucket (so we can later determine an average color)
                countLookup(newR, newG, newB) = countLookup(newR, newG, newB) + 1
                
            End If
            
            'Multiply the now-discretely divided R, G, and B values to (0-255) equivalents
            newR = Int(newR * mR + 0.5)
            newG = Int(newG * mG + 0.5)
            newB = Int(newB * mB + 0.5)
            
            srcPixels1D(x) = newB
            srcPixels1D(x + 1) = newG
            srcPixels1D(x + 2) = newR
            
            'Calculate new errors
            rError = r - newR
            gError = g - newG
            bError = b - newB
            
            'Reduce color bleed, if specified
            rError = rError * ditherStrength
            gError = gError * ditherStrength
            bError = bError * ditherStrength
            
            'Spread any remaining error to neighboring pixels, using the precalculated dither table as our guide
            For i = xLeft To xRight
            For j = 0 To yDown
                
                If (ditherTableI(i, j) <> 0) Then
                    
                    xQuickInner = xNonStride + i
                    
                    'Next, ignore target pixels that are off the image boundary
                    If (xQuickInner >= initX) Then
                        If (xQuickInner < xWidth) Then
                        
                            'If we've made it all the way here, we are able to actually spread the error to this location
                            errorMult = CSng(ditherTableI(i, j)) * ditherDivisor
                            rErrors(xQuickInner, j) = rErrors(xQuickInner, j) + (rError * errorMult)
                            gErrors(xQuickInner, j) = gErrors(xQuickInner, j) + (gError * errorMult)
                            bErrors(xQuickInner, j) = bErrors(xQuickInner, j) + (bError * errorMult)
                            
                        End If
                    End If
                    
                End If
                
            Next j
            Next i
            
        Next x
        
            'When moving to the next line, we need to "shift" all accumulated errors upward.
            ' (Basically, what was previously the "next" line, is now the "current" line.
            ' The last line of errors must also be zeroed-out.
            If (yDown > 0) Then
            
                CopyMemoryStrict VarPtr(rErrors(0, 0)), VarPtr(rErrors(0, 1)), (xWidth + 1) * 4
                CopyMemoryStrict VarPtr(gErrors(0, 0)), VarPtr(gErrors(0, 1)), (xWidth + 1) * 4
                CopyMemoryStrict VarPtr(bErrors(0, 0)), VarPtr(bErrors(0, 1)), (xWidth + 1) * 4
                
                If (yDown = 1) Then
                    FillMemory VarPtr(rErrors(0, 1)), (xWidth + 1) * 4, 0
                    FillMemory VarPtr(gErrors(0, 1)), (xWidth + 1) * 4, 0
                    FillMemory VarPtr(bErrors(0, 1)), (xWidth + 1) * 4, 0
                Else
                    CopyMemoryStrict VarPtr(rErrors(0, 1)), VarPtr(rErrors(0, 2)), (xWidth + 1) * 4
                    CopyMemoryStrict VarPtr(gErrors(0, 1)), VarPtr(gErrors(0, 2)), (xWidth + 1) * 4
                    CopyMemoryStrict VarPtr(bErrors(0, 1)), VarPtr(bErrors(0, 2)), (xWidth + 1) * 4
                    
                    FillMemory VarPtr(rErrors(0, 2)), (xWidth + 1) * 4, 0
                    FillMemory VarPtr(gErrors(0, 2)), (xWidth + 1) * 4, 0
                    FillMemory VarPtr(bErrors(0, 2)), (xWidth + 1) * 4, 0
                End If
                
            Else
                FillMemory VarPtr(rErrors(0, 0)), (xWidth + 1) * 4, 0
                FillMemory VarPtr(gErrors(0, 0)), (xWidth + 1) * 4, 0
                FillMemory VarPtr(bErrors(0, 0)), (xWidth + 1) * 4, 0
            End If
            
            'Update the progress bar, as necessary
            If (Not suppressMessages) Then
                If (y And progBarCheck) = 0 Then
                    If Interface.UserPressedESC() Then Exit For
                    SetProgBarVal y
                End If
            End If
            
        Next y
        
        dstDIB.UnwrapArrayFromDIB srcPixels1D
    
    End If
    
    'Color matching requires extra work.  Perform a second loop through the image, replacing values with their
    ' average counterparts.
    If smartColors And (Not g_cancelCurrentAction) Then
    
        'Find average colors based on color counts
        For r = 0 To rNumShades
        For g = 0 To gNumShades
        For b = 0 To bNumShades
            If (countLookup(r, g, b) <> 0) Then
                rLookup(r, g, b) = Int(rLookup(r, g, b) / countLookup(r, g, b) + 0.5)
                gLookup(r, g, b) = Int(gLookup(r, g, b) / countLookup(r, g, b) + 0.5)
                bLookup(r, g, b) = Int(bLookup(r, g, b) / countLookup(r, g, b) + 0.5)
            End If
        Next b
        Next g
        Next r
        
        'Assign average colors back into the picture
        For y = initY To finalY
            dstDIB.WrapArrayAroundScanline srcPixels1D, tmpSA1D, y
        For x = initX To finalX Step pxSize
            
            newB = bQuick(srcPixels1D(x))
            newG = gQuick(srcPixels1D(x + 1))
            newR = rQuick(srcPixels1D(x + 2))
            
            srcPixels1D(x) = bLookup(newR, newG, newB)
            srcPixels1D(x + 1) = gLookup(newR, newG, newB)
            srcPixels1D(x + 2) = rLookup(newR, newG, newB)
            
        Next x
            If (Not suppressMessages) Then
                If (y And progBarCheck) = 0 Then
                    If Interface.UserPressedESC() Then Exit For
                    SetProgBarVal finalY + y
                End If
            End If
        Next y
        
        dstDIB.UnwrapArrayFromDIB srcPixels1D
        
    End If
    
End Sub

'Several PD functions share the same dither features (monochrome, grayscale, color palettes, etc)
Public Sub PopulateDitheringDropdown(ByRef dstDropDown As pdDropDown)
    dstDropDown.SetAutomaticRedraws False
    dstDropDown.Clear
    dstDropDown.AddItem "None", 0
    dstDropDown.AddItem "Ordered (Bayer 4x4)", 1
    dstDropDown.AddItem "Ordered (Bayer 8x8)", 2
    dstDropDown.AddItem "Single neighbor", 3
    dstDropDown.AddItem "Floyd-Steinberg", 4
    dstDropDown.AddItem "Jarvis, Judice, and Ninke", 5
    dstDropDown.AddItem "Stucki", 6
    dstDropDown.AddItem "Burkes", 7
    dstDropDown.AddItem "Sierra-3", 8
    dstDropDown.AddItem "Two-Row Sierra", 9
    dstDropDown.AddItem "Sierra Lite", 10
    dstDropDown.AddItem "Atkinson / Classic Macintosh", 11
    dstDropDown.SetAutomaticRedraws True
End Sub

'If you don't want to deal with alpha values, use this function to forcibly set all alpha values in a palette to 255
' (or any other arbitrary value)
Public Sub SetFixedAlpha(ByRef srcQuads() As RGBQuad, Optional ByVal newAlpha As Byte = 255)
    Dim i As Long
    For i = 0 To UBound(srcQuads)
        srcQuads(i).Alpha = newAlpha
    Next i
End Sub

'If a palette includes alpha values, it can be helpful to forcibly set alpha premultiplication
Public Sub SetPaletteAlphaPremultiplication(ByVal applyPremultiplication As Boolean, ByRef srcQuads() As RGBQuad)
        
    Const ONE_DIV_255 As Double = 1# / 255#
    
    'Although most palettes are small (256 colors or less), PD supports "unlimited" palette sizes.
    ' As a perf failsafe, use a LUT for the conversion.
    Dim intToFloat() As Single
    ReDim intToFloat(0 To 255) As Single
    Dim i As Long
    For i = 0 To 255
        If applyPremultiplication Then
            intToFloat(i) = CSng(CDbl(i) * ONE_DIV_255)
        Else
            If (i <> 0) Then intToFloat(i) = CSng(255# / CDbl(i))
        End If
    Next i
    
    Dim r As Long, g As Long, b As Long
    Dim tmpAlpha As Byte, tmpAlphaModifier As Single
    
    For i = 0 To UBound(srcQuads)
        
        'Retrieve alpha for the current pixel
        tmpAlpha = srcQuads(i).Alpha
        
        'Branch according to applying or removing premultiplication
        If applyPremultiplication Then
        
            'When applying premultiplication, we can ignore fully opaque pixels
            If (tmpAlpha <> 255) Then
            
                'We can shortcut the calculation of full transparent pixels (they are made black)
                If (tmpAlpha = 0) Then
                    srcQuads(i).Red = 0
                    srcQuads(i).Green = 0
                    srcQuads(i).Blue = 0
                Else
            
                    r = srcQuads(i).Red
                    g = srcQuads(i).Green
                    b = srcQuads(i).Blue
                    
                    tmpAlphaModifier = intToFloat(tmpAlpha)
                    
                    'Remove premultiplied values by redistributing the colors based on this pixel's alpha value
                    r = Int(r * tmpAlphaModifier + 0.5!)
                    g = Int(g * tmpAlphaModifier + 0.5!)
                    b = Int(b * tmpAlphaModifier + 0.5!)
                    
                    srcQuads(i).Red = r
                    srcQuads(i).Green = g
                    srcQuads(i).Blue = b
                    
                End If
            
            End If
        
        Else
            
            'When removing premultiplication, we can ignore fully opaque and fully transparent values.
            ' (Note that VB doesn't short-circuit AND statements, so we manually nest the IFs.)
            If (tmpAlpha <> 255) Then
                If (tmpAlpha <> 0) Then
                
                    r = srcQuads(i).Red
                    g = srcQuads(i).Green
                    b = srcQuads(i).Blue
                    
                    tmpAlphaModifier = intToFloat(tmpAlpha)
                    
                    'Remove premultiplied values by redistributing the colors based on this pixel's alpha value
                    r = Int(r * tmpAlphaModifier + 0.5!)
                    g = Int(g * tmpAlphaModifier + 0.5!)
                    b = Int(b * tmpAlphaModifier + 0.5!)
                    
                    'Unfortunately, OOB checks are necessary for malformed colors
                    If (r > 255) Then r = 255
                    If (g > 255) Then g = 255
                    If (b > 255) Then b = 255
                    
                    srcQuads(i).Red = r
                    srcQuads(i).Green = g
                    srcQuads(i).Blue = b
                    
                End If
            End If
        
        End If
        
    Next i
    
End Sub

'Given an arbitrary source palette, apply said palette to the target image, and return the results
' not as a DIB, but as a standard byte array (in 1-byte-per-pixel format.)
'
'Dithering *is* used.  Colors are matched using a KD-tree.  Alpha values are considered when matching.
Public Function GetPalettizedImage_Dithered_IncAlpha(ByRef srcDIB As pdDIB, ByRef srcPalette() As RGBQuad, ByRef dstBytes() As Byte, Optional ByVal ditherMethod As PD_DITHER_METHOD = PDDM_FloydSteinberg, Optional ByVal ditherStrength As Single = 1!, Optional ByVal suppressMessages As Boolean = False, Optional ByVal modifyProgBarMax As Long = -1, Optional ByVal modifyProgBarOffset As Long = 0) As Boolean

    Dim srcPixels() As Byte, tmpSA As SafeArray2D
    srcDIB.WrapArrayAroundDIB srcPixels, tmpSA
        
    Dim srcPixels1D() As Byte, tmpSA1D As SafeArray1D, srcPtr As Long, srcStride As Long
    
    Dim pxSize As Long
    pxSize = srcDIB.GetDIBColorDepth \ 8
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = srcDIB.GetDIBStride - 1
    finalY = srcDIB.GetDIBHeight - 1
    
    'Prep destination array
    ReDim dstBytes(0 To srcDIB.GetDIBWidth - 1, 0 To srcDIB.GetDIBHeight - 1) As Byte
    
    'To avoid division on the inner loop, build a lut for x indices
    Dim xLookup() As Long
    ReDim xLookup(0 To finalX) As Long
    For x = 0 To finalX Step pxSize
        xLookup(x) = x \ pxSize
    Next x
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates a
    ' refresh interval based on the size of the area to be processed.
    Dim progBarCheck As Long
    If (Not suppressMessages) Then
        If (modifyProgBarMax = -1) Then SetProgBarMax finalY Else SetProgBarMax modifyProgBarMax
        progBarCheck = ProgressBars.FindBestProgBarValue()
    End If
    
    Dim r As Long, g As Long, b As Long, a As Long, i As Long, j As Long
    Dim newIndex As Long, tmpQuad As RGBQuad
    
    'Validate dither strength
    If (ditherStrength < 0!) Then ditherStrength = 0!
    If (ditherStrength > 1!) Then ditherStrength = 1!
    
    'Build A KD-tree for fast palette matching
    Dim kdTree As pdKDTree
    Set kdTree = New pdKDTree
    kdTree.BuildTreeIncAlpha srcPalette, UBound(srcPalette) + 1
    
    'Prep a dither table that matches the requested setting.  Note that ordered dithers are handled separately.
    Dim ditherTableI() As Byte, ditherDivisor As Single
    Dim xLeft As Long, xRight As Long, yDown As Long
    
    Dim orderedDitherInUse As Boolean
    orderedDitherInUse = (ditherMethod = PDDM_Ordered_Bayer4x4) Or (ditherMethod = PDDM_Ordered_Bayer8x8)
    
    If orderedDitherInUse Then
    
        'Ordered dithers are handled specially, because we don't need to track running errors (e.g. no dithering
        ' information is carried to neighboring pixels).  Instead, we simply use the dither tables to adjust our
        ' threshold values on-the-fly.
        Dim ditherRows As Long, ditherColumns As Long
        
        'First, prepare a dithering table
        Palettes.GetDitherTable ditherMethod, ditherTableI, ditherDivisor, xLeft, xRight, yDown
        
        If (ditherMethod = PDDM_Ordered_Bayer4x4) Then
            ditherRows = 3
            ditherColumns = 3
        ElseIf (ditherMethod = PDDM_Ordered_Bayer8x8) Then
            ditherRows = 7
            ditherColumns = 7
        End If
        
        'By default, ordered dither trees use a scale of [0, 255].  This works great for thresholding
        ' against pure black/white, but for color data, it leads to extreme shifts.  Reduce the strength
        ' of the table before continuing.
        For x = 0 To ditherRows
        For y = 0 To ditherColumns
            ditherTableI(x, y) = ditherTableI(x, y) \ 2
        Next y
        Next x
        
        'Apply the finished dither table to the image
        Dim ditherAmt As Long
        
        srcDIB.WrapArrayAroundScanline srcPixels1D, tmpSA1D, 0
        srcPtr = tmpSA1D.pvData
        srcStride = tmpSA1D.cElements
        
        For y = 0 To finalY
            tmpSA1D.pvData = srcPtr + (srcStride * y)
        For x = 0 To finalX Step pxSize
        
            b = srcPixels1D(x)
            g = srcPixels1D(x + 1)
            r = srcPixels1D(x + 2)
            a = srcPixels1D(x + 3)
            
            'Add dither to each component
            ditherAmt = Int(ditherTableI(Int(x \ 4) And ditherRows, y And ditherColumns)) - 63
            ditherAmt = ditherAmt * ditherStrength
            
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
            
            a = a + ditherAmt
            If (a > 255) Then
                a = 255
            ElseIf (a < 0) Then
                a = 0
            End If
            
            'Retrieve the best-match color
            tmpQuad.Blue = b
            tmpQuad.Green = g
            tmpQuad.Red = r
            tmpQuad.Alpha = a
            newIndex = kdTree.GetNearestPaletteIndexIncAlpha(tmpQuad)
            
            dstBytes(xLookup(x), y) = newIndex
            
        Next x
            If (Not suppressMessages) Then
                If (y And progBarCheck) = 0 Then
                    If Interface.UserPressedESC() Then Exit For
                    SetProgBarVal y + modifyProgBarOffset
                End If
            End If
        Next y
        
        srcDIB.UnwrapArrayFromDIB srcPixels1D
    
    'All error-diffusion dither methods are handled similarly
    Else
        
        Dim rError As Long, gError As Long, bError As Long, aError As Long
        Dim errorMult As Single
        
        'Retrieve a hard-coded dithering table matching the requested dither type
        Palettes.GetDitherTable ditherMethod, ditherTableI, ditherDivisor, xLeft, xRight, yDown
        If (ditherDivisor <> 0!) Then ditherDivisor = 1! / ditherDivisor
        
        'Next, build an error tracking array.  Some diffusion methods require three rows worth of others;
        ' others require two.  Note that errors must be tracked separately for each color component.
        Dim xWidth As Long
        xWidth = srcDIB.GetDIBWidth - 1
        Dim rErrors() As Single, gErrors() As Single, bErrors() As Single, aErrors() As Single
        ReDim rErrors(0 To xWidth, 0 To yDown) As Single
        ReDim gErrors(0 To xWidth, 0 To yDown) As Single
        ReDim bErrors(0 To xWidth, 0 To yDown) As Single
        ReDim aErrors(0 To xWidth, 0 To yDown) As Single
        
        Dim xNonStride As Long, xQuickInner As Long
        Dim newR As Long, newG As Long, newB As Long, newA As Long
        
        srcDIB.WrapArrayAroundScanline srcPixels1D, tmpSA1D, 0
        srcPtr = tmpSA1D.pvData
        srcStride = tmpSA1D.cElements
        
        'Start calculating pixels.
        For y = 0 To finalY
            tmpSA1D.pvData = srcPtr + (srcStride * y)
        For x = 0 To finalX Step pxSize
        
            b = srcPixels1D(x)
            g = srcPixels1D(x + 1)
            r = srcPixels1D(x + 2)
            a = srcPixels1D(x + 3)
            
            'Add our running errors to the original colors
            xNonStride = x \ 4
            newR = r + rErrors(xNonStride, 0)
            newG = g + gErrors(xNonStride, 0)
            newB = b + bErrors(xNonStride, 0)
            newA = a + aErrors(xNonStride, 0)
            
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
            
            If (newA > 255) Then
                newA = 255
            ElseIf (newA < 0) Then
                newA = 0
            End If
            
            'Find the best palette match
            tmpQuad.Blue = newB
            tmpQuad.Green = newG
            tmpQuad.Red = newR
            tmpQuad.Alpha = newA
            newIndex = kdTree.GetNearestPaletteIndexIncAlpha(tmpQuad)
            
            'Apply the closest discovered color to this pixel.
            dstBytes(xLookup(x), y) = newIndex
            
            'Calculate new errors
            With srcPalette(newIndex)
            
                'Calculate new errors
                rError = newR - CLng(.Red)
                gError = newG - CLng(.Green)
                bError = newB - CLng(.Blue)
                aError = newA - CLng(.Alpha)
                
            End With
            
            'Reduce color bleed, if specified
            rError = rError * ditherStrength
            gError = gError * ditherStrength
            bError = bError * ditherStrength
            aError = aError * ditherStrength
            
            'Spread any remaining error to neighboring pixels, using the precalculated dither table as our guide
            For i = xLeft To xRight
            For j = 0 To yDown
                
                If (ditherTableI(i, j) <> 0) Then
                    
                    xQuickInner = xNonStride + i
                    
                    'Next, ignore target pixels that are off the image boundary
                    If (xQuickInner >= initX) Then
                        If (xQuickInner < xWidth) Then
                        
                            'If we've made it all the way here, we are able to actually spread the error to this location
                            errorMult = CSng(ditherTableI(i, j)) * ditherDivisor
                            rErrors(xQuickInner, j) = rErrors(xQuickInner, j) + (rError * errorMult)
                            gErrors(xQuickInner, j) = gErrors(xQuickInner, j) + (gError * errorMult)
                            bErrors(xQuickInner, j) = bErrors(xQuickInner, j) + (bError * errorMult)
                            aErrors(xQuickInner, j) = aErrors(xQuickInner, j) + (aError * errorMult)
                            
                        End If
                    End If
                    
                End If
                
            Next j
            Next i
            
        Next x
        
            'When moving to the next line, we need to "shift" all accumulated errors upward.
            ' (Basically, what was previously the "next" line, is now the "current" line.
            ' The last line of errors must also be zeroed-out.
            If (yDown > 0) Then
            
                CopyMemoryStrict VarPtr(rErrors(0, 0)), VarPtr(rErrors(0, 1)), (xWidth + 1) * 4
                CopyMemoryStrict VarPtr(gErrors(0, 0)), VarPtr(gErrors(0, 1)), (xWidth + 1) * 4
                CopyMemoryStrict VarPtr(bErrors(0, 0)), VarPtr(bErrors(0, 1)), (xWidth + 1) * 4
                CopyMemoryStrict VarPtr(aErrors(0, 0)), VarPtr(aErrors(0, 1)), (xWidth + 1) * 4
                
                If (yDown = 1) Then
                    FillMemory VarPtr(rErrors(0, 1)), (xWidth + 1) * 4, 0
                    FillMemory VarPtr(gErrors(0, 1)), (xWidth + 1) * 4, 0
                    FillMemory VarPtr(bErrors(0, 1)), (xWidth + 1) * 4, 0
                    FillMemory VarPtr(aErrors(0, 1)), (xWidth + 1) * 4, 0
                Else
                    CopyMemoryStrict VarPtr(rErrors(0, 1)), VarPtr(rErrors(0, 2)), (xWidth + 1) * 4
                    CopyMemoryStrict VarPtr(gErrors(0, 1)), VarPtr(gErrors(0, 2)), (xWidth + 1) * 4
                    CopyMemoryStrict VarPtr(bErrors(0, 1)), VarPtr(bErrors(0, 2)), (xWidth + 1) * 4
                    CopyMemoryStrict VarPtr(aErrors(0, 1)), VarPtr(aErrors(0, 2)), (xWidth + 1) * 4
                    
                    FillMemory VarPtr(rErrors(0, 2)), (xWidth + 1) * 4, 0
                    FillMemory VarPtr(gErrors(0, 2)), (xWidth + 1) * 4, 0
                    FillMemory VarPtr(bErrors(0, 2)), (xWidth + 1) * 4, 0
                    FillMemory VarPtr(aErrors(0, 2)), (xWidth + 1) * 4, 0
                End If
                
            Else
                FillMemory VarPtr(rErrors(0, 0)), (xWidth + 1) * 4, 0
                FillMemory VarPtr(gErrors(0, 0)), (xWidth + 1) * 4, 0
                FillMemory VarPtr(bErrors(0, 0)), (xWidth + 1) * 4, 0
                FillMemory VarPtr(aErrors(0, 0)), (xWidth + 1) * 4, 0
            End If
            
            'Update the progress bar, as necessary
            If (Not suppressMessages) Then
                If (y And progBarCheck) = 0 Then
                    If Interface.UserPressedESC() Then Exit For
                    SetProgBarVal y + modifyProgBarOffset
                End If
            End If
            
        Next y
        
        srcDIB.UnwrapArrayFromDIB srcPixels1D
    
    End If
    
    srcDIB.UnwrapArrayFromDIB srcPixels
    
    GetPalettizedImage_Dithered_IncAlpha = True
    
End Function

'Given an arbitrary palette (including palettes > 256 colors - they work just fine!), match said palette to a
' target image and measure palette entry "popularity".  Then, redistribute said palette entries so that the
' most popular colors appear earliest in the palette.
'
'This improves performance on palette-based images in legacy RLE formats like PCX, because PCX uses high-value
' bytes as RLE flags.  (So single-occurrences of high-value bytes must be encoded as RLE runs of 1, while lower
' values can simply be written as-is.)
'
'Because this function is targeted at legacy formats specifically, alpha values are *not* considered.
'
'This operation is lossless for the DIB - it is treated as read-only - but the passed palette will obviously
' be modified by the function!
Public Function SortPaletteByPopularity_RGB(ByRef srcDIB As pdDIB, ByRef srcPalette() As RGBQuad) As Boolean
    
    Dim srcPixels() As Byte, tmpSA As SafeArray1D
    
    Dim pxSize As Long
    pxSize = srcDIB.GetDIBColorDepth \ 8
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = srcDIB.GetDIBStride - 1
    finalY = srcDIB.GetDIBHeight - 1
    
    'As with normal palette matching, we'll use basic RLE acceleration to try and skip palette
    ' searching for contiguous matching colors.
    Dim lastColor As Long: lastColor = -1
    Dim lastAlpha As Long: lastAlpha = -1
    Dim r As Long, g As Long, b As Long
    
    Dim tmpQuad As RGBQuad, palIndex As Long, lastPalIndex As Long
    
    'Build the initial tree
    Dim kdTree As pdKDTree
    Set kdTree = New pdKDTree
    kdTree.BuildTree srcPalette, UBound(srcPalette) + 1
    
    'Also construct a histogram; this is how we measure popularity
    Dim palPopularity() As Long
    ReDim palPopularity(0 To UBound(srcPalette)) As Long
    
    'Start matching pixels
    For y = 0 To finalY
        srcDIB.WrapArrayAroundScanline srcPixels, tmpSA, y
    For x = 0 To finalX Step pxSize
    
        b = srcPixels(x)
        g = srcPixels(x + 1)
        r = srcPixels(x + 2)
        
        'If this pixel matches the last pixel we tested, reuse our previous match results
        If (RGB(r, g, b) <> lastColor) Then
            
            tmpQuad.Red = r
            tmpQuad.Green = g
            tmpQuad.Blue = b
            
            'Ask the tree for its best match
            palIndex = kdTree.GetNearestPaletteIndex(tmpQuad)
            
            lastColor = RGB(r, g, b)
            lastPalIndex = palIndex
            
        Else
            palIndex = lastPalIndex
        End If
        
        'Increment the histogram for the matched palette index
        palPopularity(palIndex) = palPopularity(palIndex) + 1
        
    Next x
    Next y
    
    srcDIB.UnwrapArrayFromDIB srcPixels
    
    Dim i As Long, j As Long
    
    'Do a quick insertion sort.  Points are likely to be somewhat close to sorted, as the first color(s)
    ' we encounter are likely to consume most of the image, especially in e.g. GIFs.
    Dim numColors As Long
    numColors = UBound(srcPalette) + 1
    
    Dim tmpSortQ As RGBQuad, tmpSortL As Long, searchCont As Boolean
    i = 1
    
    Do While (i < numColors)
        tmpSortQ = srcPalette(i)
        tmpSortL = palPopularity(i)
        j = i - 1
        
        'Because VB6 doesn't short-circuit And statements, we split this check into separate parts.
        searchCont = False
        If (j >= 0) Then searchCont = (palPopularity(j) < tmpSortL)
        
        Do While searchCont
            srcPalette(j + 1) = srcPalette(j)
            palPopularity(j + 1) = palPopularity(j)
            j = j - 1
            searchCont = False
            If (j >= 0) Then searchCont = (palPopularity(j) < tmpSortL)
        Loop
        
        srcPalette(j + 1) = tmpSortQ
        palPopularity(j + 1) = tmpSortL
        i = i + 1
        
    Loop
    
    SortPaletteByPopularity_RGB = True
    
End Function

'Given an arbitrary palette (including palettes > 256 colors - they work just fine!), match said palette to a
' target image and measure palette entry "popularity".  Then, redistribute said palette entries so that the
' eight most popular colors are matched to power-of-two values (e.g. the most compressible indices).
'
'If the prioritizeAlpha parameter is set to TRUE, the palette value with transparency = 0 will be given
' highest precedence.  (This produces "more compatible" GIFs, since some GIF decoders expect index 0 to
' always represent transparency, if it exists.)
'
'This operation is lossless for the DIB - it is treated as read-only - but the passed palette will obviously
' be modified by the function!
Public Function SortPaletteForCompression_IncAlpha(ByRef srcDIB As pdDIB, ByRef srcPalette() As RGBQuad, Optional ByVal prioritizeAlpha As Boolean = True, Optional ByVal skipPopularityScan As Boolean = False) As Boolean
    
    If (Not skipPopularityScan) Then
        
        Dim srcPixels() As Byte, tmpSA As SafeArray1D
        
        Dim pxSize As Long
        pxSize = srcDIB.GetDIBColorDepth \ 8
        
        Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
        initX = 0
        initY = 0
        finalX = srcDIB.GetDIBStride - 1
        finalY = srcDIB.GetDIBHeight - 1
        
        'As with normal palette matching, we'll use basic RLE acceleration to try and skip palette
        ' searching for contiguous matching colors.
        Dim lastColor As Long: lastColor = -1
        Dim lastAlpha As Long: lastAlpha = -1
        Dim r As Long, g As Long, b As Long, a As Long
        
        Dim tmpQuad As RGBQuad, palIndex As Long, lastPalIndex As Long
        
        'Build the initial tree
        Dim kdTree As pdKDTree
        Set kdTree = New pdKDTree
        kdTree.BuildTreeIncAlpha srcPalette, UBound(srcPalette) + 1
        
        'Also construct a histogram; this is how we measure popularity
        Dim palPopularity() As Long
        ReDim palPopularity(0 To UBound(srcPalette)) As Long
        
        'Start matching pixels
        For y = 0 To finalY
            srcDIB.WrapArrayAroundScanline srcPixels, tmpSA, y
        For x = 0 To finalX Step pxSize
        
            b = srcPixels(x)
            g = srcPixels(x + 1)
            r = srcPixels(x + 2)
            a = srcPixels(x + 3)
            
            'If this pixel matches the last pixel we tested, reuse our previous match results
            If ((RGB(r, g, b) <> lastColor) Or (a <> lastAlpha)) Then
                
                tmpQuad.Red = r
                tmpQuad.Green = g
                tmpQuad.Blue = b
                tmpQuad.Alpha = a
                
                'Ask the tree for its best match
                palIndex = kdTree.GetNearestPaletteIndexIncAlpha(tmpQuad)
                
                lastColor = RGB(r, g, b)
                lastAlpha = a
                lastPalIndex = palIndex
                
            Else
                palIndex = lastPalIndex
            End If
            
            'Increment the histogram for the matched palette index
            palPopularity(palIndex) = palPopularity(palIndex) + 1
            
        Next x
        Next y
        
        srcDIB.UnwrapArrayFromDIB srcPixels
        
    End If
        
    Dim i As Long, j As Long
    
    'If transparency needs to be prioritized, find the highest popularity value and set the transparent
    ' index to that value.
    If prioritizeAlpha Then
        
        'If we're skipping the full popularity scan, don't bother with finding max popularity
        If (Not skipPopularityScan) Then
            Dim maxPopularity As Long
            For i = 0 To UBound(srcPalette)
                If (palPopularity(i) > maxPopularity) Then maxPopularity = palPopularity(i)
            Next i
        End If
        
        For i = 0 To UBound(srcPalette)
            
            'Transparent pixel...
            If (srcPalette(i).Alpha = 0) Then
                
                'If the caller doesn't want a full popularity scan, we can skip this step and
                ' simply "swap" the transparent index to the front of the image.
                If skipPopularityScan Then
                
                    If (i > 0) Then
                        tmpQuad = srcPalette(0)
                        srcPalette(0) = srcPalette(i)
                        srcPalette(i) = tmpQuad
                        Exit For
                    End If
                
                'In a normal popularity search, set all transparent pixels to "max+1" popularity
                Else
                    palPopularity(i) = maxPopularity + 1
                End If
                
            End If
            
        Next i
    
    End If
    
    'If the caller just wanted us to move alpha to the front of the palette, exit now
    If skipPopularityScan Then
        SortPaletteForCompression_IncAlpha = True
        Exit Function
    End If
    
    'Do a quick insertion sort.  Points are likely to be somewhat close to sorted, as the first color(s)
    ' we encounter are likely to consume most of the image, especially in e.g. GIFs.
    Dim numColors As Long
    numColors = UBound(srcPalette) + 1
    
    Dim tmpSortQ As RGBQuad, tmpSortL As Long, searchCont As Boolean
    i = 1
    
    Do While (i < numColors)
        tmpSortQ = srcPalette(i)
        tmpSortL = palPopularity(i)
        j = i - 1
        
        'Because VB6 doesn't short-circuit And statements, we split this check into separate parts.
        searchCont = False
        If (j >= 0) Then searchCont = (palPopularity(j) < tmpSortL)
        
        Do While searchCont
            srcPalette(j + 1) = srcPalette(j)
            palPopularity(j + 1) = palPopularity(j)
            j = j - 1
            searchCont = False
            If (j >= 0) Then searchCont = (palPopularity(j) < tmpSortL)
        Loop
        
        srcPalette(j + 1) = tmpSortQ
        palPopularity(j + 1) = tmpSortL
        i = i + 1
        
    Loop
    
    SortPaletteForCompression_IncAlpha = True
    
End Function

'Returns the number of colors written to the palette (always 16 for this function)
Public Function GetStockPalette(ByVal palID As PD_StockPalette, ByRef dstPalette() As RGBQuad, Optional ByVal initPaletteArrayForMe As Boolean = True) As Long
    
    'Different system palettes have different color counts
    Dim numColors As Long
    
    Select Case palID
        Case pdsp_EGA
            numColors = 16
        Case pdsp_PSLegacy
            numColors = 16
    End Select
    
    'If the caller doesn't require initialization, still check palette bounds for safety
    If initPaletteArrayForMe Then
        ReDim dstPalette(0 To numColors - 1) As RGBQuad
    Else
        If UBound(dstPalette) < numColors - 1 Then ReDim dstPalette(0 To numColors - 1) As RGBQuad
    End If
    
    'Grab a predefined string of palette entries
    Dim palAsHexString As String
    
    Select Case palID
        Case pdsp_EGA
            Const EGA_PAL_STRING As String = "000000,0000AA,00AA00,00AAAA,AA0000,AA00AA,AA5500,555555,AAAAAA,5555FF,55FF55,55FFFF,FF5555,FF55FF,FFFF55,FFFFFF"
            palAsHexString = EGA_PAL_STRING
        Case pdsp_PSLegacy
            Const PSLEGACY_PAL_STRING As String = "000000,0000aa,00aa00,00aaaa,aa0000,aa00aa,aaaa00,848484,c6c6c6,5555ff,55ff55,55ffff,ff5555,ff55ff,ffff55,ffffff"
            palAsHexString = PSLEGACY_PAL_STRING
    End Select
    
    'Split the color string into individual entries
    Dim listOfStrings() As String
    listOfStrings = Split(palAsHexString, ",", -1, vbBinaryCompare)
    
    If (UBound(listOfStrings) <> numColors - 1) Then PDDebug.LogAction "WARNING: Palettes.GetStockPalette failed unexpectedly"
    
    Dim i As Long
    For i = 0 To numColors - 1
        dstPalette(i) = Colors.GetRGBQuadFromHex(listOfStrings(i))
    Next i
    
End Function
