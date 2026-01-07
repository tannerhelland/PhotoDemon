VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdMakeTexture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Seamless Texture Generation class
'Copyright 2022-2026 by Tanner Helland
'Created: 21/April/22
'Last updated: 02/May/22
'Last update: split into separate class from in-painter; these two functions share a lot of code,
'             but there are some tight perf-sensitive loops that greatly benefit from specialized
'             optimization strategies that necessitate more code separation.
'
'Texture synthesis is an active field of study.  A great introduction to the topic is Paul Harrison's
' PhD thesis, "Image Texture Tools: Texture Synthesis, Texture Transfer, and Plausible Restoration",
' available here (link good as of April 2022):
'
'https://www.logarithmic.net/pfh-files/thesis/dissertation.pdf
'
'Paul wrote GIMP's "Resynthesizer" plugin to demonstrate his take on the algorithm, and as of 2022
' his algorithm is still used in many GPL programs (and is still available in GIMP, too).
'
'A few years after Paul's paper, Adobe released their take on the topic, in an algorithm called
' "PatchMatch".  You can read PatchMatch's original paper here (link good as of April 2022):
'
'https://gfx.cs.princeton.edu/pubs/Barnes_2009_PAR/patchmatch.pdf
'
'Note that both Resynthesizer and PatchMatch have "official" implementations by their original authors,
' but these implementations are license-incompatible with PhotoDemon.  Argh.
'
'So, I have taken the more challenging road and implemented my own version of Paul's "Resynthesizer"
' algorithm, as I understand it from his original paper.  While not as fast or potentially high-quality
' as GIMP's implementation (20 years of subsequent improvements are hard to beat), my version still works
' well enough to hopefully make it a reasonable take on Paul's thesis.  I definitely consider it good
' enough to warrant inclusion in PhotoDemon, and I'm particularly proud of the performance this version
' achieves, especially given VB's limitations.
'
'More than anything, though, I'm hoping that this BSD-licensed version of the algorithm draws attention
' from others who can help me improve it further, particularly performance-wise.  It seems a shame to
' let Paul's interesting algorithm languish in GIMP alone instead of being further developed by many
' different open-source developers.
'
'Broader integration of this feature into PhotoDemon itself is still a WIP, and some extended features
' available in GIMP (like "smart sharpening") are currently beyond my comprehension.  I will revisit in
' the future as time and/or my mental capacity allows.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'During debugging, it is helpful to report detailed process information.  Please disable this constant
' in production builds.
Private Const DEBUG_VERBOSE As Boolean = True

'The current inpainting mode.  Different modes require different behaviors throughout the inpainter;
' note also that some options and settings are only compatible with a single mode (e.g. edge-wrapping
' pixels is an essential part of seamless tile generation, but it makes no sense during mask inpainting).
Private Enum PD_InpaintingMode
    im_SeamlessTile = 0
    im_InpaintByMask = 1
End Enum

#If False Then
    Private Const im_SeamlessTile = 0, im_InpaintByMask = 1
#End If

Private m_CurMode As PD_InpaintingMode

'When comparing pixels against each other, it improves locality to store coordinates and colors
' next to each other.  Note that we need to store two sets of coordinates - a base coordinate
' (used for the original point that got "synthesized" into the current location) and an x/y offset
' from the target pixel (which is used in subsequent steps to test which pixel out of multiple
' candidates is the "best" one).
Private Type ip_Comparator
    cOffsets As PointLong
    cCoordinates As PointLong
    cColor As RGBQuad
End Type

'One of the first steps prior to inpainting is preparing a list of points that need to be inpainted.
' This list consists of coordinates into the destination image, whose order has been deliberately randomized.
' Note that points chosen early in the refinement process must be re-selected again later in the process,
' when more points have been filled in, so this list *will* contain duplicates, by design.  It may also be
' quite large compared to the original set of points, since the user can specify a quality parameter that
' affects how many refinement passes we apply.
Private m_PointsRemaining() As PointLong, m_numPointsLeft As Long

'When trying to match pixels, we need to randomly sample pixels from the source image as potential candidates.
' Because some source image pixels may be invalid (because e.g. they lie in the region the user wants removed),
' we can't simply grab *any* pixel - so during initialization, we'll build a list of valid coordinates to
' choose from.  (Note that some applications - like seamless tile generation - may use every source pixel.
' That's fine too!)
Private m_ValidSourcePoints() As PointLong, m_boundValidSourcePoints As Long

'We need an array (at the same dimensions as the destination image) that tracks the synthesis state of
' each destination pixel.  Because it is also necessary to track the current synthesized point location,
' we simply use a coordinate array, with unsynthesized pixels initialized to the magic value LONG_MAX.
Private m_SynthesisState() As PointLong

'A central tenet of the original Resynthesizer paper is using the Cauchy distribution when calculating
' pixel differences (instead of e.g. a more traditional gaussian distribution).  Cauchy is a much
' flatter distribution (https://en.wikipedia.org/wiki/Cauchy_distribution) which makes it more inclusive
' of outliers.  This is favorable for inpainting because we don't want *perfect* matches on each pixel -
' this will produce a noticeably artificial look.  Instead, we want lots of "good enough" matches which
' will hide seams more organically.
'
'Unfortunately, Cauchy requires an expensive Log() calculation (as used here) and it cannot be hard-coded
' because it relies on a critical input parameter "sigma" from the user - so we precalculate all possible
' pixel differences in a fixed-size table during initialization, then use that table for all color
' comparisons.  We also scale the default float-values by some arbitrary amount (currently 32768) to avoid
' the need for expensive floating-point math on each comparison.
'
'Note that the table is deliberately resized to an upper bound of 256, which we use as a special index to
' indicate "maximum difference" (which discourages further matching since the penalty is so large).
Private Const m_CauchyScale As Long = 32768
Private m_Cauchy(-255 To 256) As Long

'This term is actually the "sigma" value from the original paper (pp 43), but I'm fairly it's the analog
' of the GIMP plugin feature "sensitivity to outliers".  The range must be [0.0, 1.0], with a value of 0
' penalizing outliers most severely.  As the original paper says, "...so long as we are only concerned
' with finding the most likely match, not relative likelihoods, the Euclidean distance metric may be
' simulated in this metric by choosing a very large [sigma]."  The paper suggests a default value of
' 30 which makes no sense, so I assume it is meant to be "30 / 256".
Private m_AllowOutliers As Double

'Some fixed set of points around each pixel must be searched and compared to see if a newly placed
' pixel is a good match.  We can't do a naive raster-order search because we want to sort the pixels in
' ascending distance from the origin pixel.  As the original paper says (pp 40):  "...The best fit out of
' each of these locations is selected, based on a comparison of the pattern formed by the [n] nearest known
' neighbours of the current pixel to the pattern formed by pixels at corresponding offsets about each
' candidate location."
'
'The selection of [n] is generally left to the user as an input.  Smaller [n] values improve performance
' at some cost to matching quality.  The original paper does not provide much guidance on this point,
' but notes that you need to use some mix of both fixed points and random ones.  As it says on pp 40:
' "The locations examined are:
'  - Continuations of the regions in the input texture associated with the n nearest pixel values that
'    have already been chosen (i.e. as if the current pixel and one its neighbours formed part of a patch
'    copied exactly from the input texture).
'  - A further m random locations.
' The use of continuations allows production of good results even when a small number of locations are
'  examined in each individual search."
Private m_NearestOffsets() As PointLong, m_numNearestOffsets As Long

'From the offset list above, we need to generate [n] actual points to compare to the current point.
' These comparison point locations will vary according to input point, especially early in the inpainting
' process, because some points from the m_NearestOffsets() list won't have been synthesized yet by the
' algorithm. (Similarly, some points may lie entirelry outside image boundaries.)  Thus from the full set
' of *possible* comparison points, we must generate a *custom* list of target points for each pixel,
' consisting only of pixels that have already been synthesized.
Private m_Comparators() As ip_Comparator

'The user passes a parameter that tells us how many neighboring pixels to compare.  More can produce a
' better result, but incurs a corresponding performance penalty.  The original paper suggests a default
' value of 30.  (Note that boundary or selection issues - e.g. randomly selecting a pixel in the middle
' of a large mask, but the user has specified a very small default radius - may prevent us from reaching
' the full neighbor count, so we also track the *actual* neighbor count for the current candidate pixel.)
Private m_MaxNumNeighbors As Long, m_CurNumNeighbors As Long

'When evaluating neighboring pixels, we run the risk of evaluating the same pixel multiple times.
' This is especially true near boundaries in seamless tile mode, because wrapping ensures that
' each neighbor along a boundary appears in the neighbor list 2x - once in its actual position,
' then again as the result of mirroring a corresponding pixel mirrored across the boundary line
' (consider [-1, 0] and [1, 0], for example).  Seamless tile mode can thus be greatly accelerated
' by skipping pixels that have already been evaluated.  We do this by simply tracking the index of
' the last iteration that touched a given pixel.
Private m_LastChecked() As Long

'After comparing the nearest [n] candidates, we also want to sample some number of random pixels to
' see if any of those work better than the one we selected.  This step is especially important for early
' pixels, since they will not have *any* neighboring pixels available for comparisons (because the mask
' area has not been synthesized yet!).  This value represents a maximum upper bound for how many random
' pixels get compared.  More produces better results but with a corresponding performance penalty.
Private m_MaxRandomCandidates As Long

'In section 3.1.2 (pp 44), "Refining early-chosen pixel values", Paul makes the following observation:
' "Early-chosen pixel values are chosen on the basis of far-distant pixels, and may turn out to be
' inappropriate once nearer pixels have been filled in. Early chosen pixel values are also not chosen with
' the benefit of being able to continue the regions in the input texture associated with well-chosen nearby
' pixels. For these reasons, the algorithm re-chooses early chosen pixel values once later pixel values
' have been chosen."
'
'He then goes on to suggest a strategy of iteratively re-checking some fraction [p] of the image,
' where [p] is a value on the range [0, 1], and you re-check the first [p * n] pixels placed into the
' destination image in light of a larger collection of well-synthesized neighbors.  (The theory is that
' the final pixels you synthesize have the benefit of *tons* of well-placed neighbors, so they likely
' look good on their initial synthesis, but the first few pixels were effectively random and must be
' reevaluated after more neighbors exist.)
'
'The suggested default value in the paper is 0.75.
Private m_Refinement As Double

'Potential search radius when matching pixels.  In a perfect world, we'd search the whole image until we
' find enough pixels to reach our test threshold, but because we have to build a sorted list of offsets,
' this becomes problematic on large images.  Instead, we need to limit the potential search radius to some
' reasonable size around each pixel.
Private m_ComparatorRadius As Long

'If you want to create a tileable texture (instead of just inpainting), set this value to TRUE.  Note that
' this incurs different performance penalties than regular in-painting, so you may want to avoid it on
' large input images (as it may take an extremely long time to complete).  (Finding a way to prevent the user
' from abusing this is still TODO.)
Private m_TileMode As Boolean

'Dimensions and pixel array for the source image.  Note that this array *unsafely* aliases the source image,
' and *must* be manually unaliased before any class functions exit.
Private m_SrcWidth As Long, m_SrcHeight As Long
Private m_SrcPixels() As RGBQuad, m_SrcSA As SafeArray2D

'Dimensions and pixel array for the destination image.  Note that this array *unsafely* aliases the
' destination image, and *must* be manually unaliased before any class functions exit.
Private m_DstWidth As Long, m_DstHeight As Long
Private m_DstPixels() As RGBQuad, m_DstSA As SafeArray2D

'Point order needs to be semi-randomized for the algorithm to produce useful results.
' (This is a key element described by the original paper.)  We use PD's internal randomizer
' to accomplish this.
Private m_Random As pdRandomize

'The current best-match candidate (both its coordinates, and its "difference" from the test point,
' where a difference of 0 means "perfect match")
Private m_BestMatchPt As PointLong, m_BestMatchDiff As Long

'In inpainting mode, we need class-level access to the passed mask.  This class requires the mask array
' to be the same size as the source and destination images; this greatly simplifies bounds-checking,
' and provides a large performance boost relative to previous attempts (which took a separate mask rect)
' as a parameter.  The mask is just a byte array, and the synthesizer will treat all 0 values as
' "do not synthesize" and all non-zero values as "synthesize".  It does not perform any blending, by design -
' antialiasing or feathering must be handled by the caller after this class finishes its work.
Private m_MaskBytes() As Byte

'When applying the final version of the filter, progress-bar updates are important (because the function
' can be kinda slow, especially on huge on huge images).
Private m_ShowProgressUI As Boolean

'To improve pixel matching behavior, we need a list of neighboring pixel positions to choose from,
' sorted in ascending order from "nearest to [0, 0]".
'
'VB6 does not provide a native sort implementation, so I've written a quick-and-dirty quicksort
' for this class.  It uses a stack instead of recursion for improve performance.  (Seamless tile mode
' in particular has a large up-front initialization cost, so a fast sort algorithm is essential.)
Private Type QSStack
    sLB As Long
    sUB As Long
End Type

Private Type QSEntry
    ptDistance As Long
    idxOriginal As Long
End Type

'Use our version of the "Resynthesizer" algorithm for texture synthesis.
' All relevant parameters MUST BE SET PRIOR TO CALLING.
'
'Returns TRUE if successful, FALSE if the user cancels mid-way (or the supplied parameters are
' invalid, or the mask is empty on a content-aware fill, etc).
Private Function SynthesizeTexture() As Boolean
    
    SynthesizeTexture = False
    
    Const FUNC_NAME As String = "SynthesizeTexture"
    
    Dim startTime As Currency
    VBHacks.GetHighResTime startTime
    
    'Start by showing an arbitrary progress bar.  We won't have an *actual* upper-limit for progress reports
    ' until we finish initialization (because we don't know how many pixels need to be replaced) but we
    ' can guess based on mask boundaries.
    If m_ShowProgressUI Then
        ProgressBars.SetProgBarMax 10000    'Arbitrary non-zero value just to initialize the progress bar;
                                            ' we'll set a true value after we know how many pixels must be synthesized.
        ProgressBars.SetProgBarVal 1
        Message "Analyzing image..."
    End If
    
    'Initialization for this function is non-trivial (from a perf standpoint).
    ' TODO: minimize initialization steps if params haven't changed from a previous run?
    If (Not InitializeInpainter) Then
        InternalError FUNC_NAME, "initialization failed"
        Exit Function
    End If
    
    If DEBUG_VERBOSE Then PDDebug.LogAction "pdInpaint: initialization took " & VBHacks.GetTimeDiffNowAsString(startTime)
    PDDebug.LogAction m_numPointsLeft & " points need to be synthesized"
    VBHacks.GetHighResTime startTime
    
    Dim progCheck As Long
    If m_ShowProgressUI Then
        ProgressBars.SetProgBarMax m_numPointsLeft
        progCheck = ProgressBars.FindBestProgBarValue()
        Message "Repainting selected regions..."
    End If
    
    'Initialization prepared a list of all points in the destination image that need to be inpainted.
    ' We are now going to evaluate that list one-at-a-time until all pixels have been successfully
    ' "synthesized".
    Dim i As Long
    For i = 0 To m_numPointsLeft - 1
        
        Dim curPoint As PointLong, testPoint As PointLong, tmpPoint As PointLong
        curPoint = m_PointsRemaining(i)
        
        'Generate a list of nearby pixels that we can test to evaluate the "goodness" of the target pixel.
        ' We need to generate a custom list because...
        ' 1) Not all pixels have been synthesized yet (so many won't have a color), so there's no magic
        '    "fixed" list of valid neighbors that works for every pixel, and...
        ' 2) Even if this pixel has a lot of good neighbors, some pixels nearby may be out-of-bounds
        '    (so we don't want to use them as comparators).  Note that this criteria is only relevant
        '    for in-painting - for seamless tile generation, there are no boundaries because we just
        '    wrap OOB pixels.
        '
        'Anyway, this stage tries to build a list of valid neighboring pixel offsets *in the destination image*.
        ' That's important, because if we have neighbors that have already been synthesized, we'll use them.
        ' What we're trying to do is see if this "patch" of pixels (this one plus its neighbors) can be replaced
        ' by a more coherent "patch" from elsewhere in the image.  So first, we're going to figure out what pixels
        ' constitute this "patch", how good a "patch" this is, and whether there's a better one out there.
        '
        'At the end of the day, only one pixel will potentially be synthesized (the target one, curPoint),
        ' but how we determine what to replace it with depends on the list of coordinates we're about to
        ' assemble - the coordinates that constitute a novel "patch" on this specific iteration.
        ' (Each iteration generates a new patch of the nearest [n] valid pixels.)
        m_CurNumNeighbors = 0
        
        '(Note that in seamless tile mode, we can skip this step on the first pixel because it is 100% guaranteed
        ' to have no valid neighbors!)
        Dim runThisIteration As Boolean: runThisIteration = True
        If (m_CurMode = im_SeamlessTile) Then runThisIteration = (i > 0)
        
        If runThisIteration Then
            
            'Search as deep into the m_NearestOffsets() list (the sorted list of *potential* offsets) as we must
            ' until we assemble a full list of [m_maxNumNeighbors] valid pixels.  These pixels constitute the
            ' current "patch", which we will try to maximize the quality of.
            Dim j As Long
            For j = 0 To m_numNearestOffsets - 1
                
                'Comparison points are stored as offsets (not absolute points).  Thus we need to add offsets
                ' to the current point to produce the absolute coordinates of the point-to-test.
                testPoint.x = curPoint.x + m_NearestOffsets(j).x
                testPoint.y = curPoint.y + m_NearestOffsets(j).y
                
                'Before choosing this offset as a potential comparator, we must ensure the target pixel is:
                ' 1) in-bounds on the destination image
                '    (note: this is *not* relevant for seamless tile mode, because all boundaries get wrapped)
                ' 2) already synthesized!  (If it *hasn't* been synthesized yet we obviously don't want to use
                '    it as a comparator, because we can't evaluate its "fit quality" yet.)
                If IsPixelInBounds(testPoint) Then
                    
                    'Synthesis state is stored as a special flag value inside a separate tracking array.
                    ' Validate it before continuing.  (Note that the behavior of this settings varies by mode.
                    ' In seamless tile mode, it means the point has been synthesized already.  In content-aware
                    ' fill mode, it means either this point has been synthesized, *or* it lies outside the mask
                    ' and is fine to use as-is.)
                    tmpPoint = m_SynthesisState(testPoint.x, testPoint.y)
                    If (tmpPoint.x <> LONG_MAX) Then
                        
                        'This point has been synthesized (or in content-aware fill mode, it's not masked).
                        
                        'Store the offsets (relative to the target pixel) and current color of this pixel;
                        ' both will be used in the following step to produce a "goodness of fit" metric.
                        With m_Comparators(m_CurNumNeighbors)
                            .cOffsets = m_NearestOffsets(j)
                            .cColor = m_DstPixels(testPoint.x, testPoint.y)
                            
                            'We also want to know what this pixel's current *synthesized" coordinate is in the source
                            ' image (because we're going to compare its against neighbors in a later step).
                            .cCoordinates = tmpPoint
                        End With
                        
                        m_CurNumNeighbors = m_CurNumNeighbors + 1
                        
                        'We must never exceed the user's max neighbor value (because that's the upper limit for
                        ' all tracking arrays)
                        If (m_CurNumNeighbors >= m_MaxNumNeighbors) Then Exit For
                        
                    End If
                End If
            
            Next j
            
        '/i = 0
        End If
        
        'Reset the "best" distance to an impossibly large value.
        m_BestMatchDiff = LONG_MAX
        
        'Ensure we actually have at least one neighboring pixel to compare.  (For very large masks with small
        ' search radii, we may not, and that's fine - instead we'll just rely on the next step, sampling
        ' random pixels from the source image, to find a good first candidate for synthesizing this pixel.)
        If (m_CurNumNeighbors > 0) Then
            
            'A general rule-of-thumb with patch matching is that the best replacement for a given pixel is
            ' liekly a pixel that's relatively nearby in the source brush.  (This is the instinctive behavior
            ' for most users with a clone-brush tool, for example - use it to grab a nearby patch of pixels
            ' to overwrite the problematic one.)
            '
            'So rather than rely *solely* on randomly sampled pixels from the input image, let's first
            ' evaluate the nice list of neighboring pixels we've already assembled, because chances are that
            ' we can find a well-fitting pixel from that list!
            j = 0
            For j = 0 To m_CurNumNeighbors - 1
                
                'When we built our offset list, we cleverly grabbed the *synthesized* (or original, depending
                ' on mode) coordinate used by this offset pixel.  We now want to use that coordinate, along with
                ' the offset associated with this pixel, to choose a good pixel to sample.  To do so, we'll take
                ' the original coordinate of this pixel, and then subtract the current offset to see what pixel
                ' lies to [-offsetX, -offsetY] of this neighboring pixel.  Hypothetically, that should be an
                ' excellent replacement for the current pixel - right?
                With m_Comparators(j)
                    testPoint.x = .cCoordinates.x - .cOffsets.x
                    testPoint.y = .cCoordinates.y - .cOffsets.y
                End With
                
                'Just because the *source* point exists at this coordinate doesn't mean that our newly calculated
                ' test point does, so we need to repeat all bounds- and mask-checks from the previous step.
                '
                '(Again, note that VB does not short-circuit checks so we manually break each check out.)
                If (testPoint.x < 0) Then GoTo NextNeighbor
                If (testPoint.y < 0) Then GoTo NextNeighbor
                If (testPoint.x >= m_SrcWidth) Then GoTo NextNeighbor
                If (testPoint.y >= m_SrcHeight) Then GoTo NextNeighbor
                
                'In mask mode, we also want to ensure the target point is not within the mask.  (If it is,
                ' we don't want to consider it as a potential replacement!)
                If (m_CurMode = im_InpaintByMask) Then
                    If (m_MaskBytes(testPoint.x, testPoint.y) <> 0) Then GoTo NextNeighbor
                End If
                
                'If we've already compared this point this iteration, skip it.  This is a likely scenario as
                ' patches become well-established, since neighboring pixels are likely taken from the same
                ' "patch" in the original image.  (Also, EvaluatePixel, below, is highly expensive since we
                ' have to compare the target pixel to *each* entry in the comparator list - so skipping it is
                ' a huge perf win.)
                '
                'NOTE: this idea came from https://github.com/notwa/resynth which I was benchmarking against
                ' PD's implementation, and I could not for the life of me figure out why it kept crushing me
                ' in performance on texture synthesis.  Thank you to those authors for the clever optimization
                ' strategy (it may actually be adopted from earlier resynthesizer work, idk).
                If (m_LastChecked(testPoint.x, testPoint.y) <> i) Then
                    
                    'Evaluate the test pixel; if it's a better fit than the current one, EvaluatePixel will
                    ' update the class-level running sum accordingly.
                    EvaluatePixel testPoint
                    
                    'If this point is a perfect match against the current comparator list, stop searching because
                    ' we're not going to find anything better this time around.
                    If (m_BestMatchDiff = 0) Then Exit For
                    
                    'Flag this pixel as "checked during this pass" so we can skip it if we encounter it again.
                    m_LastChecked(testPoint.x, testPoint.y) = i
                    
                End If
                
NextNeighbor:
            Next j
        
        '/no valid neighbors
        End If
        
        'If we already have a "perfect" match, we don't need to compare random candidates for this pixel
        ' (because the current one is already as-good-as-it-gets).
        If (m_BestMatchDiff > 0) Then
            
            'We're now going to compare the best quality we've found so far to the quality of some randomly
            ' sampled points from the source image.  This is critical for early pixels because they don't
            ' have enough neighbors to determine whether they're a good match or not (so randomly selected
            ' pixels will likely perform better).
            '
            'As we get deeper into the pixel list, however, this step becomes less and less relevant
            ' because random pixels are unlikely to outperform pixels that have already survived previous
            ' iterations.  As such, we automatically scale-down the number of random candidates that we
            ' consider as we get deeper and deeper into the list.
            Dim numRandomCandidatesToTry As Long
            numRandomCandidatesToTry = m_MaxRandomCandidates
            
            'Scale down by how many points we have already processed, so that we're only doing 25%
            ' as many random comparisons by the end of the list (because those points are likely
            ' well-matched after multiple iterations).  This is an optimization of my own creation so
            ' please do not blame the original Resynthesis paper for this behavior!
            numRandomCandidatesToTry = numRandomCandidatesToTry * (0.25 + 0.75 * ((m_numPointsLeft - i) / m_numPointsLeft))
            
            'Iterate that many random pixels and see if any outperform our current "winning" choice
            For j = 0 To numRandomCandidatesToTry - 1
                
                'There is a very small (but non-zero!) chance that we'll have evaluated this random point already
                ' in the previous step.  This is particularly likely if the user set an aggressively small radius.
                ' We may as well do a quick check and skip this point if it's already been checked.
                testPoint = m_ValidSourcePoints(m_Random.GetRandomIntRange_WH(0, m_boundValidSourcePoints))
                If (m_LastChecked(testPoint.x, testPoint.y) <> i) Then
                    
                    EvaluatePixel testPoint
                    
                    'If this randomly selected pixel proved to be a "perfect" candidate (unlikely but
                    ' not impossible), stop searching.
                    If (m_BestMatchDiff = 0) Then Exit For
                    
                End If
                
            Next j
            
        End If

        'Assign the best-matching candidate to this destination pixel!
        m_DstPixels(curPoint.x, curPoint.y) = m_SrcPixels(m_BestMatchPt.x, m_BestMatchPt.y)
        
        'In our state array, note this pixel's current best-match coordinate.  (This coordinate may
        ' be replaced with a better alternative in subsequent passes, as more points get filled-in.)
        m_SynthesisState(curPoint.x, curPoint.y) = m_BestMatchPt
        
        'TODO: add cancellation support
        If m_ShowProgressUI And ((i And progCheck) = 0) Then ProgressBars.SetProgBarVal i
        'If Interface.UserPressedESC() Then Exit For
        
    Next i
    
    'If we raised the progress bar, hide it before exiting
    If m_ShowProgressUI Then ProgressBars.ReleaseProgressBar
    
    SynthesizeTexture = True
    
    If DEBUG_VERBOSE Then PDDebug.LogAction "Full resynthesis took " & VBHacks.GetTimeDiffNowAsString(startTime)
    
End Function

'Inpainting requires a large amount of prep work (and prep resources).
' Try to minimize initialization calls if you can.
Private Function InitializeInpainter() As Boolean
    
    Const FUNC_NAME As String = "InitializeInpainter"
    On Error GoTo InitializationFailed
    
    'Assume failure; if we reach the end without any fail states, we will reset this to TRUE
    InitializeInpainter = False
    
    'This function is profiled fairly aggressively, since it incurs a disproportionate performance penalty,
    ' especially when the source and/or destination image is large.  (This is especially true in "seamless tile" mode.)
    Dim startTime As Currency
    VBHacks.GetHighResTime startTime
    
    'Build an initial look-up table of the central Cauchy distribution.  This table is fixed according
    ' to differences in pixel values (which for PhotoDemon, are guaranteed to be 8-bit color components).
    ' Because this distribution requires the expensive Log() operator, we pre-build it.
    Dim i As Long
    If (m_AllowOutliers > 0#) Then
        
        'We really only need [-255, 255] values, but [256] is used as a special "*really* penalize this pixel"
        ' value (used as a default weight on pixels that lie outside the source image, which discourages
        ' over-selection of edge pixels in the input texture - a wanted behavior, as edges are harder to
        ' re-synthesize smoothly).
        Dim invAllowOutliers As Double
        invAllowOutliers = 1# / m_AllowOutliers
        
        For i = -255 To 256
            
            Dim tmpFloat As Double
            tmpFloat = GetCauchy((i / 256#) * invAllowOutliers) / GetCauchy(invAllowOutliers)
            
            'Scale by an arbitrary integer value for faster calculations on the inner loop
            tmpFloat = tmpFloat * m_CauchyScale
            m_Cauchy(i) = Int(tmpFloat + 0.5)
            
        Next i
    
    'If the user requires perfect matches (ugh) then *any* difference will receive max penalty.
    ' (This approach really isn't recommended and honestly, I may just disable this option entirely.)
    Else
        For i = -255 To 256
            m_Cauchy(i) = m_CauchyScale
        Next i
        m_Cauchy(i) = 0
    End If
    
    If DEBUG_VERBOSE Then PDDebug.LogAction "Initializing cauchy table took " & VBHacks.GetTimeDiffNowAsString(startTime)
    VBHacks.GetHighResTime startTime
    
    'Prepare a sorted list of offsets to search when evaluating pixel quality.  This step has a non-trivial
    ' startup cost due to the sort requirement.
    InitializeOffsetList
    If DEBUG_VERBOSE Then PDDebug.LogAction "Making and sorting offset list took " & VBHacks.GetTimeDiffNowAsString(startTime)
    VBHacks.GetHighResTime startTime
    
    'Generate a list of valid source points to sample *from*.  Points may be disqualified either implicitly
    ' (because they're inside the area-to-be-removed) or explicitly (maybe we could make a tool that allows
    ' the user to "paint" regions with either "use" or "dont use" flags?)
    '
    'Note that unlike the list of destination points, this one will never *exceed* the total number of source
    ' image coordinates, but it can be *smaller* than the total number of source image coordinates.
    ReDim m_ValidSourcePoints(0 To m_SrcWidth * m_SrcHeight - 1) As PointLong
    
    Dim x As Long, y As Long, idxDst As Long
    idxDst = 0
    
    'Potential places to draw sample points from varies by mode....
    ' 1) In seamless tile mode, we want to sample from the full source image.
    ' 2) In content-aware fill mode, the user can set a sample radius.  Do not exceed this.
    If (m_CurMode = im_SeamlessTile) Then
        
        For y = 0 To m_SrcHeight - 1
        For x = 0 To m_SrcWidth - 1
            m_ValidSourcePoints(idxDst).x = x
            m_ValidSourcePoints(idxDst).y = y
            idxDst = idxDst + 1
        Next x
        Next y
              
    ElseIf (m_CurMode = im_InpaintByMask) Then
        
        'Only add unmasked pixels to the "valid source sample" list
        For y = 0 To m_SrcHeight - 1
        For x = 0 To m_SrcWidth - 1
            If (m_MaskBytes(x, y) = 0) Then
                m_ValidSourcePoints(idxDst).x = x
                m_ValidSourcePoints(idxDst).y = y
                idxDst = idxDst + 1
            End If
        Next x
        Next y
        
    End If
    
    m_boundValidSourcePoints = idxDst - 1
    If (UBound(m_ValidSourcePoints) <> m_boundValidSourcePoints) Then ReDim Preserve m_ValidSourcePoints(0 To m_boundValidSourcePoints) As PointLong
    If DEBUG_VERBOSE Then PDDebug.LogAction "Initializing source point list took " & VBHacks.GetTimeDiffNowAsString(startTime)
    VBHacks.GetHighResTime startTime
    
    'The list of "comparators" (neighboring pixels that we try to evaluate on every synthesis)
    ' has a fixed size specified by the user.  Note that we may not perform this many comparisons
    ' for *every* pixel, but this gives us a safe upper limit so we can ignore array bound checks
    ' on perf-sensitive inner loops.
    ReDim m_Comparators(0 To m_MaxNumNeighbors - 1) As ip_Comparator
    
    'In the destination image, we need to track the synthesis state of each pixel-to-be-filled.
    ' (Pixels that have not yet been synthesized must be flagged, so we know to ignore them.)
    ReDim m_SynthesisState(0 To m_DstWidth - 1, 0 To m_DstHeight - 1) As PointLong
    
    'Allocate the base set of points that need to be synthesizsed.  Note that this list will ultimately be
    ' larger than the set of incoming points, potentially *much* larger, because pixels selected early in
    ' the process must be revisited later in the process, when more neighboring points have been filled.
    ' (Because of this, we deliberately over-size this list to start.)
    ReDim m_PointsRemaining(0 To m_DstWidth * m_DstHeight * 4 - 1) As PointLong
    
    idxDst = 0
    For y = 0 To m_DstHeight - 1
    For x = 0 To m_DstWidth - 1
        
        'While we're here, dump this destination point to its storage queue too
        ' (dependent on current inpainting mode)
        Select Case m_CurMode
            
            'In seamless tile mode, *all* source pixels are valid and available
            Case im_SeamlessTile
                
                'Add *all* destination pixels to the "need to be synthesized "list
                m_SynthesisState(x, y).x = LONG_MAX
                m_PointsRemaining(idxDst).x = x
                m_PointsRemaining(idxDst).y = y
                idxDst = idxDst + 1
            
            'In in-painting mode, we must compare coordinates to the mask to see if they belong
            ' in the "need-to-synthesize" point list.
            Case im_InpaintByMask
                
                'Masked pixels must be flagged as "not yet synthesized" and explicitly added to the
                ' "need-to-synthesize" list.
                If (m_MaskBytes(x, y) <> 0) Then
                    m_SynthesisState(x, y).x = LONG_MAX
                    m_PointsRemaining(idxDst).x = x
                    m_PointsRemaining(idxDst).y = y
                    idxDst = idxDst + 1
                
                'Mark as "already synthesized" so the synthesizer loop knows to ignore this pixel.
                ' (And DO NOT add this coordinate to the "need-to-synthesize" list.)
                Else
                    m_SynthesisState(x, y).x = x
                    m_SynthesisState(x, y).y = y
                End If
                
        End Select
        
    Next x
    Next y
    
    'Because we deliberately over-sized the destination list, we need to flag how many points *actually*
    ' exist in the array
    m_numPointsLeft = idxDst
    If (m_numPointsLeft = 0) Then
        InternalError FUNC_NAME, "no points to synthesize!"
        Exit Function
    End If
    
    If DEBUG_VERBOSE Then PDDebug.LogAction "Initializing destination point list (" & m_numPointsLeft & " points) took " & VBHacks.GetTimeDiffNowAsString(startTime)
    VBHacks.GetHighResTime startTime
    
    'Next, we need to randomize the order of the initial list of "points to synthesize".
    ' Shuffle using standard Fisher-Yates.
    Dim j As Long, tmpPoint As PointLong
    
    Dim uBoundPoints As Long
    uBoundPoints = m_numPointsLeft - 1
    For i = uBoundPoints To 0 Step -1
        
        'Swap the current point [i] with some random point [j] preceding it in the list.
        j = m_Random.GetRandomIntRange_WH(0, i)
        tmpPoint = m_PointsRemaining(i)
        m_PointsRemaining(i) = m_PointsRemaining(j)
        m_PointsRemaining(j) = tmpPoint
        
    Next i
    
    If DEBUG_VERBOSE Then PDDebug.LogAction "Shuffling initial set of data points took " & VBHacks.GetTimeDiffNowAsString(startTime)
    VBHacks.GetHighResTime startTime
    
    'Now that we have generated a queue of pixels that need to be synthesized, we need to re-add the first
    ' [p * n] pixels from the list to the queue.  A longer explanation for this is given at the top of the
    ' module (see the definition of m_Refinement), but basically pixels chosen early in the placement
    ' process do not have the benefit of well-established neighbors, so they are more likely to be random
    ' crap relative to later pixels.  By re-evaluating these pixels after more neighbors get synthesized,
    ' we are much more likely to produce a high-quality output.
    If (m_Refinement > 0#) Then
        
        'The number of points to add will be repeatedly reduced by the user's refinement parameter,
        ' but it begins as the full size of the incoming table.  (This exact approach is recommended
        ' by the original paper.)
        Dim numPointsToRefine As Long
        numPointsToRefine = m_numPointsLeft
        idxDst = numPointsToRefine
        
        'If the user's refinement parameter is very high (e.g. 0.99) and the area to synthesize is
        ' very big, it's entirely possible to run out of memory because we append too many data
        ' points to the list.  As such, we need a good failsafe mechanism for ensuring we don't
        ' add *too* many points to our processing list (especially because each point takes 8 bytes).
        '
        'I don't have a magic bullet for this process, but as an easy initial attempt, let's just
        ' limit it to 5x the original number of points.
        Dim maxNumPoints As Long
        maxNumPoints = numPointsToRefine * 5
        
        'We will keep reducing [n] by the user-supplied value "m_Refinement", which is a fraction on the
        ' range [0.0, < 1.0].
        Do While (numPointsToRefine > 1)
            
            'NOTE: I have strongly debated whether to apply the reduction parameter here, or *after* adding
            ' a full duplication of the initial point list.  (The paper isn't really clear on this point.)
            ' My own testing showed that doing two full passes over the input data produces a subjectively
            ' higher-quality output, but there is obviously a performance cost to this.  I think the
            ' trade-offs are worth it, and I think it's reasonable to *always* do two full passes over
            ' the target area (regardless of the user's refinement parameter).
            '
            'Thus, we will always re-add all destination points to the "points-to-be-processed" list,
            ' and then we will add e.g. 75% of the pixels back for a 3rd pass (if the user's m_Refinement
            ' parameter is 0.75), then 75% of *those* pixels back again, then 75% of each preceding group
            ' over and over until we either reach 0 points added, or we exceed the maximum number of points
            ' declared above.
            
            'Bulk copy this set of points into place (after ensuring sufficient storage space, obviously)
            If ((idxDst + numPointsToRefine) > UBound(m_PointsRemaining)) Then ReDim Preserve m_PointsRemaining(0 To UBound(m_PointsRemaining) * 2 + 1) As PointLong
            VBHacks.CopyMemoryStrict VarPtr(m_PointsRemaining(idxDst)), VarPtr(m_PointsRemaining(0)), numPointsToRefine * 8     '8 = LenB(PointLong)
            idxDst = idxDst + numPointsToRefine
            
            'Stop adding points if we've hit or surpassed the pre-calculated maximum number of points.
            If (idxDst >= maxNumPoints) Then Exit Do
            
            'Calculate how many points to add on the next iteration, and limit it to the max amount
            ' we determined earlier.
            numPointsToRefine = Int(numPointsToRefine * m_Refinement)
            If (idxDst + numPointsToRefine > maxNumPoints) Then numPointsToRefine = (maxNumPoints - idxDst)
            
        Loop
        
        'After all points have been added, resize the (potentially way too big) table to its precise size
        ReDim Preserve m_PointsRemaining(0 To idxDst - 1) As PointLong
        m_numPointsLeft = idxDst
        
    End If
    
    If DEBUG_VERBOSE Then PDDebug.LogAction "Appending repeat data points (" & m_numPointsLeft & " points) took " & VBHacks.GetTimeDiffNowAsString(startTime)
    VBHacks.GetHighResTime startTime
    
    'See the top of this module for details, but basically, in seamless tile mode, multiple neighboring pixels
    ' will map to the same position in the image due to edge-wrapping (e.g. [-1, 0] and [1, 0] will both map
    ' to [1, 0]).  We can save a lot of time by not checking duplicate values like these.
    '
    'To achieve this, we track the index of the last iteration that touched a given pixel, and if we see a
    ' repeat we skip subsequent checks.
    ReDim m_LastChecked(0 To m_SrcWidth - 1, 0 To m_SrcHeight - 1) As Long
    VBHacks.FillMemory VarPtr(m_LastChecked(0, 0)), m_SrcWidth * m_SrcHeight * 4, 255
    
    If DEBUG_VERBOSE Then PDDebug.LogAction "Initializing attempts array took " & VBHacks.GetTimeDiffNowAsString(startTime)
    VBHacks.GetHighResTime startTime
    
    InitializeInpainter = True
    Exit Function

InitializationFailed:
    InternalError FUNC_NAME, "internal error", Err.Number
    
End Function

'Compare the pixel at a given coordinate to all pixels in the current Comparator list.  If this point is
' better than all previous ones we tested, update the module-level m_BestMatch value(s) to match.
Private Sub EvaluatePixel(ByRef srcPoint As PointLong)
    
    Dim totalDiff As Long, curDiff As Long
    
    Dim r As Long, g As Long, b As Long, a As Long
    Dim x As Long, y As Long, numPixelsCompared As Long
    
    Dim i As Long
    For i = 0 To m_CurNumNeighbors - 1
        
        'The comparator list stores offsets (not absolute values); add those offsets to the passed point
        ' to produce an absolute coordinate into the source image
        With m_Comparators(i).cOffsets
            x = srcPoint.x + .x
            y = srcPoint.y + .y
        End With
        
        'Make sure the target point actually lies within the source image's boundaries.  (Note that this is
        ' constructed weirdly because VB6 does not short-circuit.)
        If (x >= 0) And (y >= 0) Then
            If (x < m_SrcWidth) And (y < m_SrcHeight) Then
                
                'In masked mode, we also need to ensure this pixel is *not* part of the mask.
                ' (We're indexing into the source image, and selected pixels will *still retain their
                ' original contents*.)
                Dim okToUsePixel As Boolean: okToUsePixel = True
                If (m_CurMode = im_InpaintByMask) Then okToUsePixel = (m_MaskBytes(x, y) = 0)
                
                If okToUsePixel Then
                    
                    'Use the pre-constructed Cauchy table to calculate distance
                    Dim origColor As RGBQuad, candidateColor As RGBQuad
                    origColor = m_SrcPixels(x, y)
                    candidateColor = m_Comparators(i).cColor
                    
                    'Swap chars into ints to prevent VB from overflowing on char arithmetic
                    b = candidateColor.Blue
                    g = candidateColor.Green
                    r = candidateColor.Red
                    a = candidateColor.Alpha
                    curDiff = m_Cauchy(b - origColor.Blue) + m_Cauchy(g - origColor.Green) + m_Cauchy(r - origColor.Red) + m_Cauchy(a - origColor.Alpha)
                    
                Else
                    If (numPixelsCompared <> 0) Then
                        curDiff = totalDiff \ numPixelsCompared
                    Else
                        curDiff = m_Cauchy(256) * 4
                    End If
                End If
                
            'If this pixel lies outside the original image's boundaries, assign it a fixed-size penalty.
            ' (We could also ignore the pixel, but this creates unwanted favoritism for edge pixels because
            ' they simply won't have as many neighboring pixels to compare - so this is a better compromise
            ' for final quality.)
            Else
                curDiff = m_Cauchy(256) * 4
            End If
        Else
            curDiff = m_Cauchy(256) * 4
        End If
        
        'Tally the running sum of differences, and if we exceed the current best candidate bail immediately
        totalDiff = totalDiff + curDiff
        If (totalDiff >= m_BestMatchDiff) Then Exit Sub
        numPixelsCompared = numPixelsCompared + 1

    Next i
    
    'If we completed a full scan, the current point looks better than previous ones.  Store it!
    m_BestMatchDiff = totalDiff
    m_BestMatchPt = srcPoint
    
End Sub

'The Cauchy distribution is a critical aspect of the original Resynthesizer paper (see pp 43)
' because unlike a Gaussian distribution, it is much more forgiving to a few outliers on each pass.
'
'Because calculating this requires a slow Log function, we pre-calculate all values and store them in
' a fixed-size lookup table.  (Note that incoming x may be 0, so we must add 1 to avoid an error.)
Private Function GetCauchy(ByVal x As Double) As Double
    GetCauchy = Log(x * x + 1#)
End Function

'Prepare a fixed list of offsets for searching around each resynthesized pixel.  Search radius must be
' constrained by one of two things:
' 1) For inpainting, some radius specified by the user (smaller radii are faster, as you'd expect)
' 2) For novel texture synthesis, constrain by the minimum of source and destination dimensions.
'    (Offsets will be used as indices into either, and it's easier to just use min values here.)
Private Sub InitializeOffsetList()
    
    'In texture synthesis mode (and as a failsafe during inpainting), constrain by the smaller of the
    ' source or destination's height.
    Dim minWidth As Long, minHeight As Long
    minWidth = PDMath.Min2Int(m_SrcWidth, m_DstWidth)
    minHeight = PDMath.Min2Int(m_SrcHeight, m_DstHeight)
    
    'In in-painting mode, we want to further limit the search radius by a user-supplied parameter
    If (m_CurMode = im_InpaintByMask) Then
        minWidth = PDMath.Min2Int(minWidth, m_ComparatorRadius)
        minHeight = PDMath.Min2Int(minHeight, m_ComparatorRadius)
    End If
    
    'Ensure available space for the offset list (x4 is needed because there are four quadrants of offsets)
    ReDim m_NearestOffsets(0 To minWidth * minHeight * 4 - 1) As PointLong
    
    'Convert minimum values to fixed bounds for the forthcoming loops
    minHeight = minHeight - 1
    minWidth = minWidth - 1
    
    Dim idxPoint As Long: idxPoint = 0
    
    'TODO: see if there is some algorithm for adding points in an already sorted way?
    Dim x As Long, y As Long
    For y = -minHeight To minHeight
    For x = 0 To minWidth
        
        'Don't add (0, 0); we don't want to compare points to themselves!
        If (x = 0) And (y = 0) Then GoTo SkipPoint
        
        'Add this point to the list, and its horizontal mirror.  (This improves sort perf by keeping
        ' similarly-distanced-points close together in the initial unsorted list.)
        m_NearestOffsets(idxPoint).x = x
        m_NearestOffsets(idxPoint).y = y
        idxPoint = idxPoint + 1
        
        If (x > 0) Then
            m_NearestOffsets(idxPoint).x = -x
            m_NearestOffsets(idxPoint).y = y
            idxPoint = idxPoint + 1
        End If
        
SkipPoint:
    Next x
    Next y
    
    'Size the list precisely
    m_numNearestOffsets = idxPoint
    If (UBound(m_NearestOffsets) <> m_numNearestOffsets - 1) Then ReDim Preserve m_NearestOffsets(0 To m_numNearestOffsets - 1) As PointLong
    
    'To ensure accurate matching, we need to sort this offset list by its proximity to (0, 0).
    ' (We only search [n] points (user-supplied parameter) around each pixel, but we disqualify pixels that
    ' have not yet been filled.  Because of this, on many pixels we will only use a small subset of coordinates
    ' from the start of the list, but for tasks like novel texture synthesis we will actually use the *full* list,
    ' especially at the start of the generation process, so a full, valid list needs to exist.)
    SortCoordinates
    
End Sub

'Returns TRUE if a given pixel is in-bounds (per the DESTINATION image), FALSE otherwise.
' (In tileable texture mode, this function ALWAYS returns true, by design - because any pixel coordinate
' gets "wrapped" into a valid position.)
Private Function IsPixelInBounds(ByRef srcPoint As PointLong) As Boolean
    
    IsPixelInBounds = False
    
    'We only wrap pixels when generating tileable textures.  (In regular in-painting mode, this produces
    ' weird results - while you could clamp pixels to boundary row/columns, this produces over-representation
    ' of edge pixels.  Subjectively better results are produced by ignoring OOB pixels.)
    If m_TileMode Then
        
        If (srcPoint.x < 0) Then
            srcPoint.x = ModuloInt(srcPoint.x, m_DstWidth)
        ElseIf (srcPoint.x >= m_DstWidth) Then
            srcPoint.x = ModuloInt(srcPoint.x, m_DstWidth)
        End If
        
        If (srcPoint.y < 0) Then
            srcPoint.y = ModuloInt(srcPoint.y, m_DstHeight)
        ElseIf (srcPoint.y >= m_DstHeight) Then
            srcPoint.y = ModuloInt(srcPoint.y, m_DstHeight)
        End If
    
    'If we're *not* in tiling mode, exit immediately on any boundary failures
    Else
        If (srcPoint.x < 0) Then Exit Function
        If (srcPoint.x >= m_DstWidth) Then Exit Function
        If (srcPoint.y < 0) Then Exit Function
        If (srcPoint.y >= m_DstHeight) Then Exit Function
    End If
    
    IsPixelInBounds = True
    
End Function

'This is a modified modulo function; it handles negative values differently from VB's built-in Mod
Private Function ModuloInt(ByVal quotient As Long, ByVal divisor As Long) As Long
    ModuloInt = quotient - Fix(quotient / divisor) * divisor
    If (ModuloInt < 0&) Then ModuloInt = ModuloInt + divisor
End Function

Private Sub SortCoordinates()
    
    If (m_numNearestOffsets <= 0) Then Exit Sub
    
    'To improve performance, we're not gonna sort the coordinate list in-place.
    ' Instead, we're going to pre-calculate all distances, sort just the distance list,
    ' then rebuild the original coordinate array accordingly.
    Dim sortList() As QSEntry
    ReDim sortList(0 To m_numNearestOffsets - 1) As QSEntry
    
    Dim i As Long, j As Long
    For i = 0 To m_numNearestOffsets - 1
        sortList(i).idxOriginal = i
        With m_NearestOffsets(i)
            'Potential neighbor offsets are sorted by their Euclidean distance from the origin.
            ' (We only need relative values, so don't waste energy on a square-root.)
            sortList(i).ptDistance = .x * .x + .y * .y
        End With
    Next i
    
    'Prep our internal stack (this is faster and more reliable than recursion).  Note that because
    ' we are using a fixed potential list of sort candidates, this upper limit is guaranteed safe
    ' for this application - but it may *not* be safe for all quick sort ops, so if you change things
    ' like the maximum upper limit for search radius, you may need to increase stack size accordingly.
    Const INIT_QUICKSORT_STACK_SIZE As Long = 256
    Dim qsRemaining() As QSStack, qsStackPtr As Long
    ReDim qsRemaining(0 To INIT_QUICKSORT_STACK_SIZE - 1) As QSStack
    qsStackPtr = 0
    qsRemaining(0).sLB = 0
    qsRemaining(0).sUB = m_numNearestOffsets - 1
    
    'Like most quick-sort implementations, VB6 sees a perf boost when switching over to insertion sort
    ' when a quicksort sub-list is small.  This threshold could probably be refined after more
    ' intensive study, but the usual default of 10 shows a ~20-30% perf improvement over baseline.
    Const USE_INSERTION_SORT_THRESHOLD As Long = 10
    
    'Execute the sort (just a bog-standard quick sort at present)
    Dim idxLow As Long, idxHigh As Long
    Dim v As QSEntry, vDist As Long
    
    Do
        
        'Load the next set of boundaries, and reset all pivots
        idxLow = qsRemaining(qsStackPtr).sLB
        idxHigh = qsRemaining(qsStackPtr).sUB
        
        'Switch over to insertion sort when the sub-list is small
        If (idxHigh - idxLow < USE_INSERTION_SORT_THRESHOLD) Then
            i = idxLow + 1
            Do While (i <= idxHigh)
                v = sortList(i)
                vDist = v.ptDistance
                j = i - 1
                Do While (j >= idxLow) And (sortList(j).ptDistance > vDist)
                    sortList(j + 1) = sortList(j)
                    j = j - 1
                Loop
                sortList(j + 1) = v
                i = i + 1
            Loop
        
        'On larger lists, use standard quicksort
        Else
            
            i = idxLow
            j = idxHigh
            vDist = sortList((i + j) \ 2).ptDistance
            
            Do
                Do While (sortList(i).ptDistance < vDist)
                    i = i + 1
                Loop
                Do While (sortList(j).ptDistance > vDist)
                    j = j - 1
                Loop
                
                'Swap final indices as necessary
                If (i <= j) Then
                    v = sortList(i)
                    sortList(i) = sortList(j)
                    sortList(j) = v
                    i = i + 1
                    j = j - 1
                End If
                
            Loop Until i > j
            
            'Conditionally add new entries to the processing stack
            If (idxLow < j) Then
                qsRemaining(qsStackPtr).sLB = idxLow
                qsRemaining(qsStackPtr).sUB = j
                qsStackPtr = qsStackPtr + 1
            End If
            
            If (i < idxHigh) Then
                qsRemaining(qsStackPtr).sLB = i
                qsRemaining(qsStackPtr).sUB = idxHigh
                qsStackPtr = qsStackPtr + 1
            End If
            
        End If
        
        'Decrement the stack pointer and continue as long as new sub-lists were added
        qsStackPtr = qsStackPtr - 1
        
    Loop While (qsStackPtr >= 0)
    
    'The coordinate list has been successfully sorted.  Clone the original list, then copy each entry
    ' from that temporary clone into its final position in the "official" list.
    Dim tmpPoints() As PointLong
    ReDim tmpPoints(0 To m_numNearestOffsets - 1) As PointLong
    CopyMemoryStrict VarPtr(tmpPoints(0)), VarPtr(m_NearestOffsets(0)), 8 * m_numNearestOffsets
    
    For i = 0 To m_numNearestOffsets - 1
        m_NearestOffsets(i) = tmpPoints(sortList(i).idxOriginal)
        'During development, it was helpful to verify sort correctness:
        'If (i > 0) Then
        '    If (m_NearestOffsets(i).x * m_NearestOffsets(i).x + m_NearestOffsets(i).y * m_NearestOffsets(i).y < m_NearestOffsets(i - 1).x * m_NearestOffsets(i - 1).x + m_NearestOffsets(i - 1).y * m_NearestOffsets(i - 1).y) Then Debug.Print "BAD SORT AT " & i
        'End If
    Next i
    
End Sub

'From an arbitrary source image, create a seamlessly tiling new image
Friend Sub MakeSeamlessTile(ByRef srcDIB As pdDIB, ByRef dstDIB As pdDIB)
        
    'Set the class-wide mode indicator.  This will toggle a number of different behaviors during
    ' different inpainting steps.
    m_CurMode = im_SeamlessTile
        
    'The process for inpainting and generating tileable textures is 95+% similar.  One of the only differences
    ' when generating tileable textures is that you need to scan the whole input image, and *wrap* pixels when
    ' scanning for neighbor matches (so that the resulting tile has cleanly tiling boundaries).
    m_TileMode = True
    
    'Maximum neighbor count must ultimately be user-provided; the original paper default is 30.
    ' (Note that these are the 30 *nearest* pixels to the target one.  Pixels that have not yet been filled
    ' or are outside the image (in non-tiling mode) are not considered.  Randomly selected pixels will also
    ' be evaluated to see if *they* are better matches; that count is specified separately.
    m_MaxNumNeighbors = 30
    
    'After searching the [n] nearest pixels to see how well the current candidate matches, we next want to
    ' sample some random subset of pixels from the image to see if they do any better than the current
    ' candidate.  The original paper suggests a default value of 80.
    m_MaxRandomCandidates = 80
    
    'Refinement is the fraction of pixels that we reevaluate after generating the output image.
    ' Must be a positive value on the range [0, <1] (1.0 does *NOT* work - it must be less than 1!)
    '
    'We currently assume the same default as the original paper: 75%
    m_Refinement = 0.75
    
    'The original paper suggests a sigma of "30" but this makes no sense, so I assume 30 is just an input
    ' that is supposed to be scaled by 256 or 100?  (256 appears to produce better results here.)
    m_AllowOutliers = 30# / 256#
    
    'Prep input image
    m_SrcWidth = srcDIB.GetDIBWidth
    m_SrcHeight = srcDIB.GetDIBHeight
    srcDIB.WrapRGBQuadArrayAroundDIB m_SrcPixels, m_SrcSA
    
    'Prep output image
    If (dstDIB Is Nothing) Then Set dstDIB = New pdDIB
    m_DstWidth = srcDIB.GetDIBWidth
    m_DstHeight = srcDIB.GetDIBHeight
    dstDIB.CreateBlank m_DstWidth, m_DstHeight, 32, 0, 0
    
    'Because no new pixel values will be created (we will only be "rearranging" existing pixels),
    ' alpha premultiplication state matches the input image.
    dstDIB.SetInitialAlphaPremultiplicationState srcDIB.GetAlphaPremultiplication
    dstDIB.WrapRGBQuadArrayAroundDIB m_DstPixels, m_DstSA
    
    'Perform filter
    SynthesizeTexture
    
    'Unalias all unsafe array wrappers
    srcDIB.UnwrapRGBQuadArrayFromDIB m_SrcPixels
    dstDIB.UnwrapRGBQuadArrayFromDIB m_DstPixels
    
End Sub

'At initialization, all parameters will be set to their default values.  You can change any of these values
' prior to calling the main SynthesizeTexture() function.
Friend Sub SetDefaultParameterValues()
    
    'How many neighbors to compare when evaluating the quality of a synthesized pixel.  30 is the default value
    ' given in the original resynthesizer paper.  It's probably overkill for most applications.
    '
    'Because neighbors will always be searched in nearest-to-furthest order, changing this doesn't change the
    ' quality of pixel matching, per se - it just controls how far outward the algorithm evaluates during the
    ' EvaluatePixel process.
    m_MaxNumNeighbors = 25
    
    'A key part of the "resynthesize" algorithm is testing a set of random pixels against each potential
    ' candidate.  This is unavoidable seeing as 1) comparing every pixel in the source image would take years,
    ' and 2) testing replacement pixels in an orderly fashion (by scanline, for example) produces noticeable
    ' banding in the final result.  80 random candidates is the suggested default in the original
    ' resynthesizer paper.
    '
    'Note that my implementation uses this maximum value early in the resynthesize process, but scales it down
    ' as better matches are found.  This is an easy optimization because as we near the end of the synthesize
    ' process, existing pixels have already survived multiple iterations and are less and less likely to be
    ' "outperformed" by a random pixel from elsewhere in the image.
    m_MaxRandomCandidates = 80
    
    'After synthesizing each pixel at least once, we then re-refine the earliest synthesized pixels to try
    ' and improve their match quality against their neighbors (who didn't exist yet on the first pass).
    ' This refinement is thus a percentage value, where e.g. 0.75 means "re-test the first 75% of synthesized
    ' pixels AGAIN after all pixels have been synthesized".
    '
    'The default value in the original synthesize paper is 0.75.  Note that PhotoDemon, by design, synthesizes
    ' all points TWICE before this value kicks in.  I do this because repeat testing shows a significant
    ' (subjective) quality improvement after revisiting each pixel at least once.  Thus if you have 10,000 pixels
    ' to synthesize, PhotoDemon will synthesize all 10,000 once, then re-synthesize them AGAIN, and then
    ' resynthesize the first e.g. 75% of pixels a third time.
    '
    'This value is used repeatedly, so after e.g. 75% of pixels have been re-synthesized a third time, the first
    ' 75% of *those* 75% of pixels (e.g. 56.25% of the full original set) are then re-synthesized a fourth time.
    ' This continues until we reach "0% of pixels", so the total number of pixels synthesized will be much
    ' larger than the actual size of the mask.  This is the only way to ensure a high-quality output.
    m_Refinement = 0.75
    
    'The original paper calls this value "sigma", and I'm fairly certain GIMP's Resynthesize plugin renames it
    ' "sensitivity to outliers".  This value controls the spread of the Cauchy distribution look-up table we
    ' generate and unlike previous parameters, it has no performance implications.  (All possible values,
    ' from 0.0 to 1.0, have the same performance.)
    '
    'Instead, this value determines how "good" of a match a synthesized pixel needs to be to its neighbors
    ' before we decide to keep it.  A value of 0 means that the new pixel must match its original patch
    ' neighbors identically.  This is not really desirable because the filled patch starts to look like an
    ' obvious copy of sub-patches from elsewhere in the image.
    '
    'Conversely, a value of 1.0 means that random patches are totally fine.  This leads to areas of obviously
    ' mismatched pixels alongside very good matches (because random sampling still provides some nice fits)
    ' and again, usually isn't desirable.
    '
    'It's hard to provide a "one-size fits all" value here, however, because it depends a lot on the current
    ' image.  For example, synthesizing something very noisy - like a top-down view of beach sand - can
    ' actually look great with a max value of 1.0, because the underlying texture is basically noise.
    ' Similarly, a highly repeatable geometric texture (like tiles or bricks, perhaps) may do better with
    ' a value closer to 0.
    '
    'The default value in the original resynthesizer paper is 30 / 256 (~0.12).
    m_AllowOutliers = 30# / 256#
    
    'Because the original paper is primarily focused on full texture synthesis (where you create a novel
    ' seamless texture from some source), it doesn't deal a lot with specifics for content-aware filling.
    ' Instead, it shows some examples and sort of says "the same ideas work for this concept, too!"
    '
    'One of the things that quickly became apparent during my own implementation attempt was that you
    ' don't want content-aware fill pulling pixels from too far away in the image.  Otherwise, you may end
    ' up mistakenly replacing something like sky with something like water - because they're *close enough*
    ' in color to confuse the algorithm.
    '
    'So I added an option to restrict the radius where the synthesizer looks for replacement pixels.
    ' Note that this value is based on the size of the underlying mask, so if you have three tiny fill
    ' areas spread across an image, it is the union rect of *all* those areas that is used for the
    ' underlying radius calculation.  (Basically, any pixel within 300 pixels of the boundary rect is OK
    ' to use.)  As such, you're better off replacing individual objects one-at-a-time rather than selecting
    ' them all and then batch-replacing them.
    m_ComparatorRadius = 300
    
End Sub

Private Sub Class_Initialize()
    
    'Initialize a random number generator
    Set m_Random = New pdRandomize
    
    'At present, PhotoDemon seeds the inpainter randomly.  You could also allow the user to provide their
    ' own random seed, which would ensure identical results every time the function is called (with the
    ' same inputs and mask, obviously).
    m_Random.SetSeed_AutomaticAndRandom
    
    'Set all parameters to their default values
    SetDefaultParameterValues
    
End Sub

Private Sub InternalError(ByRef funcName As String, ByRef errComment As String, Optional ByVal errNumber As Long = 0&)
    If (errNumber <> 0) Then
        PDDebug.LogAction "WARNING!  pdInpaint." & funcName & " VB error (#" & Err.Number & "): " & Err.Description & " || " & errComment
    Else
        PDDebug.LogAction "WARNING!  pdInpaint." & funcName & " reported: " & errComment
    End If
End Sub
