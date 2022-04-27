VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdInpaint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Inpainting (Texture Synthesis) class
'Copyright 2022-2022 by Tanner Helland
'Created: 21/April/22
'Last updated: 26/April/22
'Last update: wrap up initial build (seamless tile mode only, since it's so much simpler conceptually!)
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
'He then goes on to suggest a strategy of iteratively re-checking some fraction [p] of the image, where [p]
' is a value on the range [0, 1], and you re-check the first [p * n] pixels placed into the destination image
' in light of a larger collection of synthesized neighbors since the last iteration.  (The theory is that the
' final pixels you synthesize have the benefit of *tons* of well-placed neighbors, so they are likely to look
' good with minimal further iteration, but the first few pixels were effectively random and likely need to be
' reevaluated again.)
'
'The suggested default value in the paper is 0.75.
Private m_Refinement As Double

'If you want to create a tileable texture (instead of just inpainting), set this value to TRUE.  Note that
' this incurs different performance penalties than regular in-painting, so you may want to avoid it on
' large input images (as it may take an extremely long time to complete).  Finding a way to prevent the user
' from abusing this is still TODO
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

Private Const INIT_QUICKSORT_STACK_SIZE As Long = 256
Private m_qsStack() As QSStack
Private m_qsStackPtr As Long

'Use our version of the "Resynthesizer" algorithm for texture synthesis.
' All relevant parameters MUST BE SET PRIOR TO CALLING!
Private Sub AttemptResynthesisAlgorithm()
    
    Dim startTime As Currency
    VBHacks.GetHighResTime startTime
    
    'Initialization for this function is non-trivial (from a perf standpoint).
    ' TODO: minimize initialization steps if params haven't changed from a previous run.
    InitializeInpainter
    
    PDDebug.LogAction "pdInpaint: initialization took " & VBHacks.GetTimeDiffNowAsString(startTime)
    PDDebug.LogAction m_numPointsLeft & " points will be synthesized"
    VBHacks.GetHighResTime startTime
    
    'Initialization prepared a list of all points in the destination image that need to be inpainted.
    ' We are now going to evaluate that list one-at-a-time until all pixels have been successfully
    ' "synthesized".
    Dim i As Long
    For i = 0 To m_numPointsLeft - 1
        
        Dim curPoint As PointLong, testPoint As PointLong, tmpPoint As PointLong
        curPoint = m_PointsRemaining(i)
        
        'Generate a list of nearby pixels that we can test to evaluate the "goodness" of the target pixel.
        ' We need to generate a custom list because
        ' 1) Not all pixels have been synthesized yet (so many won't have a color), and...
        ' 2) Neighboring pixels to this one may be out-of-bounds (so we can't use them)
        '
        'This step reads pixel values *out of the destination image*, by design, because it's trying to
        ' evaluate the quality of pixels that have already been synthesized.  (The step after this one
        ' will compare this point against random points from the *source image* to see if any of those
        ' perform better; if they do, the existing synthesized value, if any, will be replaced - this is
        ' why this algorithm must iterate, because data is constantly being reevaluated in-place.)
        m_CurNumNeighbors = 0
        
        '(Note that we can skip this step on the first pixel, because it is 100% guaranteed to have no
        ' valid neighbors!)
        If (i > 0) Then
            
            Dim j As Long
            For j = 0 To m_numNearestOffsets - 1
                
                'Comparison points are stored as offsets (not absolute points).  Thus we need to add coords
                ' to the current point to produce the absolute coordinates of the point-to-test.
                testPoint.x = curPoint.x + m_NearestOffsets(j).x
                testPoint.y = curPoint.y + m_NearestOffsets(j).y
                
                'Ensure the target pixel is:
                ' 1) in-bounds on the destination image
                '    (note: this is *not* relevant for seamless tile mode, because all boundaries get wrapped)
                ' 2) synthesized!  (If it *hasn't* been synthesized yet we obviously need to ignore it.)
                If IsPixelInBounds(testPoint) Then
                    tmpPoint = m_SynthesisState(testPoint.x, testPoint.y)
                    If (tmpPoint.x <> LONG_MAX) Then
                        
                        'Store the coordinates and color of this pixel (both are required during the comparison stage)
                        With m_Comparators(m_CurNumNeighbors)
                            .cOffsets = m_NearestOffsets(j)
                            .cColor = m_DstPixels(testPoint.x, testPoint.y)
                            
                            'We also want to know what this pixel's current *synthesized" coordinate is in the source
                            ' image (because we're going to compare its against neighbors in a later step).  This is
                            ' an important value to clone directly from the state array because the synthesized
                            ' coordinate may have a special flag value LONG_MAX which means this pixel has *not* been
                            ' synthesized yet.
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
        
        'Reset the "best" distance to an impossibly large value
        m_BestMatchDiff = LONG_MAX
            
        'Ensure we actually have at least one neighboring pixel to compare
        If (m_CurNumNeighbors > 0) Then
        
            'Search through the assembled list of neighboring pixels and calculate a net sum of the
            ' differences between pixel values.  The smallest net sum represents the best candidate pixel.
            j = 0
            For j = 0 To m_CurNumNeighbors - 1
                
                'Because the stored value for this pixel is an absolute coordinate into the source image,
                ' we need to *subtract* the current offset from it to arrive at a comparison pixel coordinate.
                ' (The call to EvaluatePixel(), below, will add individual offsets *back* onto the base
                ' coordinate during evalution.)
                With m_Comparators(j)
                    testPoint.x = .cCoordinates.x - .cOffsets.x
                    testPoint.y = .cCoordinates.y - .cOffsets.y
                End With
                
                'TODO: consider move these checks up, so we can bail immediately after evaluating x
                ' (without wasting time on y)
                If (testPoint.x < 0) Or (testPoint.y < 0) Then GoTo NextNeighbor
                If (testPoint.x >= m_SrcWidth) Or (testPoint.y >= m_SrcHeight) Then GoTo NextNeighbor
                
'                // skip computing differences of points
'                // we've already done this iteration. not mandatory.
'                if (*image_atc(s->tried, point) == i) continue;
                If (m_LastChecked(testPoint.x, testPoint.y) <> i) Then
                    
                    'Evaluate the pixel at this location, which will update the running sum for this pixel
                    EvaluatePixel testPoint
                    
                    'If this point is a perfect match, stop searching
                    If (m_BestMatchDiff = 0) Then Exit For
                    
                    'Note that this pixel has already been checked during this pass
                    m_LastChecked(testPoint.x, testPoint.y) = i
                    
                End If
                
NextNeighbor:
            Next j
        
        '/no valid neighbors
        End If
        
        'If we already have a perfect match, we don't need to compare random candidates for this pixel
        If (m_BestMatchDiff > 0) Then
            
            'We're now going to compare the quality of this point to the quality of some randomly sampled
            ' points from the source image.  This is critical for early pixels because they don't have
            ' enough neighbors to determine whether they're a good match or not (so randomly selected
            ' pixels will likely perform better!).
            '
            'As we get deeper into the pixel list, however, this step becomes less and less relevant
            ' because random pixels are unlikely to outperform pixels that have already survived previous
            ' iterations.  As such, we automatically scale-down the number of random candidates that we
            ' try as we get deeper and deeper into the list.
            Dim numRandomCandidatesToTry As Long
            numRandomCandidatesToTry = m_MaxRandomCandidates
            
            'Scale down by how many points we have already processed, so that we're only doing 25%
            ' as many random comparisons by the end of the list (because those points are likely
            ' well-matched after multiple iterations)
            numRandomCandidatesToTry = numRandomCandidatesToTry * (0.25 + 0.75 * ((m_numPointsLeft - i) / m_numPointsLeft))
            
            'Iterate random pixels and see if any outperform our current "winner"
            For j = 0 To numRandomCandidatesToTry - 1
            
                'Grab a random pixel from the list of valid source image points and evaluate it.
                ' Note that this function will replace m_BestMatch if this random point is better.
                EvaluatePixel m_ValidSourcePoints(m_Random.GetRandomIntRange_WH(0, m_boundValidSourcePoints))
                
                'If this randomly selected pixel proved to be a "perfect" candidate
                ' (unlikely but not impossible), stop searching.
                If (m_BestMatchDiff = 0) Then Exit For
                
            Next j
            
        End If

        'Assign the best-matching candidate to this destination pixel!
        m_DstPixels(curPoint.x, curPoint.y) = m_SrcPixels(m_BestMatchPt.x, m_BestMatchPt.y)
        
        'In our state array, note this pixel's current best-match coordinate.  (This coordinate may
        ' be replaced with a better alternative in subsequent passes, as more points get filled-in.)
        m_SynthesisState(curPoint.x, curPoint.y) = m_BestMatchPt
        
        'TODO: switch to progress bar, obviously
        If ((i And 8191) = 8191) Then PDDebug.LogAction (i + 1) & " points complete"
        
    Next i
    
    PDDebug.LogAction "Full resynthesis took " & VBHacks.GetTimeDiffNowAsString(startTime)
    
End Sub

'Inpainting requires a large amount of prep work (and prep resources).
' Try to minimize initialization calls if you can.
Private Sub InitializeInpainter()
    
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
    
    PDDebug.LogAction "Initializing cauchy table took " & VBHacks.GetTimeDiffNowAsString(startTime)
    VBHacks.GetHighResTime startTime
    
    'Prepare a sorted list of offsets to search when evaluating pixel quality.  This step has a non-trivial
    ' startup cost due to the sort requirement.
    InitializeOffsetList
    PDDebug.LogAction "Making and sorting offset list took " & VBHacks.GetTimeDiffNowAsString(startTime)
    VBHacks.GetHighResTime startTime
    
    'Generate a list of valid source points to sample *from*.  Points may be disqualified either implicitly
    ' (because they're inside the area-to-be-removed) or explicitly (maybe we could make a tool that allows
    ' the user to "paint" regions with either "use" or "dont use" flags?)
    '
    'Note that unlike the list of destination points, this one will never *exceed* the total number of source
    ' image coordinates.
    ReDim m_ValidSourcePoints(0 To m_SrcWidth * m_SrcHeight - 1) As PointLong
    
    Dim x As Long, y As Long, idxDst As Long
    idxDst = 0
    
    For y = 0 To m_SrcHeight - 1
    For x = 0 To m_SrcWidth - 1
    
        'TODO: obviously on inpainting mode we will need to check the mask to ensure this point can be used!
        With m_ValidSourcePoints(idxDst)
            .x = x
            .y = y
        End With
        idxDst = idxDst + 1
        
    Next x
    Next y
    
    m_boundValidSourcePoints = idxDst - 1
    If (UBound(m_ValidSourcePoints) <> m_boundValidSourcePoints) Then ReDim Preserve m_ValidSourcePoints(0 To m_boundValidSourcePoints) As PointLong
    PDDebug.LogAction "Initializing source point list took " & VBHacks.GetTimeDiffNowAsString(startTime)
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
    
    'TODO: for actual in-painting, this will be a small subset rect of the actual image.
    ' We'll need to factor x/y offsets into the image when doing anything with these arrays.
    idxDst = 0
    For y = 0 To m_DstHeight - 1
    For x = 0 To m_DstWidth - 1
        
        'Flag all destination pixels as being "not synthesized" (yet)
        m_SynthesisState(x, y).x = LONG_MAX
        
        'While we're here, dump this destination point to its storage queue too
        With m_PointsRemaining(idxDst)
            .x = x
            .y = y
        End With
        idxDst = idxDst + 1
        
    Next x
    Next y
    
    'Because we deliberately over-sized the destination list, we need to flag how many points *actually*
    ' exist in the array
    m_numPointsLeft = idxDst
    
    PDDebug.LogAction "Initializing destination point list took " & VBHacks.GetTimeDiffNowAsString(startTime)
    VBHacks.GetHighResTime startTime
    
    'Next, we need to randomize the order of the initial list of "points to synthesize".
    ' Shuffle using standard Fisher-Yates.
    Dim uBoundPoints As Long
    uBoundPoints = m_numPointsLeft - 1
    For i = uBoundPoints To 0 Step -1
        
        Dim j As Long
        j = m_Random.GetRandomIntRange_WH(0, i)
        
        'Swap points i and j
        Dim tmp As PointLong
        tmp = m_PointsRemaining(i)
        m_PointsRemaining(i) = m_PointsRemaining(j)
        m_PointsRemaining(j) = tmp
        
    Next i
    
    PDDebug.LogAction "Shuffling initial set of data points took " & VBHacks.GetTimeDiffNowAsString(startTime)
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
            
            'Bulk copy this set of points into place (after ensuring sufficient storage space, obviously)
            If ((idxDst + numPointsToRefine) > UBound(m_PointsRemaining)) Then ReDim Preserve m_PointsRemaining(0 To UBound(m_PointsRemaining) * 2 + 1) As PointLong
            VBHacks.CopyMemoryStrict VarPtr(m_PointsRemaining(idxDst)), VarPtr(m_PointsRemaining(0)), numPointsToRefine * 8     '8 = LenB(PointLong)
            idxDst = idxDst + numPointsToRefine
            
            'Stop adding points if we hit the pre-calculated max value
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
    
    PDDebug.LogAction "Appending repeat data points took " & VBHacks.GetTimeDiffNowAsString(startTime)
    VBHacks.GetHighResTime startTime
    
    'See the top of this module for details, but basically, in seamless tile mode, multiple neighboring pixels
    ' will map to the same position in the image due to edge-wrapping (e.g. [-1, 0] and [1, 0] will both map
    ' to [1, 0]).  We can save a lot of time by not checking duplicate values like these.
    '
    'To achieve this, we track the index of the last iteration that touched a given pixel, and if we see a
    ' repeat we skip subsequent checks.
    ReDim m_LastChecked(0 To m_SrcWidth - 1, 0 To m_SrcHeight - 1) As Long
    VBHacks.FillMemory VarPtr(m_LastChecked(0, 0)), m_SrcWidth * m_SrcHeight * 4, 255
    
    PDDebug.LogAction "Initializing attempts array took " & VBHacks.GetTimeDiffNowAsString(startTime)
    VBHacks.GetHighResTime startTime
    
End Sub

'Compare the pixel at a given coordinate to all pixels in the current Comparator list.  If this point is
' better than all previous ones we tested, update the module-level m_BestMatch value(s) to match.
Private Sub EvaluatePixel(ByRef srcPoint As PointLong)
    
    Dim totalDiff As Long, curDiff As Long
    
    Dim r As Long, g As Long, b As Long, a As Long
    Dim x As Long, y As Long
    
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
    ' source or destination's height.  TODO: revisit when implementing inpainting, as this will need
    ' to be user-controlled there.
    Dim minWidth As Long, minHeight As Long
    minWidth = PDMath.Min2Int(m_SrcWidth, m_DstWidth)
    minHeight = PDMath.Min2Int(m_SrcHeight, m_DstHeight)
    
    'Ensure available space for the offset list (x4 is needed because there are four quadrants of offsets)
    ReDim m_NearestOffsets(0 To minWidth * minHeight * 4 - 1) As PointLong
    
    Dim idxPoint As Long: idxPoint = 0
    
    'TODO: see if there is some algorithm for adding points in an already sorted way?
    Dim x As Long, y As Long
    For y = -minHeight + 1 To minHeight - 1
    For x = 0 To minWidth - 1
        
        'Don't add (0, 0); we don't want to compare points to themselves!
        If (x = 0) And (y = 0) Then GoTo SkipPoint
        
        'Add this point to the list, and its horizontal mirror.  (This improves sort perf by keeping
        ' similarly-distanced-points close together in the initial unsorted list.)
        With m_NearestOffsets(idxPoint)
            .x = x
            .y = y
        End With
        idxPoint = idxPoint + 1
        
        If (x > 0) Then
            With m_NearestOffsets(idxPoint)
                .x = -x
                .y = y
            End With
            idxPoint = idxPoint + 1
        End If

SkipPoint:
    Next x
    Next y
    
    'Size the list precisely, then sort it
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
    ' of edge pixels.  It's typically better to just ignore OOB pixels.  TODO: double-check this once inpainting
    ' code is finalized.)
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

'Potential neighbor offsets are sorted by their Euclidean distance from the origin
Private Function PointDistanceFrom00(ByVal idxSorted As Long) As Long
    With m_NearestOffsets(idxSorted)
        PointDistanceFrom00 = .x * .x + .y * .y
    End With
End Function

Private Sub SortCoordinates()
    
    If (m_numNearestOffsets <= 0) Then Exit Sub

    'Prep our internal stack (this is faster and more reliable than recursion)
    ReDim m_qsStack(0 To INIT_QUICKSORT_STACK_SIZE - 1) As QSStack
    m_qsStackPtr = 0
    m_qsStack(0).sLB = 0
    m_qsStack(0).sUB = m_numNearestOffsets - 1
    
    'Execute the sort (just a bog-standard quick sort at present)
    Dim idxLow As Long, idxHigh As Long
    Dim i As Long, j As Long
    Dim v As PointLong, vDist As Long
    
    'TODO: consider pre-calculating all distances so we don't have to repeatedly calculate them during sorting.
    Do
        
        'Load the next set of boundaries, and reset all pivots
        idxLow = m_qsStack(m_qsStackPtr).sLB
        idxHigh = m_qsStack(m_qsStackPtr).sUB
        
        i = idxLow
        j = idxHigh
        vDist = PointDistanceFrom00((i + j) \ 2)
        
        'TODO: standard optimization of preferentially using insertion sort on small ranges
        Do
            Do While (PointDistanceFrom00(i) < vDist)
                i = i + 1
            Loop
            Do While (PointDistanceFrom00(j) > vDist)
                j = j - 1
            Loop
            
            'Swap final indices as necessary
            If (i <= j) Then
                v = m_NearestOffsets(i)
                m_NearestOffsets(i) = m_NearestOffsets(j)
                m_NearestOffsets(j) = v
                i = i + 1
                j = j - 1
            End If
            
        Loop Until i > j
        
        'Conditionally add new entries to the processing stack
        If (idxLow < j) Then
            m_qsStack(m_qsStackPtr).sLB = idxLow
            m_qsStack(m_qsStackPtr).sUB = j
            m_qsStackPtr = m_qsStackPtr + 1
        End If
        
        If (i < idxHigh) Then
            m_qsStack(m_qsStackPtr).sLB = i
            m_qsStack(m_qsStackPtr).sUB = idxHigh
            m_qsStackPtr = m_qsStackPtr + 1
        End If
        
        'Decrement the stack pointer and continue as long as new sub-lists were added
        m_qsStackPtr = m_qsStackPtr - 1
        
    Loop While (m_qsStackPtr >= 0)
    
    'Free the stack before exiting
    Erase m_qsStack
    
End Sub

'From an arbitrary source image, create a seamlessly tiling new image
Friend Sub MakeSeamlessTile(ByRef srcDIB As pdDIB, ByRef dstDIB As pdDIB)
    
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
    AttemptResynthesisAlgorithm
    
    'Unalias all unsafe array wrappers
    srcDIB.UnwrapRGBQuadArrayFromDIB m_SrcPixels
    dstDIB.UnwrapRGBQuadArrayFromDIB m_DstPixels
    
End Sub

Private Sub Class_Initialize()

    Set m_Random = New pdRandomize
    
    'Eventually we will want to seed the randomizer randomly, but it's helpful during testing to
    ' use a fixed value to ensure that I haven't broken anything since the last revision lol
    m_Random.SetSeed_Int 0
    
    'm_Random.SetSeed_AutomaticAndRandom
    
End Sub
