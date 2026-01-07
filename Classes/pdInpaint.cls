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
'Copyright 2022-2026 by Tanner Helland
'Created: 21/April/22
'Last updated: 16/June/22
'Last update: remove a few unused variables
'
'Texture synthesis is an active field of study.  A great introduction to the topic is Paul Harrison's
' PhD thesis, "Image Texture Tools: Texture Synthesis, Texture Transfer, and Plausible Restoration",
' available here (link good as of April 2022):
'
'https://www.logarithmic.net/pfh-files/thesis/dissertation.pdf
'
'Paul wrote GIMP's "Resynthesizer" plugin to demonstrate his take on the algorithm, and as of 2022 his
' algorithm is still used in many GPL programs (and is still available in GIMP, too).
'
'A few years after Paul's paper, Adobe released their take on the topic in an algorithm called "PatchMatch".
' You can read the original PatchMatch paper here (link good as of April 2022):
'
'https://gfx.cs.princeton.edu/pubs/Barnes_2009_PAR/patchmatch.pdf
'
'Note that both Resynthesizer and PatchMatch have "official" implementations by their original authors,
' but these implementations are license-incompatible with PhotoDemon.  Argh.
'
'So I have gone ahead and implemented my own version of Paul's "Resynthesizer" algorithm, as I understand
' it from his original paper.  I think the resulting code works pretty damn well, with very favorable
' results compared to both the modern GIMP Resynthesizer plugin and Adobe's official content-aware fill
' tool.  I'm especially proud of the performance this class achieves despite VB6's limitations.
'
'More than anything, though, I hope that this BSD-licensed version of the algorithm draws attention
' from others who can help improve it further, particularly performance-wise.  It seems a shame to
' let Paul's interesting algorithm languish in GIMP alone instead of being further developed by many
' different open-source projects.
'
'This algorithm can be accessed from the "Edit > Content-Aware Fill" menu.  Various settings can be toggled
' from the matching dialog.  Note that GIMP exposes some additional features that this class does not
' (such as an option for filling pixels in "outside-in" order instead of randomly); I may add analogs of
' these options to PD as time allows.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'During debugging, it is helpful to report detailed process information.
' Please disable verbose debugging in production builds.
Private Const DEBUG_VERBOSE As Boolean = False

'Pixels can be synthesized in a number of different "orders".  The original paper suggests that random
' fill order is necessary for "[avoiding] any possibility of consistent skew due to the order in which
' pixels are selected."  That said, GIMP provides options for "outside-in" and "inside-out" order,
' and on some images outside-in order in particular does appear to provide more consistently appealing
' results - so I've added a similar option here.
Private Enum PD_FillOrder
    fo_Random = 0
    fo_OutsideIn = 1
    fo_InsideOut = 2
End Enum

#If False Then
    Private Const fo_Random = 0, fo_OutsideIn = 1, fo_InsideOut = 2
#End If

Private m_FillOrder As PD_FillOrder

'When sampling pixels from the source image, the caller can toggle each sampling direction.
' Note that selection shapes can be complex, so these toggles simply limit where pixels are
' sampled *relative to the entire selection's boundary box*.  Any unselected pixels within
' the selection's boundary box will *always* be available for sampling, but beyond the borders
' of the boundary box, sampling can be limited by toggling e.g. "sample up" to FALSE.
'
'This class doesn't care about which sampling directions you limit, because it expects you to
' "crop out" the relevant region of the image prior to calling this class.  The one thing that
' sampling direction affects is how the *center* of the sampling area is calculated, which is
' relevant for the "outside-in" and "inside-out" fill orders.  (When random sampling, it doesn't
' matter if you limit sampling direction.)  This class can auto-calculate the center of the
' sampling region, obviously, but if you've limited sampling direction then an auto-calculation
' won't be correct because bordering sampling regions aren't symmetrical.
'
'Anyway, it's up to the caller to notify us of the center point of the selection.  This class
' can handle everything else.
Private m_SamplingCenter As PointLong

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

'One of the first steps prior to inpainting is preparing a list of points that must be inpainted.
' This list consists of coordinates into the destination image, whose order has been deliberately randomized
' (depending on the user's settings - by *default* it's fully randomized, but the user can also select
' "outside-in" or "inside-out" order, in which case the list is sorted accordingly).
'
'Points chosen early in the refinement process must be re-synthesized again later in the process, when more
' neighboring points have been filled in, so this list *will* contain duplicates, by design.  It may also be
' quite large (5x or more) compared to the original set of points, since the user can specify a quality
' parameter that affects how many refinement passes we apply.
Private m_PointsRemaining() As PointLong, m_numPointsLeft As Long

'When trying to match pixels, we need to randomly sample pixels from the source image as potential candidates.
' Because some source image pixels may be invalid (because e.g. they lie in the region the user wants removed),
' we can't simply grab *any* pixel - so during initialization, we'll build a list of valid coordinates to
' choose from.  (Note that some applications - like seamless tile generation - may use every source pixel.
' That's fine too!)
Private m_ValidSourcePoints() As PointLong, m_boundValidSourcePoints As Long

'We need an array (at the same dimensions as the destination image) that tracks the synthesis state of
' each destination pixel.  Because it is also necessary to track the current synthesized point location,
' we simply use a coordinate array, with unsynthesized pixels initialized to the magic value LONG_MAX
' (stored in the x-coord).
'
'While we're here, we can also add an easy performance optimization.  When evaluating neighboring pixels,
' we run the risk of evaluating the same pixel multiple times.  This is especially true near boundaries in
' seamless tile mode, because wrapping ensures that each neighbor along a boundary appears in the neighbor
' list 2x - once in its actual position, then again as the result of mirroring a corresponding pixel across
' the nearby boundary line (consider [-1, 0] and [1, 0], for example).  Seamless tile mode can thus be
' greatly accelerated by skipping pixels that have already been evaluated.  We do this by simply tracking
' the index of the last iteration that touched a given pixel, in the .z value of this state tracker.
Private m_SynthesisState() As PointLong3D

'A central tenet of the original Resynthesizer paper is using the Cauchy distribution when calculating
' pixel differences (vs e.g. a more traditional gaussian distribution).  Cauchy is a much flatter
' distribution (https://en.wikipedia.org/wiki/Cauchy_distribution) which makes it more inclusive of
' outliers.  This is favorable for inpainting because we don't want *perfect* matches on each pixel -
' this will produce a noticeably "artificial" look because your brain will notice the obvious repeating
' patterns from the source image.  Instead, we want lots of "good enough" matches that mix-and-match
' data from elsewhere in the image, hiding any seams organically.
'
'Unfortunately, Cauchy requires an expensive Log() calculation (as used here, so we can get away with
' addition during summing) and it cannot be hard-coded at compile-time because it relies on a critical
' input parameter "sigma" from the user - so we precalculate all possible pixel differences in a
' fixed-size table during initialization, then use that table for all color comparisons.  Note that we
' also scale the default float-values by some arbitrary integer amount (currently 32768) to avoid the
' need for expensive floating-point math on each comparison.
'
'Note that the table is deliberately resized to all possible values on the range [-255, 255], but we
' also calculate a special "maximum value" (equal to a difference of 256), which can be used as a
' special sum to indicate "maximum difference" (which discourages further matching due to the extreme
' penalty).
Private Const m_CauchyScale As Long = 32768
Private m_Cauchy(-255 To 255) As Long
Private m_CauchyMax As Long

'This term is actually the "sigma" value from the original paper (pp 43), but I'm fairly certain it's
' the analog of the GIMP plugin feature "sensitivity to outliers".  The range must be [0.0, 1.0],
' with a value of 0 penalizing outliers most severely.  As the original paper says, "...so long as we
' are only concerned with finding the most likely match, not relative likelihoods, the Euclidean
' distance metric may be simulated in this metric by choosing a very large [sigma]."  The paper
' suggests a default value of 30 which makes no sense, so I assume it is meant to be "30 / 256" (~0.12).
' In PhotoDemon I've bumped the default value up slightly, to 0.15, purely for aesthetic reasons
' (it looks better in the UI lol).
Private m_AllowOutliers As Double
Private Const ALLOW_OUTLIERS_DEFAULT As Double = 0.15

'Some fixed set of points around each pixel must be searched and compared to see if a newly placed
' pixel is a good match.  We can't do a naive raster-order search because we want to sort the pixels in
' ascending distance from the origin pixel.  As the original paper says (pp 40):  "...The best fit out
' of each of these locations is selected, based on a comparison of the pattern formed by the [n] nearest
' known neighbours of the current pixel to the pattern formed by pixels at corresponding offsets about
' each candidate location."
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
' algorithm. (Similarly, some points may lie entirely outside image boundaries.)  Thus from the full set
' of *possible* comparison points, we must generate a *custom* list of target points for each pixel,
' consisting only of pixels that have already been synthesized.
Private m_Comparators() As ip_Comparator

'The user passes a parameter that tells us how many neighboring pixels to compare.  More neighbors can
' produce a better result, but incurs a corresponding performance penalty.  The original paper suggests
' a default value of 30.  (Note that boundary or selection issues - e.g. randomly selecting a pixel in
' the middle of a large mask, but the user has specified a very small default radius - may prevent us
' from reaching the full neighbor count, so we also track the *actual* neighbor count for the current
' candidate pixel; as the maximum value approaches 0, we're basically selecting a new pixel at random,
' but further refinement passes will help us reject poor early candidates.)
Private m_MaxNumNeighbors As Long, m_CurNumNeighbors As Long

'After much testing, I've gone with a default radius of 21 for the neighbor comparison.  This results in
' a circle of diameter 5 being used for fully-surrounded pixels, as in the following diagram:
'  X X X
'X X X X X
'X X T X X
'X X X X X
'  X X X
'
'(The target pixel [T] is not evaluated, by design.)  This provides a significant speed improvement vs
' the original paper's default of 30, but without a noticeable quality penalty, in my opinion.
Private Const MAX_NEIGHBORS_DEFAULT As Long = 20

'After comparing the nearest [n] candidates, we also want to sample some number of random pixels to
' see if any of those work better than the one we selected.  This step is especially important for early
' pixels, since they will not have *any* neighboring pixels available for comparisons (because the mask
' area has not been synthesized yet!).  This value represents a maximum upper bound for how many random
' pixels get compared.  More produces better results but with a corresponding performance penalty.
'
'The original paper used 80 random candidates but made no discussion as to the reasoning behind this.
' For inpainting, I've found it preferable to use a smaller random pixel count combined with a higher
' refinement threshold (to ensure each pixel is tested at least *twice* against the maximum number
' of random candidates).  This produces excellent results with very good performance, and seems a
' better fit for inpainting vs the seamless-tile emphasis of the original paper.
Private m_MaxRandomCandidates As Long
Private Const RANDOM_CANDIDATES_DEFAULT As Long = 60

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
'The suggested default value in the paper is 0.75.  In PhotoDemon, I've chosen to handle this a little
' differently.  Instead of doing a single pass over the image and then applying the refinement parameter,
' my version does *two* passes over the image before applying the refinement parameter.  I've found this
' to meaningfully reduce the likelihood of "garbage" patches that match very poorly, by ensuring that
' every pixel gets refined at least once.  After that full second pass of the image, the refinement
' parameter kicks in and the first [p * n] are revisited a 3rd time, then a 4th time, etc.  Because we
' guarantee a full second pass over the image, a smaller default refinement threshold is acceptable.
Private m_Refinement As Double
Private Const REFINEMENT_DEFAULT As Double = 0.5

'Potential search radius when matching pixels.  In a perfect world, we'd search the whole image until we
' find enough pixels to reach our test threshold, but because we have to build a sorted list of offsets,
' this becomes problematic on large images.  Instead, we need to limit the potential search radius to some
' reasonable size around each pixel.
'
'This value is my own creation (it's not present in the original paper, which focuses more on "full image
' synthesis" while masked regions are mostly a side note).  GIMP's current Resynthesis plugin defaults
' to a rather paltry 50 pixels, which I find too small for proper sampling on your average 12+ megapixel
' JPEG.  Instead, we default to 200 pixels and allow for much larger radii if desired.
'
'Importantly, note that this radius *will be limited by the size of the incoming image*, so if you pass
' a 64x64 image, a radius of 200 doesn't matter.  (The smaller of an image's width/height and this radius
' will automatically be used.)
Private m_ComparatorRadius As Long
Private Const COMPARE_RADIUS_DEFAULT As Long = 200

'Dimensions and pixel array for the source/destination image.  Note that this array *unsafely* aliases
' the target image, and *must* be manually unaliased before any class functions exit.  (In early versions
' of this code, separate images were used for source and destination pixel placement.  However, reworking
' the class to use a single shared image greatly reduces cache pressure and meaningfully improved
' performance, while also reducing memory requirements.)
Private m_SrcWidth As Long, m_SrcHeight As Long
Private m_SrcPixels() As RGBQuad, m_SrcSA As SafeArray2D

'Point order needs to be semi-randomized for the algorithm to produce useful results.
' (This is a key element described by the original paper.)
'
'We use PD's internal randomizer to accomplish this, but you could also use Rnd() or similar.
Private m_Random As pdRandomize

'The current best-match candidate (both its coordinates, and its "difference" from the test point,
' where a difference of 0 means "perfect match").  It's stored at module-level because separate functions
' need to access it and this is faster than passing it frequently back-and-forth as a parameter.
' (A multi-threaded redesign would also need to rework this; each thread would need to maintain its
' own current best match, and then the best of *those* matches would need to be selected as the final value.)
Private m_BestMatchPt As PointLong, m_BestMatchDiff As Long

'In inpainting mode, we need class-level access to the passed mask.  This class requires the mask array
' to be the same size as the source and destination images; this greatly simplifies bounds-checking,
' and provides a large performance boost relative to previous attempts (which took a separate mask rect)
' as a parameter.  The mask is just a byte array, and the synthesizer will treat all 0 values as
' "do not synthesize" and all non-zero values as "synthesize".  It does not perform any blending, by design -
' antialiasing or feathering must be handled by the caller after this class finishes its work.
Private m_MaskBytes() As Byte

'This class supports both 24-bpp RGB and 32-bpp RGBA matching.  If the source image is fully opaque
' (a primary use-case) then we use a 24-bpp evaluator on the inner-loop for a welcome speed boost.
' This behavior is automatically determined based on the opacity state of the source image.
Private m_EvaluateRGBOnly As Boolean

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

'Property set/get.  Bad values will automatically be clamped to valid ranges, with appropriate warnings
' sent to the InternalError function (which you can handle as you please; by default they are dumped to
' the debug window).
Friend Sub SetAllowOutliers(ByVal newValue As Double)
    
    Const FUNC_NAME As String = "SetAllowOutliers"
    
    If (newValue < 0.01) Or (newValue > 1#) Then
        InternalError FUNC_NAME, "bad value; acceptable range is [0.01, 1.0]"
        m_AllowOutliers = ALLOW_OUTLIERS_DEFAULT
    Else
        m_AllowOutliers = newValue
    End If
    
End Sub

Friend Sub SetFillOrder(ByVal newValue As Long)
    
    Const FUNC_NAME As String = "SetFillOrder"
    
    If (newValue < fo_Random) Or (newValue > fo_InsideOut) Then
        InternalError FUNC_NAME, "bad value; acceptable range is [0, 2]"
        m_FillOrder = fo_Random
    Else
        m_FillOrder = newValue
    End If
    
End Sub

Friend Sub SetMaxNumNeighbors(ByVal newValue As Long)
    
    Const FUNC_NAME As String = "SetMaxNumNeighbors"
    
    If (newValue < 4) Or (newValue > 100) Then
        InternalError FUNC_NAME, "bad value; acceptable range is [4, 100]"
        m_MaxNumNeighbors = MAX_NEIGHBORS_DEFAULT
    Else
        m_MaxNumNeighbors = newValue
    End If
    
End Sub

Friend Sub SetMaxRandomCandidates(ByVal newValue As Long)
    
    Const FUNC_NAME As String = "SetMaxRandomCandidates"
    
    If (newValue < 5) Or (newValue > 200) Then
        InternalError FUNC_NAME, "bad value; acceptable range is [5, 200]"
        m_MaxRandomCandidates = RANDOM_CANDIDATES_DEFAULT
    Else
        m_MaxRandomCandidates = newValue
    End If
    
End Sub

Friend Sub SetRefinement(ByVal newValue As Double)
    
    Const FUNC_NAME As String = "SetRefinement"
    
    If (newValue < 0#) Or (newValue > 0.99) Then
        InternalError FUNC_NAME, "bad value; acceptable range is [0.0, 0.99]"
        m_Refinement = REFINEMENT_DEFAULT
    Else
        m_Refinement = newValue
    End If
    
End Sub

'Notify this class of the "center point" to use when "outside-in" or "inside-out" fill order
' is chosen.  You can pass [0, 0] to have this class auto-determine the center for you
' (it will use [width / 2, height / 2]).
Friend Sub SetSamplingCenter(ByVal newX As Long, ByVal newY As Long)
    m_SamplingCenter.x = newX
    m_SamplingCenter.y = newY
End Sub

Friend Sub SetSearchRadius(ByVal newValue As Long)
    
    Const FUNC_NAME As String = "SetSearchRadius"
    
    If (newValue < 5) Or (newValue > 500) Then
        InternalError FUNC_NAME, "bad value; acceptable range is [5, 500]"
        m_ComparatorRadius = COMPARE_RADIUS_DEFAULT
    Else
        m_ComparatorRadius = newValue
    End If
    
End Sub

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
        Message "Repainting selected regions..."
    End If
    
    'Initialization for this function is non-trivial (from a perf standpoint).
    ' TODO: minimize initialization steps if params haven't changed from a previous run?
    If (Not InitializeInpainter) Then
        InternalError FUNC_NAME, "initialization failed"
        Exit Function
    End If
    
    If DEBUG_VERBOSE Then PDDebug.LogAction "pdInpaint: initialization took " & VBHacks.GetTimeDiffNowAsString(startTime)
    PDDebug.LogAction m_numPointsLeft & " points will be synthesized"
    VBHacks.GetHighResTime startTime
    
    Dim progCheck As Long
    If m_ShowProgressUI Then
        ProgressBars.SetProgBarMax m_numPointsLeft
        progCheck = ProgressBars.FindBestProgBarValue()
    End If
    
    'Initialization prepared a list of all points in the destination image that need to be inpainted.
    ' We are now going to evaluate that list one-at-a-time until all pixels have been successfully
    ' "synthesized".
    Dim i As Long
    For i = 0 To m_numPointsLeft - 1
        
        Dim curPoint As PointLong, testPoint As PointLong
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
            If (testPoint.x >= 0) Then
            If (testPoint.x < m_SrcWidth) Then
            If (testPoint.y >= 0) Then
            If (testPoint.y < m_SrcHeight) Then
                
                'Synthesis state is stored as a special flag value inside a separate tracking array.
                ' Validate it before continuing.  Note that this special flag means one of two things:
                ' either this point has already been successfully synthesized, *or* it lies outside the
                ' mask and is fine to use as-is.
                If (m_SynthesisState(testPoint.x, testPoint.y).x <> LONG_MAX) Then
                    
                    'This point has been synthesized (or in content-aware fill mode, it's not masked).
                    
                    'Store the offsets (relative to the target pixel) and current color of this pixel;
                    ' both will be used in the following step to produce a "goodness of fit" metric.
                    With m_Comparators(m_CurNumNeighbors)
                        .cOffsets = m_NearestOffsets(j)
                        .cColor = m_SrcPixels(testPoint.x, testPoint.y)
                        
                        'We also want to know what this pixel's current *synthesized" coordinate is in the source
                        ' image (because we're going to compare its against neighbors in a later step).
                        .cCoordinates.x = m_SynthesisState(testPoint.x, testPoint.y).x
                        .cCoordinates.y = m_SynthesisState(testPoint.x, testPoint.y).y
                    End With
                    
                    m_CurNumNeighbors = m_CurNumNeighbors + 1
                    
                    'We must never exceed the user's max neighbor value (because that's the upper limit for
                    ' all tracking arrays)
                    If (m_CurNumNeighbors >= m_MaxNumNeighbors) Then Exit For
                    
                End If
                
            End If
            End If
            End If
            End If
        
        Next j
        
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
                If (testPoint.x >= m_SrcWidth) Then GoTo NextNeighbor
                If (testPoint.y < 0) Then GoTo NextNeighbor
                If (testPoint.y >= m_SrcHeight) Then GoTo NextNeighbor
                
                'We also want to ensure the target point is not masked.  (If it is, we don't want to consider it
                ' as a potential replacement!)
                If (m_MaskBytes(testPoint.x, testPoint.y) <> 0) Then GoTo NextNeighbor
                
                'If we've already compared this point this iteration, skip it.  This is a likely scenario as
                ' patches become well-established, since neighboring pixels are likely taken from the same
                ' "patch" in the original image.  (Also, EvaluatePixel_RGB/A, below, is highly expensive since
                ' we have to compare the target pixel to *each* entry in the comparator list - so skipping it
                ' is a huge perf win.)
                '
                'NOTE: this idea came from https://github.com/notwa/resynth which I was benchmarking against
                ' PD's implementation, and I could not for the life of me figure out why it kept crushing me
                ' in performance on texture synthesis.  Thank you to those authors for the clever optimization
                ' strategy (it may actually be adopted from earlier resynthesizer work, idk).
                If (m_SynthesisState(testPoint.x, testPoint.y).z <> i) Then
                    
                    'Evaluate the test pixel; if it's a better fit than the current one, EvaluatePixel_RGBA will
                    ' update the class-level running sum accordingly.
                    If m_EvaluateRGBOnly Then
                        EvaluatePixel_RGB testPoint
                    Else
                        EvaluatePixel_RGBA testPoint
                    End If
                    
                    'If this point is a perfect match against the current comparator list, stop searching because
                    ' we're not going to find anything better this time around.
                    If (m_BestMatchDiff = 0) Then Exit For
                    
                    'Flag this pixel as "checked during this pass" so we can skip it if we encounter it again.
                    m_SynthesisState(testPoint.x, testPoint.y).z = i
                    
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
            
            'Scale down by how many points we have already processed, so that we're only doing 25%
            ' as many random comparisons by the end of the list (because those points are likely
            ' well-matched after multiple iterations).  This is an optimization of my own creation so
            ' please do not blame the original Resynthesis paper for this behavior!
            numRandomCandidatesToTry = m_MaxRandomCandidates * (0.25 + 0.75 * ((m_numPointsLeft - i) / m_numPointsLeft))
            
            'Iterate that many random pixels and see if any outperform our current "winning" choice
            For j = 0 To numRandomCandidatesToTry - 1
                
                If m_EvaluateRGBOnly Then
                    EvaluatePixel_RGB m_ValidSourcePoints(m_Random.GetRandomIntRange_WH(0, m_boundValidSourcePoints))
                Else
                    EvaluatePixel_RGBA m_ValidSourcePoints(m_Random.GetRandomIntRange_WH(0, m_boundValidSourcePoints))
                End If
                
                'If this randomly selected pixel proved to be a "perfect" candidate (unlikely but
                ' not impossible), stop searching.
                If (m_BestMatchDiff = 0) Then Exit For
                
            Next j
            
        End If

        'Assign the best-matching candidate to this destination pixel!
        m_SrcPixels(curPoint.x, curPoint.y) = m_SrcPixels(m_BestMatchPt.x, m_BestMatchPt.y)
        
        'In our state array, note this pixel's current best-match coordinate.  (This coordinate may
        ' be replaced with a better alternative in subsequent passes, as more points get filled-in.)
        With m_SynthesisState(curPoint.x, curPoint.y)
            .x = m_BestMatchPt.x
            .y = m_BestMatchPt.y
        End With
        
        'TODO: add cancellation support
        If m_ShowProgressUI And ((i And progCheck) = 0) Then
            If Interface.UserPressedESC() Then Exit For
            ProgressBars.SetProgBarVal i
        End If
        
    Next i
    
    'If we raised the progress bar, hide it before exiting
    If m_ShowProgressUI Then ProgressBars.ReleaseProgressBar
    
    SynthesizeTexture = (Not g_cancelCurrentAction)
    
    If DEBUG_VERBOSE Then PDDebug.LogAction "Total inpainting process took " & VBHacks.GetTimeDiffNowAsString(startTime)
    
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
        ' over-selection of edge pixels in the input texture - a desireable behavior, as edges are hard to
        ' re-synthesize smoothly).
        Dim invAllowOutliers As Double
        invAllowOutliers = 1# / m_AllowOutliers
        
        For i = -255 To 256
            
            Dim tmpFloat As Double
            tmpFloat = GetCauchy(i / 256# * invAllowOutliers) / GetCauchy(invAllowOutliers)
            
            'Scale by an arbitrary integer value for faster calculations on the inner loop
            tmpFloat = tmpFloat * m_CauchyScale
            If (i < 256) Then
                m_Cauchy(i) = Int(tmpFloat + 0.5)
            Else
                m_CauchyMax = Int(tmpFloat + 0.5)
            End If
            
        Next i
    
    'If the user requires perfect matches (ugh) then *any* difference will receive max penalty.
    ' (This approach really isn't recommended and honestly, I may just disable this option entirely.)
    Else
        For i = -255 To 255
            m_Cauchy(i) = m_CauchyScale
        Next i
        m_CauchyMax = m_CauchyScale
        m_Cauchy(0) = 0
    End If
    
    'Scale the maximum cauchy value by 3 in RGB mode, 4 in RGBA mode (because for each pixel,
    ' we're comparing that many channels)
    If m_EvaluateRGBOnly Then m_CauchyMax = m_CauchyMax * 3 Else m_CauchyMax = m_CauchyMax * 4
    
    If DEBUG_VERBOSE Then PDDebug.LogAction "Initializing cauchy table took " & VBHacks.GetTimeDiffNowAsString(startTime)
    VBHacks.GetHighResTime startTime
    
    'Prepare a sorted list of offsets to search when evaluating pixel quality.
    ' This step has a non-trivial startup cost due to the sort requirement (although this penalty has
    ' been greatly reduced in later versions of this class).
    InitializeOffsetList
    If DEBUG_VERBOSE Then PDDebug.LogAction "Making and sorting offset list took " & VBHacks.GetTimeDiffNowAsString(startTime)
    VBHacks.GetHighResTime startTime
    
    'The list of "comparators" (neighboring pixels that we try to evaluate on every synthesis)
    ' has a fixed size specified by the user.  Note that we may not perform this many comparisons
    ' for *every* pixel, but this gives us a safe upper limit so we can ignore array bound checks
    ' on perf-sensitive inner loops.
    ReDim m_Comparators(0 To m_MaxNumNeighbors - 1) As ip_Comparator
    
    'Now we need to generate a bunch of different lists.
    
    'First, a list of valid source points to sample *from*.  Points may be disqualified either implicitly
    ' (because they're inside the area-to-be-removed) or explicitly (maybe we could make a tool that allows
    ' the user to "paint" regions with either "use" or "dont use" flags, a la Photoshop?)
    '
    'Note that unlike the list of destination points, this one will never *exceed* the total number of source
    ' image coordinates, but it can be *smaller* than the total number of source image coordinates.
    ReDim m_ValidSourcePoints(0 To m_SrcWidth * m_SrcHeight - 1) As PointLong
    
    'We also need a list of points that need to be synthesizsed.  Note that this list will ultimately be
    ' larger than the set of incoming points, potentially *much* larger, because pixels selected early in
    ' the process must be revisited again later in the process, when more neighboring points have been
    ' synthesized.  (Because of this, we deliberately over-size this list to start, and may resize it
    ' further as additional points are added.)
    ReDim m_PointsRemaining(0 To m_SrcWidth * m_SrcHeight * 4 - 1) As PointLong
    
    'We also need to track the "synthesis state" of each pixel-to-be-filled.  (Pixels that have not yet
    ' been synthesized must be flagged with a special value, so we know to ignore them during comparisons
    ' between neighboring pixels.)
    ReDim m_SynthesisState(0 To m_SrcWidth - 1, 0 To m_SrcHeight - 1) As PointLong3D
    
    Dim x As Long, y As Long, idxSrc As Long, idxDst As Long
    idxSrc = 0
    idxDst = 0
    
    'All unmasked pixels get added to the "valid potential replacement pixel" list.
    For y = 0 To m_SrcHeight - 1
    For x = 0 To m_SrcWidth - 1
    
        'The user does *not* want this point synthesized
        If (m_MaskBytes(x, y) = 0) Then
            m_ValidSourcePoints(idxSrc).x = x
            m_ValidSourcePoints(idxSrc).y = y
            idxSrc = idxSrc + 1
            
            'Also flag the synthesis state tracker as "already synthesized by the pixel at this location"
            m_SynthesisState(x, y).x = x
            m_SynthesisState(x, y).y = y
        
        'The user *does* want this point synthesized
        Else
            m_PointsRemaining(idxDst).x = x
            m_PointsRemaining(idxDst).y = y
            idxDst = idxDst + 1
            
            'Flag the synthesis state tracker as "not synthesized yet"; we need to know this when comparing
            ' potential pixels to their neighbors when evaluating patch quality (unsynthesized neighbors
            ' must be ignored).
            m_SynthesisState(x, y).x = LONG_MAX
            
        End If
        
        'Flag the synthesis state of each pixel as "not tested during this iteration".
        ' (In the perf-sensitive inner loop, we use this flag to know if a pixel has already been evaluated,
        ' which lets us skip additional tests until we switch to a new pixel.)
        m_SynthesisState(x, y).z = -1
        
    Next x
    Next y
    
    'Failsafe check to ensure we have at least one valid source point
    If (idxSrc <= 0) Then
        InternalError FUNC_NAME, "no source points to synthesize from!"
        Exit Function
    End If
    
    'Resize the "valid source point list" to its precise boundary limit.  (We need to randomly select
    ' points from this list during the synthesis process, and we're just going to grab random points
    ' from this array to do it - since we already know every point in this list is valid!)
    m_boundValidSourcePoints = idxSrc - 1
    If (UBound(m_ValidSourcePoints) <> m_boundValidSourcePoints) Then ReDim Preserve m_ValidSourcePoints(0 To m_boundValidSourcePoints) As PointLong
    
    'Because we deliberately over-sized the destination list, we need to flag how many points *actually*
    ' exist in the array before appending additional points.
    m_numPointsLeft = idxDst
    If (m_numPointsLeft = 0) Then
        InternalError FUNC_NAME, "no points to synthesize!"
        Exit Function
    End If
    
    If DEBUG_VERBOSE Then PDDebug.LogAction "Initializing coordinate lists took " & VBHacks.GetTimeDiffNowAsString(startTime)
    VBHacks.GetHighResTime startTime
    
    'Next, we need to deal with the initial list of "points that must be synthesized".  The original paper
    ' suggests using a fully randomized order, with the caveat that if this produces poor results you can
    ' just run it again.
    '
    'Modern versions of GIMP's resynthesizer plugin also provide a choice for "outside-in" or "inside-out" order.
    ' It's possible these provide better results on specific types of images (e.g. strong patterns may benefit
    ' from "outside-in" order, because the pattern will be better matched from the get-go vs completely random).
    
    'If we need to sort the initial list in ascending (inside-out) or descending (outside-in) order,
    ' do so now.
    If (m_FillOrder <> fo_Random) Then
        
        SortPointsRemaining
        If DEBUG_VERBOSE Then PDDebug.LogAction "Sorting initial set of data points took " & VBHacks.GetTimeDiffNowAsString(startTime)
        
    'Otherwise, just do an in-place shuffle of the "points that must be synthesized" list
    ' using standard Fisher-Yates.
    Else
        
        'Next, we need to randomize the order of the initial list of "points that must be synthesized".
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
    
    End If
    
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
    
    InitializeInpainter = True
    Exit Function

InitializationFailed:
    InternalError FUNC_NAME, "internal error", Err.Number
    
End Function

'Compare the pixel at a given coordinate to all pixels in the current Comparator list.  If this point is
' better than all previous ones we tested, update the module-level m_BestMatch value(s) to match.
' (Per the name, this compares *color only*, and should only be used on 32-bpp sources with fully opaque pixels.)
Private Sub EvaluatePixel_RGB(ByRef srcPoint As PointLong)
    
    Dim totalDiff As Long, curDiff As Long
    
    Dim r As Long, g As Long, b As Long
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
                
                'We also need to ensure this pixel is *not* part of the mask.
                ' (We're indexing into the source image, and selected pixels will *still retain their
                ' original contents*.)
                If (m_MaskBytes(x, y) = 0) Then
                    
                    'Use the pre-constructed Cauchy table to calculate distance
                    Dim origColor As RGBQuad, candidateColor As RGBQuad
                    origColor = m_SrcPixels(x, y)
                    candidateColor = m_Comparators(i).cColor
                    
                    'Swap chars into ints to prevent VB from overflowing on char arithmetic
                    b = candidateColor.Blue
                    g = candidateColor.Green
                    r = candidateColor.Red
                    curDiff = m_Cauchy(b - origColor.Blue) + m_Cauchy(g - origColor.Green) + m_Cauchy(r - origColor.Red)
                
                'If this pixel has not yet been synthesized, add the *average* value of the current comparison
                ' to the core.  (This effectively ignores this pixel without penalizing it.)
                Else
                    If (numPixelsCompared <> 0) Then
                        curDiff = totalDiff \ numPixelsCompared
                    Else
                        curDiff = m_CauchyMax
                    End If
                End If
                
            'If this pixel lies outside the original image's boundaries, assign it a fixed-size penalty.
            ' (We could also ignore the pixel, but this creates unwanted favoritism for edge pixels because
            ' they simply won't have as many neighboring pixels to compare - so this is a better compromise
            ' for final quality.)
            Else
                curDiff = m_CauchyMax
            End If
        Else
            curDiff = m_CauchyMax
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

'Compare the pixel at a given coordinate to all pixels in the current Comparator list.  If this point is
' better than all previous ones we tested, update the module-level m_BestMatch value(s) to match.
' (Per the name, this compares color *and* alpha values, and must be used on 32-bpp sources with non-opaque pixels.)
Private Sub EvaluatePixel_RGBA(ByRef srcPoint As PointLong)
    
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
                
                'We also need to ensure this pixel is *not* part of the mask.
                ' (We're indexing into the source image, and selected pixels will *still retain their
                ' original contents*.)
                If (m_MaskBytes(x, y) = 0) Then
                    
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
                
                'If this pixel has not yet been synthesized, add the *average* value of the current comparison
                ' to the core.  (This effectively ignores this pixel without penalizing it.)
                Else
                    If (numPixelsCompared <> 0) Then
                        curDiff = totalDiff \ numPixelsCompared
                    Else
                        curDiff = m_CauchyMax
                    End If
                End If
                
            'If this pixel lies outside the original image's boundaries, assign it a fixed-size penalty.
            ' (We could also ignore the pixel, but this creates unwanted favoritism for edge pixels because
            ' they simply won't have as many neighboring pixels to compare - so this is a better compromise
            ' for final quality.)
            Else
                curDiff = m_CauchyMax
            End If
        Else
            curDiff = m_CauchyMax
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
    
    'Constrain by the smaller of the source or destination's height, or the user-supplied radius parameter.
    Dim minWidth As Long, minHeight As Long
    minWidth = PDMath.Min2Int(m_SrcWidth, m_ComparatorRadius)
    minHeight = PDMath.Min2Int(m_SrcHeight, m_ComparatorRadius)
    
    'Ensure available space for the offset list (x4 is needed because there are four quadrants of offsets)
    ReDim m_NearestOffsets(0 To minWidth * minHeight * 4 - 1) As PointLong
    
    'Convert minimum values to fixed bounds for the forthcoming loops
    minHeight = minHeight - 1
    minWidth = minWidth - 1
    
    Dim idxPoint As Long: idxPoint = 0
    
    'There may be an algorithm for adding rectangular points in an already-sorted way, but I don't know it
    ' so we're just going to add points "rectangularly" before manually sorting them into ascending distance order.
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
    SortNeighborCoordinates
    
End Sub

'This is a modified modulo function; it handles negative values differently from VB's built-in Mod
Private Function ModuloInt(ByVal quotient As Long, ByVal divisor As Long) As Long
    ModuloInt = quotient - Fix(quotient / divisor) * divisor
    If (ModuloInt < 0&) Then ModuloInt = ModuloInt + divisor
End Function

Private Sub SortNeighborCoordinates()
    
    If (m_numNearestOffsets <= 0) Then Exit Sub
    
    'To improve performance, we're not gonna sort the coordinate list in-place.
    ' Instead, we're going to pre-calculate all distances, sort just the distance list,
    ' then rebuild the original coordinate array accordingly.
    Dim sortList() As QSEntry
    ReDim sortList(0 To m_numNearestOffsets - 1) As QSEntry
    
    Dim i As Long
    For i = 0 To m_numNearestOffsets - 1
        
        sortList(i).idxOriginal = i
        With m_NearestOffsets(i)
            
            'Potential neighbor offsets are sorted by their Euclidean distance from the origin.
            ' (We only need relative values, so don't waste energy on a square-root.)
            sortList(i).ptDistance = .x * .x + .y * .y
            
        End With
        
    Next i
    
    'Hand the rest of the sort off to a generic coordinate sorter
    OptimizedSort_CoordinateList_Ascending sortList, m_numNearestOffsets
    
    'The coordinate list has been successfully sorted.  Clone the original list, then copy each entry
    ' from that temporary clone into its final position in the "official" list.
    Dim tmpPoints() As PointLong
    ReDim tmpPoints(0 To m_numNearestOffsets - 1) As PointLong
    CopyMemoryStrict VarPtr(tmpPoints(0)), VarPtr(m_NearestOffsets(0)), 8 * m_numNearestOffsets
    
    For i = 0 To m_numNearestOffsets - 1
        m_NearestOffsets(i) = tmpPoints(sortList(i).idxOriginal)
    Next i
    
    'During development, it was helpful to verify sort correctness (since I had to write my own sort function).
    ' You can verify correctness with a simple check like this:
    'For i = 0 To m_numNearestOffsets - 1
    '    If (i > 0) Then
    '        If (m_NearestOffsets(i).x * m_NearestOffsets(i).x + m_NearestOffsets(i).y * m_NearestOffsets(i).y < m_NearestOffsets(i - 1).x * m_NearestOffsets(i - 1).x + m_NearestOffsets(i - 1).y * m_NearestOffsets(i - 1).y) Then Debug.Print "BAD SORT AT " & i
    '    End If
    'Next i
    
End Sub

Private Sub SortPointsRemaining()
    
    'Don't waste time sorting arbitrarily small lists
    If (m_numPointsLeft <= 4) Then Exit Sub
    
    'Calculate the origin of the current shape.  (Rather than a complex marching-squares approach
    ' to the "which pixels are outermost" question, we just sort by distance from the centroid of
    ' the entire rectangle.  I have no idea if this is how GIMP does it, but some other inpainting
    ' algorithms - e.g. Microsoft's at this link...
    ' https://www.microsoft.com/en-us/research/wp-content/uploads/2016/02/tr-2003-84.pdf
    ' ... propose more complex priority rules.  I have not yet pursued a strategy like that.)
    '
    'Note that the caller may have specified their *own* custom center point (PD uses this if the
    ' user turns off some sampling directions), but if the caller *hasn't* specified a custom
    ' center point, we'll just use the center of the image.
    Dim oX As Single, oY As Single
    If (m_SamplingCenter.x = 0) And (m_SamplingCenter.y = 0) Then
        oX = (m_SrcWidth / 2!)
        oY = (m_SrcHeight / 2!)
    Else
        oX = m_SamplingCenter.x
        oY = m_SamplingCenter.y
    End If
    
    'To improve performance, we're not gonna sort the coordinate list in-place.
    ' Instead, we're going to pre-calculate all distances, sort just the distance list,
    ' then rebuild the original coordinate array accordingly.
    Dim sortList() As QSEntry
    ReDim sortList(0 To m_numPointsLeft - 1) As QSEntry
    
    Dim i As Long
    For i = 0 To m_numPointsLeft - 1
        
        sortList(i).idxOriginal = i
        With m_PointsRemaining(i)
            
            'Potential neighbor offsets are sorted by their Euclidean distance from the origin.
            ' (We only need relative values, so don't waste time on the square-root.)
            sortList(i).ptDistance = (.x - oX) * (.x - oX) + (.y - oY) * (.y - oY)
            
        End With
        
    Next i
    
    'Hand the rest of the sort off to a generic coordinate sorter
    OptimizedSort_CoordinateList_Ascending sortList, m_numPointsLeft
    
    'The coordinate list has been successfully sorted.  Clone the original list, then copy each entry
    ' from that temporary clone into its final position in the "official" list (while accounting for
    ' the possibility of either ascending or descending order).
    Dim tmpPoints() As PointLong
    ReDim tmpPoints(0 To m_numPointsLeft - 1) As PointLong
    CopyMemoryStrict VarPtr(tmpPoints(0)), VarPtr(m_PointsRemaining(0)), 8 * m_numPointsLeft
    
    Dim uBoundRemainingPoints As Long
    uBoundRemainingPoints = m_numPointsLeft - 1
    For i = 0 To uBoundRemainingPoints
        If (m_FillOrder = fo_InsideOut) Then
            m_PointsRemaining(i) = tmpPoints(sortList(i).idxOriginal)
        Else
            m_PointsRemaining(uBoundRemainingPoints - i) = tmpPoints(sortList(i).idxOriginal)
        End If
    Next i
    
    'If you want to verify sort order, consult the SortNeighborCoordinates function for a sample function.
    
End Sub

'Sort an arbitrary list of coordinates in ascending order.  It's up to the caller to transform their data
' into QSEntry format, and rearrange their list according to the final order when this function finishes.
' (PhotoDemon uses this approach because different coordinate lists may be sorted by this function.)
Private Sub OptimizedSort_CoordinateList_Ascending(ByRef sortList() As QSEntry, ByVal numPoints As Long)

    'Prep an internal stack (this is faster and more reliable than recursion).  Note that because
    ' we are using a fixed potential list of sort candidates, this upper limit is guaranteed safe
    ' for this application - but it may *not* be safe for all quick sort ops, so if you change things
    ' like the maximum upper limit for search radius, you may need to increase stack size accordingly.
    Const INIT_QUICKSORT_STACK_SIZE As Long = 256
    Dim qsRemaining() As QSStack, qsStackPtr As Long
    ReDim qsRemaining(0 To INIT_QUICKSORT_STACK_SIZE - 1) As QSStack
    qsStackPtr = 0
    qsRemaining(0).sLB = 0
    qsRemaining(0).sUB = numPoints - 1
    
    'Like most quick-sort implementations, VB6 sees a perf boost by switching to insertion sort
    ' when a quicksort sub-list is small.  This threshold could probably be refined after more
    ' intensive study, but the usual default of 10 shows a ~20-30% perf improvement over baseline.
    Const USE_INSERTION_SORT_THRESHOLD As Long = 10
    
    Dim idxLow As Long, idxHigh As Long
    Dim i As Long, j As Long, v As QSEntry, vDist As Long
    
    'Head of the quick sort loop
    Do
        
        'Load the next set of sub-list boundaries
        idxLow = qsRemaining(qsStackPtr).sLB
        idxHigh = qsRemaining(qsStackPtr).sUB
        
        'Switch over to insertion sort when the sub-list is small
        If (idxHigh - idxLow < USE_INSERTION_SORT_THRESHOLD) Then
            
            i = idxLow + 1
            Do While (i <= idxHigh)
                v = sortList(i)
                vDist = v.ptDistance
                j = i - 1
                Do While (sortList(j).ptDistance > vDist)
                    sortList(j + 1) = sortList(j)
                    j = j - 1
                    If (j < idxLow) Then Exit Do    'VB6 doesn't short-circuit so this check must be performed separately
                Loop
                sortList(j + 1) = v
                i = i + 1
            Loop
        
        'On larger lists, continue with standard quicksort
        Else
            
            'Establish initial pivots
            i = idxLow
            j = idxHigh
            vDist = sortList((i + j) \ 2).ptDistance
            
            'Shift this item into position
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
    
End Sub

'Given a source image (32-bpp only) and mask (a byte array with the SAME X/Y DIMENSIONS as the image),
' replace any masked pixels in the source image with "texture synthesized" results taken from nearby
' unmasked pixels.
'
'The mask array is also treated as a const.  The source image *will* be overwritten anywhere that
' the mask array has a non-zero value.  (Pixels where the mask has a 0 value are guaranteed unchanged.)
'
'Please note that the synthesizer does not blend pixels, so mask values from 1 to 255 are treated
' the same. It is up to the caller to blend the results if desired (for e.g. feathered selections).
'
'Originally, this function allowed the selection mask to be smaller than the source image, and a rect
' needed to be passed to define the mask boundary.  However, this had large perf implications because
' *every* mask access needed to be manually bounds-checked.  By forcing the mask to be the same size
' as the source image, we can skip boundary checking for a meaningful performance boost.
'
'If the passed reportProgress parameter is TRUE, progress bar updates will be supplied to the main window.
'
'Returns TRUE if the operations completed successfully, FALSE otherwise.  FALSE can mean several things
' (user canceled the result, there were no pixels marked to fill, etc), but regardless of the underlying
' reason you must ignore the results of srcDIB when FALSE is returned, as its state is not guaranteed.
' (It is STRONGLY RECOMMENDED to maintain a backup copy of srcDIB in case the user cancels, so you can
' restore their original image.)
Friend Function ContentAwareFill(ByRef srcDIB As pdDIB, ByRef srcMask() As Byte, Optional ByVal reportProgress As Boolean = True) As Boolean
    
    m_ShowProgressUI = reportProgress
    
    'To test with default values, uncomment the block below.
    'm_MaxNumNeighbors = MAX_NEIGHBORS_DEFAULT
    'm_MaxRandomCandidates = RANDOM_CANDIDATES_DEFAULT
    'm_Refinement = REFINEMENT_DEFAULT
    'm_AllowOutliers = ALLOW_OUTLIERS_DEFAULT
    'm_ComparatorRadius = COMPARE_RADIUS_DEFAULT
    
    'Unsafely wrap a class-level array around the user mask
    Dim maskSA2D As SafeArray2D
    With maskSA2D
        .cbElements = 1
        .cDims = 2
        .cLocks = 1
        .Bounds(0).lBound = 0
        .Bounds(0).cElements = srcDIB.GetDIBHeight
        .Bounds(1).lBound = 0
        .Bounds(1).cElements = srcDIB.GetDIBWidth
        .pvData = VarPtr(srcMask(0, 0))
    End With
    PutMem4 VarPtrArray(m_MaskBytes()), VarPtr(maskSA2D)
    
    'See if we can get away with 24-bpp RGB comparisons only (ignoring alpha)
    m_EvaluateRGBOnly = (Not DIBs.IsDIBTransparent(srcDIB))
    
    'Wrap class-level arrays around the image itself
    m_SrcWidth = srcDIB.GetDIBWidth
    m_SrcHeight = srcDIB.GetDIBHeight
    srcDIB.WrapRGBQuadArrayAroundDIB m_SrcPixels, m_SrcSA
    
    'Perform filter
    ContentAwareFill = SynthesizeTexture()
    
    'Regardless of success or failure, we *must* unalias all array wrappers before exiting.
    PutMem4 VarPtrArray(m_MaskBytes), 0&
    srcDIB.UnwrapRGBQuadArrayFromDIB m_SrcPixels
    
End Function

'At initialization, all parameters will be set to their default values.  You can change any of these values
' prior to calling the main SynthesizeTexture() function.
Friend Sub SetDefaultParameterValues()
    
    'How many neighbors to compare when evaluating the quality of a synthesized pixel.  30 is the default value
    ' given in the original resynthesizer paper.  It's probably overkill for most applications.
    '
    'Because neighbors will always be searched in nearest-to-furthest order, changing this doesn't change the
    ' quality of pixel matching, per se - it just controls how far outward the algorithm evaluates during the
    ' EvaluatePixel_RGBA process.
    m_MaxNumNeighbors = MAX_NEIGHBORS_DEFAULT
    
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
    m_MaxRandomCandidates = RANDOM_CANDIDATES_DEFAULT
    
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
    m_Refinement = REFINEMENT_DEFAULT
    
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
    m_AllowOutliers = ALLOW_OUTLIERS_DEFAULT
    
    'Because the original paper is primarily focused on full texture synthesis (where you create a novel
    ' seamless texture from some source), it doesn't deal a lot with specifics for content-aware filling.
    ' Instead, it shows some examples and sort of says "the same concepts work here, too!"
    '
    'One of the things that quickly became apparent during my own implementation attempt was that you
    ' don't want content-aware fill pulling pixels from too far away in the image.  Otherwise, you may end
    ' up mistakenly replacing something like sky with something like water - because they're *close enough*
    ' in color to confuse the algorithm.
    '
    'So I added an option to restrict the radius where the synthesizer looks for replacement pixels.
    ' Note that this value is based on the size of the underlying mask, so if you have three tiny fill
    ' areas spread across an image, it is the union rect of *all* those areas that is used for the
    ' underlying radius calculation.  (Basically, any pixel within [n] pixels of the boundary rect is OK
    ' to use.)  As such, you're better off replacing individual objects one-at-a-time rather than selecting
    ' them all and then batch-replacing them.
    m_ComparatorRadius = COMPARE_RADIUS_DEFAULT
    
    'For "inside-out" and "outside-in" fill modes, this class needs to know where the center of the
    ' sampling region lies.  Rather than force this class to try and solve that itself, the caller can
    ' simply tell us (via .SetSamplingCenter).  Pass [0, 0] to that function to use the default center
    ' point at [width / 2, height / 2].
    m_SamplingCenter.x = 0
    m_SamplingCenter.y = 0
    
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
