VERSION 5.00
Begin VB.Form FormRedEye 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Red eye removal"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   12030
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   436
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   802
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.pdSlider sltShapeStrictness 
      Height          =   675
      Left            =   6480
      TabIndex        =   6
      Top             =   2880
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   1191
      Caption         =   "strictness"
      FontSizeCaption =   10
      Min             =   1
      Max             =   100
      SigDigits       =   1
      SliderTrackStyle=   1
      Value           =   50
      GradientColorRight=   1703935
      NotchPosition   =   2
      NotchValueCustom=   50
   End
   Begin PhotoDemon.pdCheckBox chkShape 
      Height          =   375
      Left            =   6000
      TabIndex        =   4
      Top             =   2400
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   661
      Caption         =   "enforce shape restrictions"
   End
   Begin PhotoDemon.pdSlider sltColor 
      Height          =   705
      Left            =   6000
      TabIndex        =   2
      Top             =   720
      Width           =   5880
      _ExtentX        =   10372
      _ExtentY        =   1270
      Caption         =   "color sensitivity"
      Min             =   1
      Max             =   200
      SigDigits       =   1
      Value           =   100
      NotchPosition   =   2
      NotchValueCustom=   100
   End
   Begin PhotoDemon.pdFxPreviewCtl pdFxPreview 
      Height          =   5625
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9922
   End
   Begin PhotoDemon.pdCommandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5790
      Width           =   12030
      _ExtentX        =   21220
      _ExtentY        =   1323
   End
   Begin PhotoDemon.pdSlider sltObject 
      Height          =   705
      Left            =   6000
      TabIndex        =   3
      Top             =   1560
      Width           =   5880
      _ExtentX        =   10372
      _ExtentY        =   1270
      Caption         =   "object sensitivity"
      Min             =   1
      Max             =   200
      SigDigits       =   1
      Value           =   100
      NotchPosition   =   2
      NotchValueCustom=   100
   End
   Begin PhotoDemon.pdCheckBox chkHighlight 
      Height          =   375
      Left            =   6000
      TabIndex        =   5
      Top             =   5040
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   661
      Caption         =   "highlight detected regions (preview only)"
   End
   Begin PhotoDemon.pdCheckBox chkSize 
      Height          =   375
      Left            =   6000
      TabIndex        =   7
      Top             =   3720
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   661
      Caption         =   "enforce size restrictions"
   End
   Begin PhotoDemon.pdSlider sltSizeStrictness 
      Height          =   675
      Left            =   6480
      TabIndex        =   8
      Top             =   4200
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   1191
      Caption         =   "strictness"
      FontSizeCaption =   10
      Min             =   1
      Max             =   100
      SigDigits       =   1
      SliderTrackStyle=   1
      Value           =   50
      GradientColorRight=   1703935
      NotchPosition   =   2
      NotchValueCustom=   50
   End
End
Attribute VB_Name = "FormRedEye"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Automated Red Eye Correction Tool
'Copyright 2015-2017 by Tanner Helland
'Created: 29/December/15
'Last updated: 29/December/15
'Last update: initial build
'
'Comments TODO
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'During debug, you can render debug data onto the final image by setting this to TRUE.
Private Const RENDER_DEBUG_REDEYE_DATA As Boolean = False

'Apply automated red-eye correction
Public Sub ApplyRedEyeCorrection(ByVal parameterList As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)
    
    'Parse out the parameter list
    Dim cParams As pdParamXML
    Set cParams = New pdParamXML
    cParams.SetParamString parameterList
    
    Dim colorSensitivity As Double, objectSensitivity As Double
    colorSensitivity = cParams.GetDouble("color-sensitivity", 100#)
    objectSensitivity = cParams.GetDouble("object-sensitivity", 100#)
    
    'Passed sensitivity values are on the range [0, 200].  Normalize these to [-0.1, 0.1] and [-0.5, 0.5], respectively.
    colorSensitivity = (colorSensitivity - 100#) / 1000#
    objectSensitivity = (objectSensitivity - 100#) / 200#
    
    'This function can restrict red-eye regions by "shape".  (In this case, shape refers purely to aspect ratio.)
    ' A higher strictness requires a tighter aspect ratio, but note that "confirmShape" can be FALSE, in which case no
    ' aspect ratio enforcement takes place.
    Dim confirmShape As Boolean, shapeStrictness As Double
    confirmShape = cParams.GetBool("confirm-shape", True)
    shapeStrictness = cParams.GetDouble("shape-strictness", 50#)
    
    'Aspect ratio strictness enters on the range [1, 100], default = 50.
    ' We want the calculation default to be 2.5, on the range [1.1, 4.1], with HIGH strictness yielding the LOWEST value.
    shapeStrictness = (shapeStrictness / 100#) * 3#
    shapeStrictness = 1.1 + (3 - shapeStrictness)
    
    'This function can also restrict red-eye regions by "size".  I've tried using relative size (e.g. eye size relative
    ' to total image size), but that gets tricky if the function is applied to a selection, as the eye may be enormous
    ' relative to the selection area.  So instead, this value operates purely on pixel measurements.
    Dim confirmSize As Boolean, sizeStrictness As Double
    confirmSize = cParams.GetBool("confirm-size", True)
    sizeStrictness = cParams.GetDouble("size-strictness", 50#)
    
    'Since we use area as our primary comparison, square the incoming "size strictness" value
    sizeStrictness = 20 + (100 - sizeStrictness)
    sizeStrictness = sizeStrictness * sizeStrictness
    If toPreview Then sizeStrictness = sizeStrictness * curDIBValues.previewModifier
    
    'While in preview mode, this function can highlight the red-eye regions that have been detected.  Regardless of value,
    ' this setting has no meaning when not in preview mode.
    Dim previewHighlight As Boolean
    previewHighlight = cParams.GetBool("preview-highlight", True)
    
    If Not toPreview Then Message "Searching image for potential red-eye locations..."
    
    'Create a local array and point it at the pixel data we want to operate on
    Dim ImageData() As Byte
    Dim tmpSA As SAFEARRAY2D
    
    PrepImageData tmpSA, toPreview, dstPic
    CopyMemory ByVal VarPtrArray(ImageData()), VarPtr(tmpSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left
    initY = curDIBValues.Top
    finalX = curDIBValues.Right
    finalY = curDIBValues.Bottom
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim quickX As Long, qvDepth As Long
    qvDepth = curDIBValues.BytesPerPixel
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    If Not toPreview Then
        SetProgBarMax 3
        SetProgBarVal 0
    End If
    
    'Color and grayscale variables
    Dim r As Long, g As Long, b As Long
    Dim rRatio As Double, gRatio As Double, bRatio As Double, pxSum As Double
    
    'We need an array the size of the image to track various pixel statistics.  Each pixel will be sorted into a variety
    ' of potential categories, and because we're applying region analysis to the image, we need to gather statistical data
    ' on large numbers of pixels at a time.
    Dim redEyeData() As Byte
    ReDim redEyeData(initX To finalX, initY To finalY) As Byte
    
    'For large segments of our heuristics, we're only going to be referring to the red channel in the image.
    ' By stripping out red and green bytes, we can reduce memory access times and cache competition.
    Dim redMap() As Byte
    ReDim redMap(initX To finalX, initY To finalY) As Byte
    
    'A few constants to make this code easier to read.  We use a lot of "magic numbers" during red-eye analysis, alas.
    Const PIXEL_IS_NON_SKIN As Long = 1
    Const PIXEL_IS_MOSTLY_RED As Long = 2
    Const PIXEL_IS_INTERIOR_HIGHLIGHT As Long = 3
    
    'Determine cut-off values for valid red-eye pixels.  These start as magic numbers, but they can be modified according
    ' to the "color-sensitivity" parameter passed to the function.
    ' (The magic numbers come from this paper: http://research.microsoft.com/en-us/um/people/leizhang/paper/icip04-lei.pdf)
    Const RED_CUTOFF As Long = 50
    Const RED_RATIO_CUTOFF As Single = 0.4
    Const GREEN_RATIO_CUTOFF As Single = 0.31
    Const BLUE_RATIO_CUTOFF As Single = 0.36
    
    Dim rCutoff As Long, rRatioCutoff As Single, gRatioCutoff As Single, bRatioCutoff As Single
    rCutoff = RED_CUTOFF
    rRatioCutoff = RED_RATIO_CUTOFF + (RED_RATIO_CUTOFF * colorSensitivity)
    gRatioCutoff = GREEN_RATIO_CUTOFF - (GREEN_RATIO_CUTOFF * colorSensitivity)
    bRatioCutoff = BLUE_RATIO_CUTOFF - (BLUE_RATIO_CUTOFF * colorSensitivity)
    
    'Start with a basic red-eye analysis heuristic.  In this step, we simply want to mark "red" pixels.  This initial
    ' data set will then be sorted into "red regions", and because we pre-check redness, we can perform our region
    ' analysis much more quickly.
    For y = initY To finalY
    For x = initX To finalX
        quickX = x * qvDepth
    
        'Get the source pixel color values
        b = ImageData(quickX, y)
        g = ImageData(quickX + 1, y)
        r = ImageData(quickX + 2, y)
        
        'Strip red bytes into a separate tracking array
        redMap(x, y) = r
        
        'Calculate relative RGB sums
        pxSum = r + g + b
        If pxSum <> 0 Then
        
            rRatio = r / pxSum
            gRatio = g / pxSum
            bRatio = b / pxSum
        
            'Compare against our predetermined cutoff values.
            If r > rCutoff Then
                If rRatio > rRatioCutoff Then
                    If gRatio < gRatioCutoff Then
                        If bRatio < bRatioCutoff Then
                            redEyeData(x, y) = PIXEL_IS_MOSTLY_RED
                        End If
                    End If
                End If
            End If
            
            'If this is a non-red pixel, see if we can mark it as non-skin.  This allows us to completely bypass the
            ' pixel on subsequent heuristic passes.
            If redEyeData(x, y) <> PIXEL_IS_MOSTLY_RED Then
                If gRatio > 0.4 Then
                    redEyeData(x, y) = PIXEL_IS_NON_SKIN
                ElseIf bRatio > 0.45 Then
                    redEyeData(x, y) = PIXEL_IS_NON_SKIN
                End If
            End If
            
        End If
        
    Next x
        If Not toPreview Then
            If (y And progBarCheck) = 0 Then
                If UserPressedESC() Then Exit For
            End If
        End If
    Next y
    
    'With a redness map generated, we are now going to apply a second pass to the image, using our redness data as
    ' one of our inputs.  The goal of this step is to mark "highlight" pixels.
    
    'Because we are performing neighborhood searches, and red eyes are unlikely to appear exactly on image borders,
    ' we can shrink our processing area to save some time and resources.
    Dim hlInitX As Long, hlInitY As Long, hlFinalX As Long, hlFinalY As Long
    hlInitX = initX + 2
    hlInitY = initY + 2
    hlFinalX = finalX - 3
    hlFinalY = finalY - 3
    
    Dim hTotal As Long, sTotal As Long, rTotal As Long
    Dim i As Long, j As Long, k As Long
    
    For y = hlInitY To hlFinalY
    For x = hlInitX To hlFinalX
        
        'Apply a basic shadow mask to this pixel; the goal here is to attempt to flag "highlight" pixels in the
        ' center of a red-eye region.  By checking for highlight regions surrounded by red regions, we can greatly
        ' reduce the occurence of false-positives.
        
        'Code blocks here are grouped by row; six rows in total are processed for each pixel.
        sTotal = redMap(x - 1, y - 2)
        sTotal = sTotal + redMap(x, y - 2)
        sTotal = sTotal + redMap(x + 1, y - 2)
        sTotal = sTotal + redMap(x + 2, y - 2)
            
        sTotal = sTotal + redMap(x - 1, y - 1)
        sTotal = sTotal + redMap(x + 2, y - 1)
            
        sTotal = sTotal + redMap(x - 1, y)
        hTotal = redMap(x, y)
        hTotal = hTotal + redMap(x + 1, y)
        sTotal = sTotal + redMap(x + 2, y)
            
        sTotal = sTotal + redMap(x - 1, y + 1)
        hTotal = hTotal + redMap(x, y + 1)
        hTotal = hTotal + redMap(x + 1, y + 1)
        sTotal = sTotal + redMap(x + 2, y + 1)
            
        sTotal = sTotal + redMap(x - 1, y + 2)
        sTotal = sTotal + redMap(x, y + 2)
        sTotal = sTotal + redMap(x + 1, y + 2)
        sTotal = sTotal + redMap(x + 2, y + 2)
        
        'If the highlight vs shadow ratio is acceptable, continue processing this pixel.  Note that the original MS paper
        ' strangely says "> 140" which is an astronomical difference, and one that never results in actual regions
        ' being found.  14 seems to be a good compromise between accuracy and false-positive potential, so I'm assuming
        ' their original 140 value was just a typo.
        If ((hTotal \ 4) - (sTotal \ 16)) > 14 Then
            
            'Count the number of "red" pixels in this sub-region.  To be a true "highlight" pixel, there must be
            ' at least ten red pixels in the subregion
            rTotal = 0
            
            For j = y - 2 To y + 3
            For i = x - 2 To x + 3
                If redEyeData(i, j) = PIXEL_IS_MOSTLY_RED Then rTotal = rTotal + 1
            Next i
            Next j
            
            'Ignore subregions that have 10 or less red pixels
            If rTotal > 10 Then
            
                'There are a good amount of red pixels in this subregion.  Mark it as a potential highlight.
                redEyeData(x, y) = PIXEL_IS_INTERIOR_HIGHLIGHT
                
            End If
            
        End If
        
    Next x
        If Not toPreview Then
            If (y And progBarCheck) = 0 Then
                If UserPressedESC() Then Exit For
            End If
        End If
    Next y
    
    'With potential red-eye, eye-highlight, and non-skin regions identified, it is now time to sort the highlights
    ' into contiguous regions.  Each region will be assessed in turn, and we'll try to remove as many false-positives
    ' as we can.
    
    If Not toPreview Then
        Message "Refining list of eye candidates..."
        SetProgBarVal 1
    End If
    
    'A dedicated "red-eye" class helps with this step.  It's basically an optimized region detector, with some
    ' optimizations applied against this dedicated use-case.
    Dim cRedEye As pdRedEye
    Set cRedEye = New pdRedEye
    
    'The red-eye class requires two input arrays: one is a byte array that contains our various red-eye pixel IDs
    ' (e.g. red, highlight, non-skin, etc).  The other input array is a "Region ID" array, currently of Integer type.
    ' This array will mark each pixel with a region ID > 0, IFF the pixel belongs to a potential red-eye region.
    Dim regionIDs() As Integer
    ReDim regionIDs(initX To finalX, initY To finalY) As Integer
    
    Dim iWidth As Long, iHeight As Long
    iWidth = finalX - initX
    iHeight = finalY - initY
    
    cRedEye.InitializeRedEyeEngine iWidth, iHeight, redEyeData, regionIDs
    
    'We're now going to use a floodfill-like algorithm to generate highlight pixel regions.  This happens in two steps.
    
    'In this function, we are going to scan the redEyeData() array and look for pixels that meet two criteria:
    ' 1) Highlight pixels...
    ' 2) ...that have not yet been added to a valid region.
    
    'When such pixels are found, we'll pass them to the red-eye class.  It will generate region IDs for the all pixels
    ' touching the passed pixel, and also add a region descriptor (position and bounds) to an ever-growing region stack.
    For y = initY To finalY
    For x = initX To finalX
        
        'Is this pixel a highlight pixel?
        If redEyeData(x, y) = PIXEL_IS_INTERIOR_HIGHLIGHT Then
        
            'Does it not yet belong to a region?
            If regionIDs(x, y) = 0 Then
            
                'Let the red-eye handler generate a new contiguous region, starting with this pixel
                cRedEye.FindHighlightRegion x, y, PIXEL_IS_INTERIOR_HIGHLIGHT
            
            End If
        
        End If
        
    Next x
        If Not toPreview Then
            If (y And progBarCheck) = 0 Then
                If UserPressedESC() Then Exit For
            End If
        End If
    Next y
    
    If Not toPreview Then
        Message "Applying final false-positive checks..."
        SetProgBarVal 2
    End If
    
    'All potential highlight regions have now been detected.  Retrieve a copy of the region stack from the red-eye class.
    Dim regionStack() As PD_Dynamic_Region, numOfRegions As Long
    If cRedEye.GetCopyOfRegionStack(regionStack, numOfRegions) Then
    
        'At least one candidate red-eye highlight region exists in the target image.
        
        'Next, we're going to try and remove as many false-positive regions as we can.  We use multiple criteria to
        ' determine whether regions are invalid; some of these are also modified by user inputs to the function.
        Dim regID As Long
        Dim rSum As Long, gSum As Long, bSum As Long, rgbSum As Long
        Dim avePctR As Double, avePctG As Double
        Dim aveR As Long, aveG As Long, aveB As Long, aveL As Long
        Dim numSimilar As Long, numNotInRegion As Long, simThreshold As Long, similarityThresholdReached As Boolean
        Dim numRegionTotal As Long, numRegionRed As Long, numValidRegions As Long, numRegionsProcessed As Long
        Dim aspectRatio As Double, simRejectThreshold As Single
        Const REGION_EXPANSION_RADIUS As Long = 12
        Const DEFAULT_SIMILARITY_THRESHOLD As Single = 0.1
        
        numValidRegions = 0
        numRegionsProcessed = 0
        
        'The rejection threshold for "pixels too similar to their surroundings" is modified by the user's
        ' "object sensitivity" parameter.
        simRejectThreshold = DEFAULT_SIMILARITY_THRESHOLD - (DEFAULT_SIMILARITY_THRESHOLD * objectSensitivity)
        
        'Loop through all highlight regions and attempt to discard regions where pixels surrounding the region bare
        ' strong color similarity to the region itself.  This step is crucial for removing false-positive regions
        ' caused by red-eye-like patterns in clothing and surrounding scenery.
        For i = 0 To numOfRegions - 1
            
            'First, we're going to calculate a few different average color metrics for this region.  These provide
            ' a nice, quick-to-calculate benchmark for assessing region color validity.
            rSum = 0&
            gSum = 0&
            bSum = 0&
            
            avePctR = 0#
            avePctG = 0#
            aveR = 0&
            
            numSimilar = 0
            numNotInRegion = 0
            similarityThresholdReached = False
            
            With regionStack(i)
                
                regID = .RegionID
                
                hlInitX = .RegionLeft
                hlInitY = .RegionTop
                hlFinalX = .RegionLeft + .RegionWidth
                hlFinalY = .RegionTop + .RegionHeight
                
                For y = hlInitY To hlFinalY
                For x = hlInitX To hlFinalX
                    
                    'For this initial step, we're only checking pixels that are actually PART of the region
                    If regionIDs(x, y) = regID Then
                        
                        'Generating a running sum the original RGB values of each pixel in the region
                        quickX = x * qvDepth
                        bSum = bSum + ImageData(quickX, y)
                        gSum = gSum + ImageData(quickX + 1, y)
                        rSum = rSum + ImageData(quickX + 2, y)
                        
                    Else
                        numNotInRegion = numNotInRegion + 1
                    End If
                    
                Next x
                Next y
                
                'Calculate averages for all pixels lying within this region
                avePctR = (CDbl(rSum) / CDbl(rSum + gSum + bSum))
                avePctG = (CDbl(gSum) / CDbl(rSum + gSum + bSum))
                aveR = CDbl(rSum) / .RegionPixelCount
                
                'Expand the region to include a few extra pixels from the surrounding area
                hlInitX = hlInitX - REGION_EXPANSION_RADIUS
                hlFinalX = hlFinalX + REGION_EXPANSION_RADIUS
                hlInitY = hlInitY - REGION_EXPANSION_RADIUS
                hlFinalY = hlFinalY + REGION_EXPANSION_RADIUS
                
                If hlInitX < 0 Then hlInitX = 0
                If hlFinalX > finalX Then hlFinalX = finalX
                If hlInitY < 0 Then hlInitY = 0
                If hlFinalY > finalY Then hlFinalY = finalY
                
                'Update the "neighboring-but-not-inside-region" pixel count
                numNotInRegion = numNotInRegion + (hlFinalX - hlInitX) * (.RegionTop - hlInitY)
                numNotInRegion = numNotInRegion + (hlFinalX - hlInitX) * (hlFinalY - (.RegionTop + .RegionHeight))
                numNotInRegion = numNotInRegion + (.RegionHeight * (.RegionLeft - hlInitX))
                numNotInRegion = numNotInRegion + (.RegionHeight * (hlFinalX - (.RegionLeft + .RegionWidth)))
                
                'Calculate a dynamic "matching-but-not-in-region" value based on the size of the scanned region
                simThreshold = CDbl(numNotInRegion) * simRejectThreshold
                
                For y = hlInitY To hlFinalY
                For x = hlInitX To hlFinalX
                    
                    'If this pixel is NOT part of the region, perform a similarity check between it and our average
                    ' region RGB values.
                    If regionIDs(x, y) <> regID Then
                        
                        'If this pixel is highly similar to its neighboring region, add it to a running tally
                        quickX = x * qvDepth
                        b = ImageData(quickX, y)
                        g = ImageData(quickX + 1, y)
                        r = ImageData(quickX + 2, y)
                        rgbSum = b + g + r
                        If rgbSum = 0 Then rgbSum = 1
                        
                        If Abs(CDbl(r / rgbSum) - avePctR) < 0.03 Then
                            If Abs(CDbl(g / rgbSum) - avePctG) < 0.03 Then
                                If Abs(r - aveR) < 20 Then
                                    numSimilar = numSimilar + 1
                                    If numSimilar > simThreshold Then similarityThresholdReached = True
                                End If
                            End If
                        End If
                        
                    End If
                    
                    'If we've already found too many similarity pixels, exit the loop immediately
                    If similarityThresholdReached Then Exit For
                    
                Next x
                    If similarityThresholdReached Then Exit For
                Next y
                
                'If the similarity threshold was exceeded, mark this region as invalid
                If similarityThresholdReached Then .RegionValid = False
                
                'If this region is still valid, perform some naive failsafe checks on things like aspect ratio and size
                
                'Reject single-pixel regions
                If .RegionValid Then
                    If (.RegionHeight <= 1) Or (.RegionWidth <= 1) Then .RegionValid = False
                End If
                
                'Check size.  The final size check actually happens after combining red and highlight regions into
                ' a single union, but if the current highlight region is ALREADY too large, we can skip that step.
                If .RegionValid And confirmSize Then
                    aspectRatio = (.RegionHeight * .RegionWidth)
                    If aspectRatio > sizeStrictness Then .RegionValid = False
                End If
                
                'Increment the valid region counter
                If .RegionValid Then numValidRegions = numValidRegions + 1
                
            End With
        
        'We've successfully validated this region.  Move to the next one.
        Next i
        
        'With all regions validated, we now need to merge the red-eye data with the highlight data.  This is accomplished
        ' by dynamically "growing" the highlight regions by any neighboring red-eye pixels; the end result is a list of
        ' regions that cover both highlight and red-eye pixels.
        
        If Not toPreview Then
            SetProgBarMax 3 + numValidRegions
            SetProgBarVal 3
        End If
        
        'Start by resetting the region ID array
        For y = initY To finalY
        For x = initX To finalX
            regionIDs(x, y) = 0
        Next x
        Next y
        
        Dim correctionFactor As Double
        Dim regionRectF As RECTF
        
        Dim innerInitX As Long, innerInitY As Long, innerFinalX As Long, innerFinalY As Long
        
        'Next, iterate through all valid regions.
        For i = 0 To numOfRegions - 1
            If regionStack(i).RegionValid Then
                
                If Not toPreview Then
                    numRegionsProcessed = numRegionsProcessed + 1
                    Message "Applying final red-eye corrections (%1 of %2)...", numRegionsProcessed, numValidRegions
                    SetProgBarVal 3 + numRegionsProcessed
                End If
                
                'Tell the red-eye class to expand this region to include any neighboring red-eye pixels.
                cRedEye.ExpandToIncludeRedEye regionStack(i), PIXEL_IS_INTERIOR_HIGHLIGHT, PIXEL_IS_MOSTLY_RED
                
                'regionStack(i) now describes an updated region that includes the red-eye pixels surrounding the
                ' original highlight region.  We're almost ready to correct the region - first, however, we want to
                ' perform a failsafe check of the region's shape and size.  If the new, combined highlight and
                ' red-eye and highlight region is too large, we should reject it prior to actually applying the
                ' redness correction algorithm.
                
                'Check aspect ratio.  Valid regions should be roughly square-shaped.
                If regionStack(i).RegionValid And confirmShape Then
                    If regionStack(i).RegionHeight > regionStack(i).RegionWidth Then
                        aspectRatio = CDbl(regionStack(i).RegionHeight) / CDbl(regionStack(i).RegionWidth)
                    Else
                        aspectRatio = CDbl(regionStack(i).RegionWidth) / CDbl(regionStack(i).RegionHeight)
                    End If
                    If aspectRatio > shapeStrictness Then regionStack(i).RegionValid = False
                End If
                
                'Check size.  Valid regions should not be too large.
                If regionStack(i).RegionValid And confirmSize Then
                    aspectRatio = (regionStack(i).RegionHeight * regionStack(i).RegionWidth)
                    If aspectRatio > sizeStrictness Then regionStack(i).RegionValid = False
                End If
                
                'If the region has survived all attempts to invalidate it, it is finally time to apply the
                ' redness correction part of the algorithm.  Start by applying some heursitics to the region.
                ' The data we gather will increase our odds of successfully reconstructing the target subject's
                ' original eye-color, which may be partially visible outside the pupil.
                If regionStack(i).RegionValid Then
                    
                    'First, calculate average RGB values for the region
                    aveR = 0
                    aveG = 0
                    aveB = 0
                    aveL = 0
                    numRegionTotal = 0
                    numRegionRed = 0
                    
                    With regionStack(i)
                        regID = .RegionID
                        hlInitX = .RegionLeft - 1
                        hlInitY = .RegionTop - 1
                        hlFinalX = .RegionLeft + .RegionWidth + 1
                        hlFinalY = .RegionTop + .RegionHeight + 1
                    End With
                    
                    If hlInitX < initX Then hlInitX = initX
                    If hlInitY < initY Then hlInitY = initY
                    If hlFinalX > finalX Then hlFinalX = finalX
                    If hlFinalY > finalY Then hlFinalY = finalY
                    
                    For y = hlInitY To hlFinalY
                    For x = hlInitX To hlFinalX
                        
                        'Make sure the pixel actually belongs to this region
                        If (regionIDs(x, y) = regID) Then
                        
                            quickX = x * qvDepth
                            b = ImageData(quickX, y)
                            g = ImageData(quickX + 1, y)
                            r = ImageData(quickX + 2, y)
                            
                            'Calculate running luminance for the ENTIRE region (including the eye highlight)
                            aveL = aveL + Colors.GetHQLuminance(r, g, b)
                            numRegionTotal = numRegionTotal + 1
                            
                            'Perform a modified red-eye check.  This steps is where we assess the actual "redness" of the
                            ' underlying pixel; we don't want to correct pixels unless they are obviously red.  (We will
                            ' use these averages to determine how much red to strip out of the red-eye region.)
                            rgbSum = r + g + b
                            If rgbSum = 0 Then rgbSum = 1
                            If CDbl(r / rgbSum) > 0.35 Then
                                aveR = aveR + r
                                aveG = aveG + g
                                aveB = aveB + b
                                numRegionRed = numRegionRed + 1
                            End If
                            
                        End If
                    Next x
                    Next y
                    
                    'With averages successfully detected, we can now (FINALLY) apply actual red-eye correction.
                    If numRegionTotal = 0 Then numRegionTotal = 1
                    If numRegionRed = 0 Then numRegionRed = 1
                    
                    aveL = aveL \ numRegionTotal
                    aveR = aveR \ numRegionRed
                    aveG = aveG \ numRegionRed
                    aveB = aveB \ numRegionRed
                    
                    'Calculate correction factors specific to this region, based on its overall "redness"
                    If aveG > aveB Then
                        correctionFactor = (aveR - aveG) / 255 * 3.2
                    Else
                        correctionFactor = (aveR - aveB) / 255 * 3.2
                    End If
                    
                    'Loop through all pixels and apply the correction results
                    For y = hlInitY To hlFinalY
                    For x = hlInitX To hlFinalX
                        
                        'For pixels inside the region, correction is easy: reduce the redness and contrast, using our
                        ' best guess of what the area's original color was.
                        If (regionIDs(x, y) = regID) Then
                            
                            quickX = x * qvDepth
                            b = ImageData(quickX, y)
                            g = ImageData(quickX + 1, y)
                            r = ImageData(quickX + 2, y)
                            
                            r = cRedEye.FixRedEyeColor(r, -correctionFactor, -0.1, aveL)
                            g = cRedEye.FixRedEyeColor(g, correctionFactor / 3, -0.1, aveL)
                            b = cRedEye.FixRedEyeColor(b, correctionFactor / 3, -0.1, aveL)
                            
                            'We now want to alias this result against neighboring pixels, to try and reduce any harsh
                            ' edges between this pixel and the result.
                            rSum = 0
                            gSum = 0
                            bSum = 0
                            numSimilar = 0
                            
                            innerInitX = x - 1
                            innerInitY = y - 1
                            innerFinalX = x + 1
                            innerFinalY = y + 1
                            
                            If innerInitX < initX Then innerInitX = initX
                            If innerInitY < initY Then innerInitY = initY
                            If innerFinalX > finalX Then innerFinalX = finalX
                            If innerFinalY > finalY Then innerFinalY = finalY
                            
                            For j = innerInitY To innerFinalY
                            For k = innerInitX To innerFinalX
                                If k <> j Then
                                    If regionIDs(k, j) <> regID Then
                                        bSum = bSum + ImageData(k * qvDepth, j)
                                        gSum = gSum + ImageData(k * qvDepth + 1, j)
                                        rSum = rSum + ImageData(k * qvDepth + 2, j)
                                        numSimilar = numSimilar + 1
                                    End If
                                End If
                            Next k
                            Next j
                            
                            'If at least one non-red pixel was found, use its value to soften the correction result.
                            If numSimilar > 0 Then
                                r = Colors.BlendColors(r, rSum \ numSimilar, numSimilar / 8)
                                g = Colors.BlendColors(g, gSum \ numSimilar, numSimilar / 8)
                                b = Colors.BlendColors(b, bSum \ numSimilar, numSimilar / 8)
                            End If
                            
                            ImageData(quickX, y) = b
                            ImageData(quickX + 1, y) = g
                            ImageData(quickX + 2, y) = r
                            
                        End If
                        
                    Next x
                    Next y
                    
                    'Take one final pass through the region.  In this last step, we will fade any non-region pixels against their
                    ' "inside-region" pixel neighbors, as a final attempt to smoothen out the edges of the corrected region.
                    For y = hlInitY To hlFinalY
                    For x = hlInitX To hlFinalX
                        
                        If (regionIDs(x, y) <> regID) Then
                            
                            quickX = x * qvDepth
                            b = ImageData(quickX, y)
                            g = ImageData(quickX + 1, y)
                            r = ImageData(quickX + 2, y)
                            
                            'We now want to alias this result against neighboring pixels, to try and reduce any harsh
                            ' edges between this pixel and the result.
                            rSum = 0
                            gSum = 0
                            bSum = 0
                            numSimilar = 0
                            
                            innerInitX = x - 1
                            innerInitY = y - 1
                            innerFinalX = x + 1
                            innerFinalY = y + 1
                            
                            If innerInitX < initX Then innerInitX = initX
                            If innerInitY < initY Then innerInitY = initY
                            If innerFinalX > finalX Then innerFinalX = finalX
                            If innerFinalY > finalY Then innerFinalY = finalY
                            
                            For j = innerInitY To innerFinalY
                            For k = innerInitX To innerFinalX
                                If k <> j Then
                                    If regionIDs(k, j) = regID Then
                                        bSum = bSum + ImageData(k * qvDepth, j)
                                        gSum = gSum + ImageData(k * qvDepth + 1, j)
                                        rSum = rSum + ImageData(k * qvDepth + 2, j)
                                        numSimilar = numSimilar + 1
                                    End If
                                End If
                            Next k
                            Next j
                            
                            'If at least one non-red pixel was found, use its value to soften the correction result.
                            If numSimilar > 0 Then
                                r = Colors.BlendColors(r, rSum \ numSimilar, numSimilar / 8)
                                g = Colors.BlendColors(g, gSum \ numSimilar, numSimilar / 8)
                                b = Colors.BlendColors(b, bSum \ numSimilar, numSimilar / 8)
                            
                                ImageData(quickX, y) = b
                                ImageData(quickX + 1, y) = g
                                ImageData(quickX + 2, y) = r
                            End If
                            
                        End If
                        
                    Next x
                    Next y
                    
                    'If this is a preview, mark the eyes with a region highlight
                    If toPreview And previewHighlight Then
                        regionRectF.Left = regionStack(i).RegionLeft - 1
                        regionRectF.Top = regionStack(i).RegionTop - 1
                        regionRectF.Width = regionStack(i).RegionWidth + 2
                        regionRectF.Height = regionStack(i).RegionHeight + 2
                        GDI_Plus.GDIPlusDrawCanvasRectF workingDIB.GetDIBDC, regionRectF
                    End If
                
                'End final validity check
                End If
            
            'End initial validity check
            End If
        
        'We've successfully processed this region.  Move to the next one.
        Next i
        
    End If
    
    
    'DEBUG ONLY!  If the form-level debug constant is active, paint the detected regions onto the image, so we can
    ' see how well our classifier algorithms worked.
    If RENDER_DEBUG_REDEYE_DATA Then
        
        For y = initY To finalY
        For x = initX To finalX
            quickX = x * qvDepth
            b = ImageData(quickX, y)
            g = ImageData(quickX + 1, y)
            r = ImageData(quickX + 2, y)
            
            If redEyeData(x, y) = PIXEL_IS_MOSTLY_RED Then
                ImageData(quickX, y) = 0
                ImageData(quickX + 1, y) = 0
                ImageData(quickX + 2, y) = 255
            ElseIf redEyeData(x, y) = PIXEL_IS_INTERIOR_HIGHLIGHT Then
                ImageData(quickX, y) = 0
                ImageData(quickX + 1, y) = 255
                ImageData(quickX + 2, y) = 0
            End If
            
        Next x
        Next y
        
    End If
    
    'Release the red-eye engine.  (This is extremely important, as the red-eye class unsafely aliases multiple local arrays.)
    cRedEye.ReleaseRedEyeEngine redEyeData, regionIDs
    
    'With our work complete, point ImageData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(ImageData), 0&, 4
    Erase ImageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    FinalizeImageData toPreview, dstPic

End Sub

Private Sub chkHighlight_Click()
    UpdatePreview
End Sub

Private Sub chkShape_Click()
    UpdatePreview
End Sub

Private Sub chkSize_Click()
    UpdatePreview
End Sub

Private Sub cmdBar_OKClick()
    Process "Red-eye removal", , GetLocalParamString(), UNDO_LAYER
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub cmdBar_ResetClick()
    sltColor.Value = 100#
    sltObject.Value = 100#
    sltShapeStrictness.Value = 50#
    sltSizeStrictness.Value = 50#
End Sub

Private Sub Form_Load()
    cmdBar.MarkPreviewStatus False
    ApplyThemeAndTranslations Me
    cmdBar.MarkPreviewStatus True
    UpdatePreview
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub pdFxPreview_ViewportChanged()
    UpdatePreview
End Sub

Private Sub UpdatePreview()
    If cmdBar.PreviewsAllowed Then Me.ApplyRedEyeCorrection GetLocalParamString(), True, pdFxPreview
End Sub

Private Function GetLocalParamString() As String
    GetLocalParamString = BuildParamList("color-sensitivity", sltColor.Value, "object-sensitivity", sltObject.Value, "confirm-shape", CBool(chkShape.Value), "shape-strictness", sltShapeStrictness.Value, "confirm-size", CBool(chkSize.Value), "size-strictness", sltSizeStrictness.Value, "preview-highlight", CBool(chkHighlight.Value))
End Function

Private Sub sltColor_Change()
    UpdatePreview
End Sub

Private Sub sltObject_Change()
    UpdatePreview
End Sub

Private Sub sltShapeStrictness_Change()
    UpdatePreview
End Sub

Private Sub sltSizeStrictness_Change()
    UpdatePreview
End Sub
