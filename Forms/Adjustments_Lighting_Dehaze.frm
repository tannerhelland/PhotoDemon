VERSION 5.00
Begin VB.Form FormDehaze 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   " Dehaze"
   ClientHeight    =   6540
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11550
   DrawStyle       =   5  'Transparent
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HasDC           =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   436
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   770
   Begin PhotoDemon.pdCommandBar cmdBar 
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5790
      Width           =   11550
      _ExtentX        =   20373
      _ExtentY        =   1323
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
   Begin PhotoDemon.pdSlider sldBackground 
      Height          =   705
      Left            =   6000
      TabIndex        =   2
      Top             =   1080
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   1244
      Caption         =   "background"
      Min             =   1
      Max             =   100
      Value           =   85
      NotchPosition   =   2
      NotchValueCustom=   85
   End
   Begin PhotoDemon.pdSlider sldThreshold 
      Height          =   705
      Left            =   6000
      TabIndex        =   3
      Top             =   2040
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   1244
      Caption         =   "threshold"
      Min             =   1
      Max             =   100
      Value           =   80
      NotchPosition   =   2
      NotchValueCustom=   80
   End
   Begin PhotoDemon.pdSlider sldForeground 
      Height          =   705
      Left            =   6000
      TabIndex        =   4
      Top             =   3960
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   1244
      Caption         =   "foreground"
      Min             =   1
      Max             =   100
      Value           =   90
      NotchPosition   =   2
      NotchValueCustom=   90
   End
   Begin PhotoDemon.pdSlider sldGamma 
      Height          =   705
      Left            =   6000
      TabIndex        =   5
      Top             =   3000
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   1244
      Caption         =   "gamma"
      Min             =   0.01
      Max             =   3
      SigDigits       =   2
      Value           =   1
      NotchPosition   =   2
      NotchValueCustom=   1
   End
End
Attribute VB_Name = "FormDehaze"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Image Dehazing Tool
'Copyright 2021-2026 by Tanner Helland
'Created: 08/September/21
'Last updated: 09/September/21
'Last update: wrap up initial build
'
'Automatic image dehazing is an area of active study.  It is of special interest to automative manufacturers,
' who need access to fast, high-quality dehazing algorithms for automative safety system cameras.
'
'PD's implementation is currently based on a 2018 paper in the open-access journal "Mathematical Problems in
' Engineering" by Zhong Luan, Hao Zeng, Yuanyuan Shang, Zhuhong Shao, and Hui Ding.  The paper is titled
' "Fast Video Dehazing Using Per-Pixel Minimum Adjustment" and you can you download it here (link good as of
' September 2021): https://www.hindawi.com/journals/mpe/2018/9241629/
'
'Many thanks to Zhong Luan et al for sharing their work under a Creative Commons Attribution License.
'
'I read quite a few dehazing papers from the past two decades before settling on this one for PhotoDemon's
' implementation.  One of the biggest issues I have in PD development is implementing methods that meet a
' few strict criteria:
' 1) no massive 3rd-party libraries (e.g. OpenCV is out) or amd64-only libraries
' 2) good run-time performance without modern intrinsics, including SIMD (necessary for XP compatibility)
' 3) better performance than O(n^2) because considering (1) and (2), I still have to make the function fast
'    enough to be useable on 20+ megapixel photos, even on PCs with strong RAM/CPU limitations.
'
'These criteria get harder to meet every year, especially as I attempt to add things to PD beyond the usual
' low-hanging photo editing fruit of brightness, contrast, etc.
'
'Anyway, I liked this paper because most dehazing approaches rely on either massive regional analysis on each
' pixel (which is incredibly slow), optimizing lengthy differential equations on-the-fly (which is even worse),
' or machine-learning algorithms that require huge training sets (ugh no).  For example, this link describes
' a classical dehaze implementation based on the original dark-channel prior theory by Kaiming He:
' - https://sites.google.com/site/computervisionadinastoica/final-project
' As the authors at that link state: "Running the algorithm on any larger images would require days of
' computation... He, Sun, and Tang mentioned using the Preconditioned Conjugate Gradient algorithm as a solver,
' allowing them to process a 600x400 pixel image in roughly 10-20 seconds."
'
'10-20 seconds for a 600x400 image!  Not gonna work for PhotoDemon.
'
'Convesely, the paper by Zhong Luan et al uses a much-simpler per-pixel analysis with good - and predictable! -
' results, excellent performance, low memory requirements, and many areas of potential further optimization
' that play nicely with VB6's limitations.  Using this strategy, a 10-megapixel image processes in less than a
' second, regardless of inputs.  Much better.
'
'As such, while this work is heavily inspired by the original paper, it is *not* identical.  (Nothing built
' in VB ever is lol.)  I have heavily commented the code to note deliberate modifications and/or enhancements,
' and I am open to input on further changes.  As with other functions in PD, I generally try to err on the side
' of "more subtle, useful real-world results" vs "intense HDR-style changes" that look impressive in a paper
' but produce wildly unrealistic-looking results on things like regular iPhone photos.
'
'Ideas for further improvements are of course welcome.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Apply a "dehaze" filter to an image, including automatic estimation of atmospheric lighting.
Public Sub ApplyDehaze(ByVal effectParams As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)
        
    If (Not toPreview) Then Message "Removing haze from image..."
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString effectParams
    
    'Parse incoming parameters.  Note that default values for epsilon, threshold, and mu
    ' come from the original paper at https://www.hindawi.com/journals/mpe/2018/9241629/#EEq2
    Dim epsilon As Double, threshold As Double, gamma As Double, mu As Double
    epsilon = cParams.GetDouble("background", 85#)
    threshold = cParams.GetDouble("threshold", 80#)
    gamma = cParams.GetDouble("gamma", 1#)
    mu = cParams.GetDouble("foreground", 89#)
    
    'epsilon and mu need to be on the range [0, 1],
    ' but they are presented to the user on [0, 100] (for readability)
    epsilon = epsilon / 100#
    mu = mu / 100#
    
    Const EPSILON_MIN As Double = 0#, EPSILON_MAX As Double = 1#
    If (epsilon < EPSILON_MIN) Then epsilon = EPSILON_MIN
    If (epsilon > EPSILON_MAX) Then epsilon = EPSILON_MAX
    
    Const MU_MIN As Double = 0#, MU_MAX As Double = 1#
    If (mu < MU_MIN) Then mu = MU_MIN
    If (mu > MU_MAX) Then mu = MU_MAX
    
    'gamma is applied using a LUT
    Const GAMMA_MIN As Double = 0.01, GAMMA_MAX As Double = 3#
    If (gamma < GAMMA_MIN) Then gamma = GAMMA_MIN
    If (gamma > GAMMA_MAX) Then gamma = GAMMA_MAX
    
    Dim gammaLUT(0 To 255) As Byte
    Dim gTmp As Double
    
    Dim i As Long
    For i = 0 To 255
        gTmp = CDbl(i) / 255#
        gTmp = gTmp ^ (1# / gamma)
        gTmp = gTmp * 255#
        If (gTmp > 255#) Then gTmp = 255#
        If (gTmp < 0#) Then gTmp = 0#
        gammaLUT(i) = Int(gTmp + 0.5)
    Next i
    
    'Generate a preview copy of the current image (if previews are active; otherwise, this will simply
    ' prep a new scratch copy of the image for editing).
    Dim dstSA As SafeArray2D, tmpSA1D As SafeArray1D, imgPixels() As Byte
    EffectPrep.PrepImageData dstSA, toPreview, dstPic
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left
    initY = curDIBValues.Top
    finalX = curDIBValues.Right
    finalY = curDIBValues.Bottom
    
    'Dehazing works by first estimating atmostpheric light.  A number of approaches are available for this;
    ' in PD, we use the quad-tree method proposed here: https://www.hindawi.com/journals/mpe/2018/9241629/
    ' This approach converges quickly and works well on any size source image.
    Dim atmValue As Single, atmQuad As RGBQuad
    atmValue = EstimateAtmosphericLighting(workingDIB, atmQuad)
    
    'Do not allow atmospheric light to be 0.  (We need to divide by it later, and we don't want DBZ checks.)
    ' Technically it should be near-impossible for atmospheric values to be 0 since the quad-tree converges
    ' on local maxima, but if the user does something dumb like pass an all-black image we won't crash.
    If (atmQuad.Blue < 0) Then atmQuad.Blue = 1
    If (atmQuad.Green < 0) Then atmQuad.Green = 1
    If (atmQuad.Red < 0) Then atmQuad.Red = 1
    
    'We will also be dividing per-pixel color components by these atmospheric values, so precalculate
    ' inverse values to avoid division in the inner loop.
    Dim iAb As Double, iAg As Double, iAr As Double
    iAb = 1# / CDbl(atmQuad.Blue)
    iAg = 1# / CDbl(atmQuad.Green)
    iAr = 1# / CDbl(atmQuad.Red)
    
    'Calculate new atmospheric values on the range [0, 1] for each color component
    Const ONE_DIV_255 As Double = 1# / 255#
    Dim ab As Double, ag As Double, ar As Double
    ab = atmQuad.Blue * ONE_DIV_255
    ag = atmQuad.Green * ONE_DIV_255
    ar = atmQuad.Red * ONE_DIV_255
    
    'Note the minimum atmospheric channel now; we'll use it to optimize the dehaze function later
    Dim aMin As Double
    aMin = PDMath.Min3Float(ab, ag, ar)
    
    'With atmospheric value calculated, we can proceed with dehazing.
    Dim r As Long, g As Long, b As Long
    Dim t As Long, tf As Double, invTf As Double
    Dim dx As Double, aX As Double  'aX is a correction factor called "alpha", but *not* related to the alpha channel
    
    'Progress bar updates are only provided on non-preview applications
    Dim progBarCheck As Long
    If (Not toPreview) Then ProgressBars.SetProgBarMax finalY
    progBarCheck = ProgressBars.FindBestProgBarValue()
    
    For y = initY To finalY
        workingDIB.WrapArrayAroundScanline imgPixels, tmpSA1D, y
    For x = initX To finalX * 4 Step 4
        
        'Retrieve original color values
        b = imgPixels(x)
        g = imgPixels(x + 1)
        r = imgPixels(x + 2)
        
        'Per "Single Image Haze Removal Using Dark Channel Prior" by Kaming He, the haze equation is
        ' simply defined as (for each channel in an image):
        
        ' pxHaze(x, y) = pxOriginal(x, y) * t(x, y) + a * (1 - t(x, y))
        
        ' Where a = global atmospheric light, and t = the medium transmission (roughly thought of as
        ' how thick the haze is, or more accurately, how strongly it obscures the original pixel).
        ' When t = 1, the haze is fully transparent, and the original pixel appears unmodified.
        ' When t = 0, the original pixel is fully obscured, and only the atmospheric light a is visible.
        '
        'Rearranging terms to solve for pxOriginal (which is what we ultimately want to calculate),
        ' we arrive at:
        ' pxOriginal(x, y) = (pxHaze(x, y) - a * (1 - t(x, y))) / t(x, y)
        '
        'Calculating a and t is the messy part of this whole operation.  a has already been estimated in
        ' a previous step.  We will attempt to calculate t for each pixel "as we go".
        
        'Start by establishing the minimum channel in this pixel
        If (b < g) Then
            If (b < r) Then t = b Else t = r
        Else
            If (g < r) Then t = g Else t = r
        End If
        
        'Convert it to the range [0, 1]
        tf = (CDbl(t) * ONE_DIV_255)
        
        'Calculate dX, or the relationship between the minimum of the current pixel and the atmospheric light.
        ' (This is used to mitigate darkening of pixels that experience severe haze correction.)
        dx = Abs(tf - aMin) * 255#
        
        'Use dX to determine correction factor alpha for this pixel
        If (dx > 0#) Then
            aX = Sqr(threshold / dx)
            If (mu > aX) Then aX = mu       'mu is used to establish a "safe" upper bound on correction
        Else
            aX = 1#
        End If
        
        'TESTING ONLY: to ignore this "correction factor", force aX to 1.  This is useful for evaluating
        ' the "pure" algorithmic approach as defined early in the source paper.
        'aX = 1#
        
        'Find the smallest ratio between each color in this pixel and its atmospheric equivalent.  This works
        ' as a control that provides stronger correction to pixels very close to the atmospheric color,
        ' while minimizing changes to colors that differ greatly.
        Dim tmpMin As Double
        tmpMin = PDMath.Min3Float(b * iAb, g * iAg, r * iAr)
        
        'Now solve for t
        tf = (1# - epsilon * tmpMin) * aX
        
        'Because we're going to calculate directly into an int further down, prevent extremely small
        ' values that could cause overflow.
        Const T_MIN As Double = 0.0000001
        If (tf < T_MIN) Then tf = T_MIN
        
        'To avoid multiple divides, divide once and cache the result
        invTf = 1# / tf
        
        'Solve the final equation for each pixel.  Note that this is sort of like a reverse-alpha-blend,
        ' where we are attempting to fade out the atmospheric component from the assumed "original"
        ' pixel color value (the "dark channel prior").  Because this is subtractive, it tends to have
        ' a darkening effect - so we also provide a subsequent user-defined gamma correction.
        b = (b - atmQuad.Blue * (1# - tf)) * invTf
        g = (g - atmQuad.Green * (1# - tf)) * invTf
        r = (r - atmQuad.Red * (1# - tf)) * invTf
        
        'Clamp final values appropriately
        If (b < 0) Then b = 0
        If (b > 255) Then b = 255
        If (g < 0) Then g = 0
        If (g > 255) Then g = 255
        If (r < 0) Then r = 0
        If (r > 255) Then r = 255
        
        'Apply gamma, if any
        b = gammaLUT(b)
        g = gammaLUT(g)
        r = gammaLUT(r)
        
        'Apply the new pixel in-place (e.g. a separate destination image is NOT required)
        imgPixels(x) = b
        imgPixels(x + 1) = g
        imgPixels(x + 2) = r
        
    Next x
        If (Not toPreview) Then
            If (y And progBarCheck) = 0 Then
                If Interface.UserPressedESC() Then Exit For
                ProgressBars.SetProgBarVal y
            End If
        End If
    Next y
    
    'Safely deallocate imageData()
    workingDIB.UnwrapArrayFromDIB imgPixels
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    EffectPrep.FinalizeImageData toPreview, dstPic
        
End Sub

Private Function EstimateAtmosphericLighting(ByRef srcDIB As pdDIB, ByRef dstAtmosphericQuad As RGBQuad) As Single

    'PD's atmospheric light estimator works similar to the quad-tree method proposed here:
    ' https://www.hindawi.com/journals/mpe/2018/9241629/
    ' My implementation of their algorithm is a novel one designed against VB6's particular quirks.
    
    'Start by generating a channel-minimum map of the full image (e.g the smallest value of RGB for
    ' each pixel, regardless of channel).
    Dim minVal() As Byte
    ReDim minVal(0 To srcDIB.GetDIBWidth - 1, 0 To srcDIB.GetDIBHeight - 1) As Byte
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = srcDIB.GetDIBWidth - 1
    finalY = srcDIB.GetDIBHeight - 1
    
    Dim r As Long, g As Long, b As Long, cMin As Long
    
    Dim srcSA As SafeArray1D, srcPixels() As Byte
    For y = 0 To finalY
        srcDIB.WrapArrayAroundScanline srcPixels, srcSA, y
    For x = 0 To finalX
        
        b = srcPixels(x * 4)
        g = srcPixels(x * 4 + 1)
        r = srcPixels(x * 4 + 2)
        
        If (b < g) Then
            If (b < r) Then cMin = b Else cMin = r
        Else
            If (g < r) Then cMin = g Else cMin = r
        End If
        
        minVal(x, y) = cMin
        
    Next x
    Next y
    
    srcDIB.UnwrapArrayFromDIB srcPixels
    
    'With minimum values calculated, we now need to recursively subdivide the image into quadrants.
    ' For each quadrant (in each pass), find the quad with the *largest average value*, then subdivide
    ' that quad and repeat the process until some minimum threshold size is reached.
    ' (In PD's implementation, we stop when *either* the width or height reaches the minimum size.
    ' This could easily be modified, below, if different termination behavior is desired.)
    Dim MinSize As Long
    MinSize = 16
    
    'Start by pre-setting a target rect
    Dim curRect As RectL
    With curRect
        .Left = 0
        .Top = 0
        .Right = finalX
        .Bottom = finalY
    End With
    
    Dim curMean As Single, testMean As Single
    
    'Ensure the initial rect is a valid size; if it *isn't*, immediately return the average value.
    If ((curRect.Bottom - curRect.Top + 1) <= MinSize) Or ((curRect.Right - curRect.Left + 1) < MinSize) Then
        curMean = FindMeanOfQuad(minVal, curRect)
        GoTo SkipLoopEntirely
    End If
    
    'Child rects are precalculated on each pass
    Dim childRects(0 To 3) As RectL
    
    Do
        
'        'TESTING ONLY: draw a rectangle around the target rect; this is fun for seeing where the algorithm converges
'        Dim tmpPen As pd2DPen, tmpSurface As pd2DSurface
'        Drawing2D.QuickCreateSurfaceFromDIB tmpSurface, srcDIB, False
'        Drawing2D.QuickCreateSolidPen tmpPen, 1!, RGB(255, 0, 0)
'        PD2D.DrawRectangleI_FromRectL tmpSurface, tmpPen, curRect
        
        'Always start by ensuring the current rectangle is large enough to sub-divide.  If it doesn't,
        ' we've reached the end of the function.  (Note that, by design, this check must *not* be
        ' triggered by the whole image, because we won't have calculated an average value for the rect yet!
        ' Instead, that is treated as a special case, above, for perf reasons.)
        If ((curRect.Bottom - curRect.Top + 1) <= MinSize) Or ((curRect.Right - curRect.Left + 1) < MinSize) Then
            EstimateAtmosphericLighting = curMean
            Exit Do
        End If
        
        'Subdivide the current area into 4 child areas.  Rects are calculated in the following order:
        ' 0 1
        ' 2 3
        With childRects(0)
            .Left = curRect.Left
            .Right = curRect.Left + (curRect.Right - curRect.Left) \ 2
            .Top = curRect.Top
            .Bottom = curRect.Top + (curRect.Bottom - curRect.Top) \ 2
        End With
        With childRects(1)
            .Left = childRects(0).Right + 1
            .Right = curRect.Right
            .Top = curRect.Top
            .Bottom = childRects(0).Bottom
        End With
        With childRects(2)
            .Left = curRect.Left
            .Right = childRects(0).Right
            .Top = childRects(0).Bottom + 1
            .Bottom = curRect.Bottom
        End With
        With childRects(3)
            .Left = childRects(1).Left
            .Right = curRect.Right
            .Top = childRects(2).Top
            .Bottom = curRect.Bottom
        End With
        
        'Always reset the current maximum mean to an impossible value
        curMean = 0!
        
        'Iterate all sub-rects and cache the largest mean value (and associated rect)
        Dim i As Long
        For i = 0 To 3
            testMean = FindMeanOfQuad(minVal, childRects(i))
            If (testMean > curMean) Then
                curMean = testMean
                curRect = childRects(i)
            End If
        Next i
        
    'Repeat the process on the new target rect!
    Loop

SkipLoopEntirely:

    'curRect now contains the desired target rect.  To estimate atmospheric lighting from this rect,
    ' we want to find the max value of each channel (RGB).
    Dim bMax As Long, gMax As Long, rMax As Long
    
    For y = curRect.Top To curRect.Bottom
        srcDIB.WrapArrayAroundScanline srcPixels, srcSA, y
    For x = curRect.Left To curRect.Right
        
        b = srcPixels(x * 4)
        g = srcPixels(x * 4 + 1)
        r = srcPixels(x * 4 + 2)
        
        If (b > bMax) Then bMax = b
        If (g > gMax) Then gMax = g
        If (r > rMax) Then rMax = r
        
    Next x
    Next y
    
    srcDIB.UnwrapArrayFromDIB srcPixels
    
    'Return the mean value as well (this may or may not be useful to the caller) and the estimated
    ' atmospheric color contribution.
    EstimateAtmosphericLighting = curMean
    dstAtmosphericQuad.Blue = bMax
    dstAtmosphericQuad.Green = gMax
    dstAtmosphericQuad.Red = rMax
    dstAtmosphericQuad.Alpha = 255      'Alpha is not relevant in this calculation; you can supply any arbitrary value
    
End Function

'Given a source byte array and a valid (and it MUST be valid) sub-rect, return the mean value of
' entries inside said rect.
Private Function FindMeanOfQuad(ByRef srcArray() As Byte, ByRef srcRect As RectL) As Single

    Dim x As Long, y As Long, curSum As Long
    For y = srcRect.Top To srcRect.Bottom
    For x = srcRect.Left To srcRect.Right
        curSum = curSum + srcArray(x, y)
    Next x
    Next y
    
    FindMeanOfQuad = CDbl(curSum) / CDbl((srcRect.Right - srcRect.Left + 1) * (srcRect.Bottom - srcRect.Top + 1))

End Function

'OK button
Private Sub cmdBar_OKClick()
    Process "Dehaze", , GetLocalParamString(), UNDO_Layer
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub Form_Load()

    'Disable previews while we initialize everything
    cmdBar.SetPreviewStatus False
    
    'Apply translations and visual themes
    ApplyThemeAndTranslations Me, True, True
    cmdBar.SetPreviewStatus True
    UpdatePreview
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub UpdatePreview()
    If cmdBar.PreviewsAllowed Then ApplyDehaze GetLocalParamString(), True, pdFxPreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub pdFxPreview_ViewportChanged()
    UpdatePreview
End Sub

Private Function GetLocalParamString() As String
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    
    With cParams
        .AddParam "background", sldBackground.Value
        .AddParam "threshold", sldThreshold.Value
        .AddParam "gamma", sldGamma.Value
        .AddParam "foreground", sldForeground.Value
    End With
    
    GetLocalParamString = cParams.GetParamString()
    
End Function

Private Sub sldBackground_Change()
    UpdatePreview
End Sub

Private Sub sldForeground_Change()
    UpdatePreview
End Sub

Private Sub sldGamma_Change()
    UpdatePreview
End Sub

Private Sub sldThreshold_Change()
    UpdatePreview
End Sub
