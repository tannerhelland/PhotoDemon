VERSION 5.00
Begin VB.Form FormLens 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   " Apply lens distortion"
   ClientHeight    =   6540
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12090
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
   ScaleWidth      =   806
   Begin PhotoDemon.pdCommandBar cmdBar 
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5790
      Width           =   12090
      _ExtentX        =   21325
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
      DisableZoomPan  =   -1  'True
      PointSelection  =   -1  'True
   End
   Begin PhotoDemon.pdSlider sltRadius 
      Height          =   705
      Left            =   6000
      TabIndex        =   2
      Top             =   2640
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
      Caption         =   "radius"
      Min             =   1
      Max             =   100
      Value           =   50
      NotchPosition   =   2
      NotchValueCustom=   50
   End
   Begin PhotoDemon.pdSlider sldCurvature 
      Height          =   705
      Left            =   6000
      TabIndex        =   3
      Top             =   1800
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
      Caption         =   "curvature"
      Max             =   5
      SigDigits       =   2
   End
   Begin PhotoDemon.pdSlider sltXCenter 
      Height          =   405
      Left            =   6000
      TabIndex        =   4
      Top             =   720
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   873
      Max             =   1
      SigDigits       =   2
      Value           =   0.5
      NotchPosition   =   2
      NotchValueCustom=   0.5
   End
   Begin PhotoDemon.pdSlider sltYCenter 
      Height          =   405
      Left            =   9000
      TabIndex        =   5
      Top             =   720
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   873
      Max             =   1
      SigDigits       =   2
      Value           =   0.5
      NotchPosition   =   2
      NotchValueCustom=   0.5
   End
   Begin PhotoDemon.pdSlider sltQuality 
      Height          =   705
      Left            =   6000
      TabIndex        =   6
      Top             =   4320
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
      Caption         =   "quality"
      Min             =   1
      Max             =   5
      Value           =   2
      NotchPosition   =   2
      NotchValueCustom=   2
   End
   Begin PhotoDemon.pdLabel lblExplanation 
      Height          =   435
      Index           =   0
      Left            =   6120
      Top             =   1200
      Width           =   5655
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "you can also set a center position by clicking the preview window"
      ForeColor       =   4210752
      Layout          =   1
   End
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   285
      Index           =   0
      Left            =   6000
      Top             =   360
      Width           =   5925
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "center position (x, y)"
      FontSize        =   12
      ForeColor       =   4210752
   End
   Begin PhotoDemon.pdSlider sldShape 
      Height          =   705
      Left            =   6000
      TabIndex        =   7
      Top             =   3480
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
      Caption         =   "aspect ratio"
      Min             =   -100
      Max             =   100
      SigDigits       =   1
   End
End
Attribute VB_Name = "FormLens"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Lens Correction and Distortion
'Copyright 2013-2026 by Tanner Helland
'Created: 05/January/13
'Last updated: 21/February/20
'Last update: large performance improvements
'
'This tool allows the user to apply a lens distortion to an image.  It is comparable to the "Spherize" tool
' in PhotoShop.  (For correcting lens distortion, please see FormLensCorrect.)
'
'As of January '14, the user can now set a custom center point by clicking the image or using the x/y sliders.
'
'This transformation is a modified version of a transformation originally written by Jerry Huxtable of JH Labs.
' Jerry's original code is licensed under an Apache 2.0 license.  You may download his original version at the
' following link (good as of 07 January '13): http://www.jhlabs.com/ip/filters/index.html
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Apply a new lens distortion to an image
Public Sub ApplyLensDistortion(ByVal effectParams As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString effectParams
    
    Dim refractiveIndex As Double, lensRadius As Double, lensShape As Double, centerX As Double, centerY As Double
    Dim superSamplingAmount As Long
    
    With cParams
        refractiveIndex = .GetDouble("strength", sldCurvature.Value)
        lensRadius = .GetDouble("radius", sltRadius.Value)
        lensShape = .GetDouble("shape", sldShape.Value)
        superSamplingAmount = .GetLong("quality", sltQuality.Value)
        centerX = .GetDouble("centerx", 0.5)
        centerY = .GetDouble("centery", 0.5)
    End With
    
    refractiveIndex = 1# / (refractiveIndex + 1#)

    If (Not toPreview) Then Message "Projecting image through simulated lens..."
    
    'Create a local array and point it at the pixel data of the current image
    Dim dstImageData() As Byte, dstSA As SafeArray2D, dstSA1D As SafeArray1D
    EffectPrep.PrepImageData dstSA, toPreview, dstPic
    
    'Create a copy of the current image; we will use it as our source reference.
    Dim srcDIB As pdDIB
    Set srcDIB = New pdDIB
    srcDIB.CreateFromExistingDIB workingDIB
    
    'At present, stride is always width * 4 (32-bit RGBA)
    Dim xStride As Long
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left
    initY = curDIBValues.Top
    finalX = curDIBValues.Right
    finalY = curDIBValues.Bottom
    
    'Create a filter support class, which will aid with edge handling and interpolation
    Dim fSupport As pdFilterSupport
    Set fSupport = New pdFilterSupport
    fSupport.SetDistortParameters pdeo_Erase, (superSamplingAmount <> 1), curDIBValues.maxX, curDIBValues.maxY
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    If (Not toPreview) Then ProgressBars.SetProgBarMax finalY
    progBarCheck = ProgressBars.FindBestProgBarValue()
    
    '***************************************
    ' /* BEGIN SUPERSAMPLING PREPARATION */
    
    'Due to the way this filter works, supersampling yields much better results.  Because supersampling is extremely
    ' energy-intensive, this tool uses a sliding value for quality, as opposed to a binary TRUE/FALSE for antialiasing.
    ' (For all but the lowest quality setting, antialiasing will be used, and higher quality values will simply increase
    '  the amount of supersamples taken.)
    Dim newR As Long, newG As Long, newB As Long, newA As Long
    Dim r As Long, g As Long, b As Long, a As Long
    Dim tmpSum As Long, tmpSumFirst As Long
    
    'Use the passed super-sampling constant (displayed to the user as "quality") to come up with a number of actual
    ' pixels to sample.  (The total amount of sampled pixels will range from 1 to 13).  Note that supersampling
    ' coordinates are precalculated and cached using a modified rotated grid function, which is consistent throughout PD.
    Dim numSamples As Long
    Dim ssX() As Single, ssY() As Single
    Filters_Area.GetSupersamplingTable superSamplingAmount, numSamples, ssX, ssY
    
    'Because supersampling will be used in the inner loop as (samplecount - 1), permanently decrease the sample
    ' count in advance.
    numSamples = numSamples - 1
    
    'Additional variables are needed for supersampling handling
    Dim j As Double, k As Double
    Dim sampleIndex As Long, numSamplesUsed As Long
    Dim superSampleVerify As Long, ssVerificationLimit As Long
    
    'Adaptive supersampling allows us to bypass supersampling if a pixel doesn't appear to benefit from it.  The superSampleVerify
    ' variable controls how many pixels are sampled before we perform an adaptation check.  At present, the rule is:
    ' Quality 3: check a minimum of 2 samples, Quality 4: check minimum 3 samples, Quality 5: check minimum 4 samples
    superSampleVerify = superSamplingAmount - 2
    
    'Alongside a variable number of test samples, adaptive supersampling requires some threshold that indicates samples
    ' are close enough that further supersampling is unlikely to improve output.  We calculate this as a minimum variance
    ' as 1.5 per channel (for a total of 6 variance per pixel), multiplied by the total number of samples taken.
    ssVerificationLimit = superSampleVerify * 6
    
    'To improve performance for quality 1 and 2 (which perform no supersampling), we can forcibly disable supersample checks
    ' by setting the verification checker to some impossible value.
    If (superSampleVerify <= 0) Then superSampleVerify = LONG_MAX
    
    ' /* END SUPERSAMPLING PREPARATION */
    '*************************************
    
    'Lensing requires a collection of specialized variables
    
    'Calculate the center of the image, modified by the user's custom center point inputs
    Dim midX As Double, midY As Double
    midX = CDbl(finalX - initX) * centerX + initX
    midY = CDbl(finalY - initY) * centerY + initY
    
    'Calculation values
    Dim theta As Double, theta2 As Double
    Dim xAngle As Double, yAngle As Double
    Dim lensAngle As Double
    
    'X and Y values, remapped around a center point of (0, 0)
    Dim nX As Double, nY As Double
    Dim nX2 As Double, nY2 As Double
    
    'All PD distort filters use reverse-mapping, which means we loop through pixels in the final image,
    ' and reverse-map their positions back to the original image.
    Dim srcX As Double, srcY As Double
        
    'By default, the lens effect is calculated as a perfect circle, with a radius based off the smaller of
    ' the two image dimensions.
    Dim imgWidth As Long, imgHeight As Long
    imgWidth = curDIBValues.Width
    imgHeight = curDIBValues.Height
    
    Dim minDimension As Double
    If (imgWidth < imgHeight) Then minDimension = imgWidth Else minDimension = imgHeight
    
    'The user's "lens shape" parameter allows us to adjust the aspect ratio of the lens on-the-fly.
    Dim shapeAdjH As Double, shapeAdjV As Double
    If (lensShape = 0#) Then
        shapeAdjH = 0#
        shapeAdjV = 0#
    ElseIf (lensShape > 0#) Then
        shapeAdjH = (minDimension * lensShape) * 0.01
        shapeAdjV = 0#
    Else
        shapeAdjH = 0#
        shapeAdjV = (minDimension * lensShape * -1#) * 0.01
    End If
    
    'We can precalculate a bunch of inner-loop variables to make the supersampling loop faster
    Dim sRadiusW As Double, sRadiusH As Double
    Dim sRadiusW2 As Double, sRadiusH2 As Double, invSRadiusH2 As Double
    sRadiusW = minDimension * (lensRadius * 0.01) + shapeAdjH
    If (sRadiusW <= 0#) Then sRadiusW = 0.0001
    sRadiusW2 = 1# / (sRadiusW * sRadiusW)
    
    sRadiusH = minDimension * (lensRadius * 0.01) + shapeAdjV
    If (sRadiusH <= 0#) Then sRadiusH = 0.0001
    sRadiusH2 = sRadiusH * sRadiusH
    invSRadiusH2 = 1# / sRadiusH2
    
    Dim sRadiusMult As Double
    sRadiusMult = sRadiusW * sRadiusH
    
    Dim tmpQuad As RGBQuad
    fSupport.AliasTargetDIB srcDIB
    
    'Loop through each pixel in the image, converting values as we go
    For y = initY To finalY
        workingDIB.WrapArrayAroundScanline dstImageData, dstSA1D, y
    For x = initX To finalX
       
        'Reset all supersampling values
        newR = 0
        newG = 0
        newB = 0
        newA = 0
        numSamplesUsed = 0
        
        'Remap the coordinates around the user's center point of (0, 0)
        nX = x - midX
        nY = y - midY
        
        'Sample a number of source pixels corresponding to the user's supplied quality value; more quality means
        ' more samples, and much better representation in the final output.
        For sampleIndex = 0 To numSamples
            
            j = nX + ssX(sampleIndex)
            k = nY + ssY(sampleIndex)
            
            nX2 = j * j
            nY2 = k * k
            
            'If the values are going to be out-of-bounds, force them to a set OOB position to ensure that the
            ' resampler uses a blank pixel for them.  (Note that we do this *inside* the supersampling loop,
            ' which allows for smooth antialiasing along the lens boundary at higher quality levels.)
            If (nY2 >= (sRadiusH2 - ((sRadiusH2 * nX2) * sRadiusW2))) Then
                srcX = -1
                srcY = -1
            
            'Otherwise, reverse-map x and y back onto the original image using a reversed lens refraction calculation
            Else
                
                theta = Sqr((1# - (nX2 * sRadiusW2) - (nY2 * invSRadiusH2)) * sRadiusMult)
                theta2 = theta * theta
                
                'Calculate the angle for x
                xAngle = Acos(j / Sqr(nX2 + theta2))
                lensAngle = PI_HALF - xAngle - Asin(Sin(PI_HALF - xAngle) * refractiveIndex)
                srcX = x - Tan(lensAngle) * theta
                
                'Now do the same thing for y
                yAngle = Acos(k / Sqr(nY2 + theta2))
                lensAngle = PI_HALF - yAngle - Asin(Sin(PI_HALF - yAngle) * refractiveIndex)
                srcY = y - Tan(lensAngle) * theta
                
            End If
            
            'Use the filter support class to interpolate and edge-wrap pixels as necessary
            tmpQuad = fSupport.GetColorsFromSource(srcX, srcY, x, y)
            b = tmpQuad.Blue
            g = tmpQuad.Green
            r = tmpQuad.Red
            a = tmpQuad.Alpha
            
            'If adaptive supersampling is active, apply the "adaptive" aspect.  Basically, calculate a variance for the currently
            ' collected samples.  If variance is low, assume this pixel does not require further supersampling.
            ' (Note that this is an ugly shorthand way to calculate variance, but it's fast, and the chance of false outliers is
            '  small enough to make it preferable over a true variance calculation.)
            If (sampleIndex = superSampleVerify) Then
                
                'Calculate variance for the first two pixels (Q3), three pixels (Q4), or four pixels (Q5)
                tmpSum = (r + g + b + a) * superSampleVerify
                tmpSumFirst = newR + newG + newB + newA
                
                'If variance is below 1.5 per channel per pixel, abort further supersampling
                If (Abs(tmpSum - tmpSumFirst) < ssVerificationLimit) Then Exit For
            
            End If
            
            'Increase the sample count
            numSamplesUsed = numSamplesUsed + 1
            
            'Add the retrieved values to our running averages
            newR = newR + r
            newG = newG + g
            newB = newB + b
            newA = newA + a
            
        Next sampleIndex
        
        'Find the average values of all samples, apply to the pixel, and move on!
        If (numSamplesUsed > 1) Then
            newR = newR \ numSamplesUsed
            newG = newG \ numSamplesUsed
            newB = newB \ numSamplesUsed
            newA = newA \ numSamplesUsed
        End If
        
        xStride = x * 4
        dstImageData(xStride) = newB
        dstImageData(xStride + 1) = newG
        dstImageData(xStride + 2) = newR
        dstImageData(xStride + 3) = newA
        
    Next x
        If (Not toPreview) Then
            If (y And progBarCheck) = 0 Then
                If Interface.UserPressedESC() Then Exit For
                SetProgBarVal y
            End If
        End If
    Next y
    
    'Safely deallocate all image arrays
    fSupport.UnaliasTargetDIB
    workingDIB.UnwrapArrayFromDIB dstImageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    EffectPrep.FinalizeImageData toPreview, dstPic
        
End Sub

'OK button
Private Sub cmdBar_OKClick()
    Process "Apply lens distortion", , GetLocalParamString(), UNDO_Layer
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub cmdBar_ResetClick()
    sltRadius.Value = 50
    sltQuality.Value = 2
End Sub

Private Sub Form_Load()
    cmdBar.SetPreviewStatus False
    ApplyThemeAndTranslations Me, True, True
    cmdBar.SetPreviewStatus True
    UpdatePreview
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub sldCurvature_Change()
    UpdatePreview
End Sub

Private Sub sldShape_Change()
    UpdatePreview
End Sub

Private Sub sltQuality_Change()
    UpdatePreview
End Sub

Private Sub sltRadius_Change()
    UpdatePreview
End Sub

'Redraw the on-screen preview of the transformed image
Private Sub UpdatePreview()
    If cmdBar.PreviewsAllowed Then ApplyLensDistortion GetLocalParamString(), True, pdFxPreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub pdFxPreview_ViewportChanged()
    UpdatePreview
End Sub

'The user can right-click the preview area to select a new center point
Private Sub pdFxPreview_PointSelected(xRatio As Double, yRatio As Double)
    cmdBar.SetPreviewStatus False
    sltXCenter.Value = xRatio
    sltYCenter.Value = yRatio
    cmdBar.SetPreviewStatus True
    UpdatePreview
End Sub

Private Sub sltXCenter_Change()
    UpdatePreview
End Sub

Private Sub sltYCenter_Change()
    UpdatePreview
End Sub

Private Function GetLocalParamString() As String
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    
    With cParams
        .AddParam "strength", sldCurvature.Value
        .AddParam "shape", sldShape.Value
        .AddParam "radius", sltRadius.Value
        .AddParam "quality", sltQuality.Value
        .AddParam "centerx", sltXCenter.Value
        .AddParam "centery", sltYCenter.Value
    End With
    
    GetLocalParamString = cParams.GetParamString()
    
End Function
