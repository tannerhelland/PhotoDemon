VERSION 5.00
Begin VB.Form FormLensCorrect 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   " Correct existing distortion"
   ClientHeight    =   6690
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
   ScaleHeight     =   446
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   806
   Begin PhotoDemon.pdCommandBar cmdBar 
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5940
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
   Begin PhotoDemon.pdButtonStrip btsOptions 
      Height          =   960
      Left            =   6000
      TabIndex        =   2
      Top             =   120
      Width           =   5835
      _ExtentX        =   10292
      _ExtentY        =   1693
      Caption         =   "correction type"
   End
   Begin PhotoDemon.pdContainer pnlMode 
      Height          =   4575
      Index           =   1
      Left            =   5880
      Top             =   1200
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   8070
      Begin PhotoDemon.pdSlider sldAdvanced 
         Height          =   705
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   1320
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   1244
         Caption         =   "(a) edges"
         Min             =   -5
         Max             =   5
         SigDigits       =   3
      End
      Begin PhotoDemon.pdSlider sldAdvanced 
         Height          =   705
         Index           =   1
         Left            =   120
         TabIndex        =   11
         Top             =   2040
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   1244
         Caption         =   "(b) midpoints"
         Min             =   -5
         Max             =   5
         SigDigits       =   3
      End
      Begin PhotoDemon.pdSlider sldAdvanced 
         Height          =   705
         Index           =   2
         Left            =   3120
         TabIndex        =   12
         Top             =   1320
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   1244
         Caption         =   "(c) whole image"
         Min             =   -5
         Max             =   5
         SigDigits       =   3
      End
      Begin PhotoDemon.pdSlider sldAdvanced 
         Height          =   705
         Index           =   3
         Left            =   3120
         TabIndex        =   13
         Top             =   2040
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   1244
         Caption         =   "(d) zoom"
         Min             =   -5
         Max             =   5
         SigDigits       =   3
      End
      Begin PhotoDemon.pdSlider sltQuality 
         Height          =   705
         Index           =   1
         Left            =   120
         TabIndex        =   14
         Top             =   2880
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   1270
         Caption         =   "quality"
         Min             =   1
         Max             =   3
         Value           =   2
         NotchPosition   =   2
         NotchValueCustom=   2
      End
      Begin PhotoDemon.pdDropDown cboEdges 
         Height          =   735
         Index           =   1
         Left            =   120
         TabIndex        =   15
         Top             =   3720
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   1296
         Caption         =   "if pixels lie outside the corrected area"
      End
      Begin PhotoDemon.pdSlider sltXCenter 
         Height          =   405
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   873
         Max             =   1
         SigDigits       =   3
         Value           =   0.5
         NotchPosition   =   2
         NotchValueCustom=   0.5
      End
      Begin PhotoDemon.pdSlider sltYCenter 
         Height          =   405
         Left            =   3120
         TabIndex        =   4
         Top             =   480
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   873
         Max             =   1
         SigDigits       =   3
         Value           =   0.5
         NotchPosition   =   2
         NotchValueCustom=   0.5
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   330
         Index           =   0
         Left            =   120
         Top             =   120
         Width           =   5925
         _ExtentX        =   10451
         _ExtentY        =   582
         Caption         =   "center position (x, y)"
         FontSize        =   12
         ForeColor       =   4210752
      End
      Begin PhotoDemon.pdLabel lblExplanation 
         Height          =   435
         Index           =   0
         Left            =   240
         Top             =   960
         Width           =   5655
         _ExtentX        =   0
         _ExtentY        =   0
         Alignment       =   2
         Caption         =   "you can also set a center position by clicking the preview window"
         FontSize        =   9
         ForeColor       =   4210752
         Layout          =   1
      End
   End
   Begin PhotoDemon.pdContainer pnlMode 
      Height          =   4575
      Index           =   0
      Left            =   5880
      Top             =   1200
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   8070
      Begin PhotoDemon.pdSlider sltStrength 
         Height          =   705
         Left            =   120
         TabIndex        =   5
         Top             =   0
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   1270
         Caption         =   "correction strength"
         Max             =   20
         SigDigits       =   2
         Value           =   3
         DefaultValue    =   3
      End
      Begin PhotoDemon.pdSlider sltZoom 
         Height          =   705
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   1270
         Caption         =   "correction zoom"
         Min             =   1
         Max             =   3
         SigDigits       =   2
         Value           =   1.5
         NotchPosition   =   2
         NotchValueCustom=   1
         DefaultValue    =   1.5
      End
      Begin PhotoDemon.pdSlider sltRadius 
         Height          =   705
         Left            =   120
         TabIndex        =   7
         Top             =   1680
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   1270
         Caption         =   "radius (percentage)"
         Min             =   1
         Max             =   100
         Value           =   100
         NotchPosition   =   2
         NotchValueCustom=   100
      End
      Begin PhotoDemon.pdSlider sltQuality 
         Height          =   705
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   2520
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   1270
         Caption         =   "quality"
         Min             =   1
         Max             =   3
         Value           =   2
         NotchPosition   =   2
         NotchValueCustom=   2
      End
      Begin PhotoDemon.pdDropDown cboEdges 
         Height          =   735
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   3360
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   1296
         Caption         =   "if pixels lie outside the corrected area"
      End
   End
End
Attribute VB_Name = "FormLensCorrect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Fix Lens Distort Tool
'Copyright 2013-2026 by Tanner Helland
'Created: 22/January/13
'Last updated: 21/February/20
'Last update: large performance improvements
'
'This tool allows the user to correct an existing lens distortion on an image.  Bilinear interpolation
' (via reverse-mapping) is available for a higher quality correction.
'
'A zoom parameter is also provided to help the user determine how much of the image they are willing
' to sacrifice as part of the correction.  If the distort is quite high, there is no real way to
' correct the image without cutting off parts of it (see sample images at http://photo.net/learn/fisheye/).
'
'For optimal quality, I suggest zooming out a ways, applying the correction, then cropping the resultant
' image to the desired shape.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

Private Sub btsOptions_Click(ByVal buttonIndex As Long)
    UpdatePreview
    UpdateOptionsPanel
End Sub

Private Sub UpdateOptionsPanel()
    Dim i As Long
    For i = pnlMode.lBound To pnlMode.UBound
        pnlMode(i).Visible = (i = btsOptions.ListIndex)
    Next i
End Sub

Private Sub cboEdges_Click(Index As Integer)
    UpdatePreview
End Sub

Public Sub CorrectLensDistortion(ByVal effectParameters As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)

    'This form supports two different lens distortion correction models.  Parse out the model required, and forward
    ' the request to the appropriate destination function.
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString effectParameters
    
    If (cParams.GetLong("lenscorrect_model", 0) = 0) Then
        ApplyLensCorrection_Basic effectParameters, toPreview, dstPic
    Else
        ApplyLensCorrection_Advanced effectParameters, toPreview, dstPic
    End If
    
End Sub

'Correct lens distortion in an image using a full-featured, multi-parameter model
Public Sub ApplyLensCorrection_Advanced(ByVal effectParameters As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)

    If (Not toPreview) Then Message "Correcting image distortion..."
    
    'Parse out individual effect parameters
    Dim paramA As Double, paramB As Double, paramC As Double, paramD As Double, edgeHandling As Long, superSamplingAmount As Long
    Dim centerX As Double, centerY As Double
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString effectParameters
    
    With cParams
        paramA = .GetDouble("lensadvanced_a", 0#)
        paramB = .GetDouble("lensadvanced_b", 0#)
        paramC = .GetDouble("lensadvanced_c", 0#)
        paramD = .GetDouble("lensadvanced_d", 1#)
        edgeHandling = .GetLong("lensbasic_edgepixels", 0)
        superSamplingAmount = .GetLong("lensbasic_quality", 2)
        centerX = .GetDouble("lensadvanced_x", 0.5)
        centerY = .GetDouble("lensadvanced_y", 0.5)
    End With
    
    'For now, auto-calculate d
    paramD = 2# ^ paramD
    
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
    fSupport.SetDistortParameters edgeHandling, (superSamplingAmount <> 1), curDIBValues.maxX, curDIBValues.maxY
    
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
    
    'Lens distort correction requires a number of specialized variables
    
    'Calculate the center of the image
    Dim midX As Double, midY As Double
    midX = CDbl(finalX - initX) * centerX
    midX = midX + initX
    midY = CDbl(finalY - initY) * centerY
    midY = midY + initY
    
    'Rotation values
    Dim theta As Double, sRadius As Double, sRadius2 As Double, sDistance As Double
    Dim radius As Double, rSrc As Double
    
    'X and Y values, remapped around a center point of (0, 0)
    Dim nX As Double, nY As Double
    
    'Source X and Y values, which may or may not be used as part of a bilinear interpolation function
    Dim srcX As Double, srcY As Double
    
    'Max radius is calculated as the distance from the center of the image to a corner
    Dim tWidth As Long, tHeight As Long
    tWidth = curDIBValues.Width
    tHeight = curDIBValues.Height
    sRadius = Sqr(tWidth * tWidth + tHeight * tHeight)
    
    Dim refDistance As Double, invRefDistance As Double
    refDistance = sRadius
    If (refDistance <> 0#) Then invRefDistance = 1# / refDistance
    
    'Lens radius could theoretically be user-controlled; at present, we default to the size
    ' of the underlying image (as the UI is already overflowing with toggles).
    Dim lensRadius As Double
    lensRadius = 100#
    sRadius = sRadius * (lensRadius / 100#)
    sRadius2 = sRadius * sRadius
    
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
        
        'Remap the coordinates around a center point of (0, 0)
        nX = x - midX
        nY = y - midY
        
        'Sample a number of source pixels corresponding to the user's supplied quality value; more quality means
        ' more samples, and much better representation in the final output.
        For sampleIndex = 0 To numSamples
            
            'Offset the pixel amount by the supersampling lookup table
            j = nX + ssX(sampleIndex)
            k = nY + ssY(sampleIndex)
            
            'Calculate distance automatically
            sDistance = (j * j) + (k * k)
            
            'Only pixels within the user-specified radius are addressed
            If (sDistance <= sRadius2) Then
                
                'Calculate a normalized radius and angle
                sDistance = Sqr(sDistance)
                radius = sDistance * invRefDistance
                theta = PDMath.Atan2(k, j)
                
                'Calculate a new radius, using the distortion correction parameters to modify
                rSrc = (paramA * (radius * radius * radius) + paramB * (radius * radius) + paramC * (radius) + paramD) * radius
                
                'Un-normalize the newly calculated radius
                radius = rSrc * refDistance
                
                'Convert them back to the Cartesian plane
                srcX = midX + (radius * Cos(theta))
                srcY = midY + (radius * Sin(theta))
                
            Else
                srcX = x
                srcY = y
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

'Correct lens distortion in an image using a simplifed model
Public Sub ApplyLensCorrection_Basic(ByVal effectParameters As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)
    
    If (Not toPreview) Then Message "Correcting image distortion..."
    
    'Parse out individual effect parameters
    Dim fixStrength As Double, fixZoom As Double, lensRadius As Double, edgeHandling As Long, superSamplingAmount As Long
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString effectParameters
    
    With cParams
        fixStrength = .GetDouble("lensbasic_strength", 0#)
        fixZoom = .GetDouble("lensbasic_zoom", 1#)
        lensRadius = .GetDouble("lensbasic_radius", 100#)
        edgeHandling = .GetLong("lensbasic_edgepixels", 0)
        superSamplingAmount = .GetLong("lensbasic_quality", 2)
    End With
    
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
    fSupport.SetDistortParameters edgeHandling, (superSamplingAmount <> 1), curDIBValues.maxX, curDIBValues.maxY
    
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
    
    'Lens distort correction requires a number of specialized variables
    
    'Calculate the center of the image
    Dim midX As Double, midY As Double
    midX = CDbl(finalX - initX) / 2
    midX = midX + initX
    midY = CDbl(finalY - initY) / 2
    midY = midY + initY
    
    'Rotation values
    Dim theta As Double, sRadius As Double, sRadius2 As Double, sDistance As Double
    Dim radius As Double
    
    'X and Y values, remapped around a center point of (0, 0)
    Dim nX As Double, nY As Double
    
    'Source X and Y values, which may or may not be used as part of a bilinear interpolation function
    Dim srcX As Double, srcY As Double
        
    'Max radius is calculated as the distance from the center of the image to a corner
    Dim tWidth As Long, tHeight As Long
    tWidth = curDIBValues.Width
    tHeight = curDIBValues.Height
    sRadius = Sqr(tWidth * tWidth + tHeight * tHeight) / 2
              
    Dim refDistance As Double
    If fixStrength = 0 Then fixStrength = 0.00000001
    refDistance = sRadius * 2 / fixStrength
                  
    sRadius = sRadius * (lensRadius / 100#)
    sRadius2 = sRadius * sRadius
    
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
        
        'Remap the coordinates around a center point of (0, 0)
        nX = x - midX
        nY = y - midY
        
        'Sample a number of source pixels corresponding to the user's supplied quality value; more quality means
        ' more samples, and much better representation in the final output.
        For sampleIndex = 0 To numSamples
            
            'Offset the pixel amount by the supersampling lookup table
            j = nX + ssX(sampleIndex)
            k = nY + ssY(sampleIndex)
            
            'Calculate distance automatically
            sDistance = (j * j) + (k * k)
            
            If (sDistance <= sRadius2) Then
                
                sDistance = Sqr(sDistance)
                radius = sDistance / refDistance
                
                If (radius = 0#) Then theta = 1# Else theta = Atn(radius) / radius
                srcX = midX + theta * j * fixZoom
                srcY = midY + theta * k * fixZoom
                
            Else
                srcX = x
                srcY = y
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

Private Sub cmdBar_OKClick()
    Process "Correct lens distortion", , GetEffectParams(), UNDO_Layer
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub cmdBar_ResetClick()
    cboEdges(0).ListIndex = pdeo_Clamp
    cboEdges(1).ListIndex = pdeo_Clamp
End Sub

Private Sub Form_Load()

    'Disable previews until all controls have been initialized
    cmdBar.SetPreviewStatus False
    
    btsOptions.AddItem "basic", 0
    btsOptions.AddItem "advanced", 1
    btsOptions.ListIndex = 0
    UpdateOptionsPanel
    
    'I use a central function to populate the edge handling combo box; this way, I can add new methods and have
    ' them immediately available to all distort functions.
    PopDistortEdgeBox cboEdges(0), pdeo_Clamp
    PopDistortEdgeBox cboEdges(1), pdeo_Clamp
    
    ApplyThemeAndTranslations Me, True, True
    cmdBar.SetPreviewStatus True
    UpdatePreview

End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub pdFxPreview_PointSelected(xRatio As Double, yRatio As Double)
    cmdBar.SetPreviewStatus False
    sltXCenter.Value = xRatio
    sltYCenter.Value = yRatio
    cmdBar.SetPreviewStatus True
    UpdatePreview
End Sub

Private Sub sldAdvanced_Change(Index As Integer)
    UpdatePreview
End Sub

Private Sub sltQuality_Change(Index As Integer)
    UpdatePreview
End Sub

Private Sub sltRadius_Change()
    UpdatePreview
End Sub

Private Sub sltStrength_Change()
    UpdatePreview
End Sub

Private Sub sltXCenter_Change()
    UpdatePreview
End Sub

Private Sub sltYCenter_Change()
    UpdatePreview
End Sub

Private Sub sltZoom_Change()
    UpdatePreview
End Sub

'Redraw the on-screen preview of the transformed image
Private Sub UpdatePreview()
    If cmdBar.PreviewsAllowed Then CorrectLensDistortion GetEffectParams(), True, pdFxPreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub pdFxPreview_ViewportChanged()
    UpdatePreview
End Sub

Private Function GetEffectParams() As String
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    
    With cParams
        .AddParam "lenscorrect_model", btsOptions.ListIndex
        .AddParam "lensbasic_strength", sltStrength.Value
        .AddParam "lensbasic_zoom", sltZoom.Value
        .AddParam "lensbasic_radius", sltRadius.Value
        .AddParam "lensbasic_quality", sltQuality(btsOptions.ListIndex).Value
        .AddParam "lensbasic_edgepixels", cboEdges(btsOptions.ListIndex).ListIndex
        .AddParam "lensadvanced_a", sldAdvanced(0).Value
        .AddParam "lensadvanced_b", sldAdvanced(1).Value
        .AddParam "lensadvanced_c", sldAdvanced(2).Value
        .AddParam "lensadvanced_d", sldAdvanced(3).Value
        .AddParam "lensadvanced_x", sltXCenter.Value
        .AddParam "lensadvanced_y", sltYCenter.Value
    End With
    
    GetEffectParams = cParams.GetParamString()
    
End Function
