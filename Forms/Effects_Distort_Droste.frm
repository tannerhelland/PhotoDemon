VERSION 5.00
Begin VB.Form FormDroste 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   " Droste"
   ClientHeight    =   6360
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12105
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
   ScaleHeight     =   424
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   807
   Begin PhotoDemon.pdCheckBox chkOrder 
      Height          =   375
      Left            =   6000
      TabIndex        =   9
      Top             =   3360
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   661
      Caption         =   "swap order"
   End
   Begin PhotoDemon.pdCommandBar cmdBar 
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5610
      Width           =   12105
      _ExtentX        =   21352
      _ExtentY        =   1323
   End
   Begin PhotoDemon.pdFxPreviewCtl pdFxPreview 
      Height          =   5355
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9446
      DisableZoomPan  =   -1  'True
      PointSelection  =   -1  'True
   End
   Begin PhotoDemon.pdSlider sltAngle 
      Height          =   705
      Left            =   6000
      TabIndex        =   3
      Top             =   1560
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
      Caption         =   "angle"
      Min             =   -180
      Max             =   180
      SigDigits       =   1
   End
   Begin PhotoDemon.pdSlider sldRadiusInner 
      Height          =   705
      Left            =   6000
      TabIndex        =   4
      Top             =   2400
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   1244
      Caption         =   "inner radius"
      Min             =   1
      Max             =   100
      SigDigits       =   2
      Value           =   10
      NotchPosition   =   2
      NotchValueCustom=   10
   End
   Begin PhotoDemon.pdSlider sltQuality 
      Height          =   705
      Left            =   6000
      TabIndex        =   5
      Top             =   3840
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
   Begin PhotoDemon.pdSlider sltXCenter 
      Height          =   405
      Left            =   6000
      TabIndex        =   6
      Top             =   480
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
      TabIndex        =   7
      Top             =   480
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   873
      Max             =   1
      SigDigits       =   2
      Value           =   0.5
      NotchPosition   =   2
      NotchValueCustom=   0.5
   End
   Begin PhotoDemon.pdDropDown cboEdges 
      Height          =   735
      Left            =   6000
      TabIndex        =   2
      Top             =   4680
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1296
      Caption         =   "if pixels lie outside the image..."
   End
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   285
      Index           =   0
      Left            =   6000
      Top             =   120
      Width           =   5925
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "center position (x, y)"
      FontSize        =   12
      ForeColor       =   4210752
   End
   Begin PhotoDemon.pdLabel lblExplanation 
      Height          =   435
      Index           =   0
      Left            =   6120
      Top             =   1050
      Width           =   5655
      _ExtentX        =   0
      _ExtentY        =   0
      Alignment       =   2
      Caption         =   "you can also set a center position by clicking the preview window"
      FontSize        =   9
      ForeColor       =   4210752
      Layout          =   1
   End
   Begin PhotoDemon.pdSlider sldRadiusOuter 
      Height          =   705
      Left            =   9000
      TabIndex        =   8
      Top             =   2400
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   1244
      Caption         =   "outer radius"
      Min             =   1
      Max             =   100
      SigDigits       =   1
      Value           =   25
      NotchPosition   =   2
      NotchValueCustom=   25
   End
End
Attribute VB_Name = "FormDroste"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'"Droste" Transform (or distortion? who knows lol)
'Copyright 2021-2026 by Tanner Helland
'Created: 24/August/21
'Last updated: 25/August/21
'Last update: wrap up initial build
'
'The Droste effect is not just a cool effect (think M.C. Escher) - there's also an interesting story behind the name!
' See https://en.wikipedia.org/wiki/Droste_effect for details.
'
'PhotoDemon's version of this tool was inspired by a Paint.NET plugin originally by PJayTycy, with additional
' modifications by toe_head2001.  Unfortunately, neither author provided any sort of license with their work,
' so I'm not sure how best to credit them, but I at least welcome interested users to check out their PDN
' plugin here (link good as of August 2021):
' https://forums.getpaint.net/topic/32240-droste-v11-may-8-2019/
'
'Note that this tool relies heavily on complex number math.  Look in the ComplexNumbers module for
' implementation details, including additional copyright and license information for certain functions.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

Private Sub cboEdges_Click()
    UpdatePreview
End Sub

'Apply a "Droste" effect to an image
Public Sub FxDroste(ByVal effectParams As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)

    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString effectParams
    
    Dim fxAngle As Double, fxInnerRadius As Double, fxOuterRadius As Double
    Dim centerX As Double, centerY As Double
    Dim edgeHandling As Long, superSamplingAmount As Long
    Dim fxSwapOrder As Boolean
    
    With cParams
        fxAngle = .GetDouble("angle", sltAngle.Value)
        fxInnerRadius = .GetDouble("radius-inner", sldRadiusInner.Value)
        fxOuterRadius = .GetDouble("radius-outer", sldRadiusOuter.Value)
        edgeHandling = .GetLong("edges", cboEdges.ListIndex)
        superSamplingAmount = .GetLong("quality", sltQuality.Value)
        fxSwapOrder = .GetBool("swap-order", chkOrder.Value)
        centerX = .GetDouble("centerx", 0.5)
        centerY = .GetDouble("centery", 0.5)
    End With
    
    'Reverse the rotationAngle value so that POSITIVE values indicate CLOCKWISE rotation.
    ' (Also, remap it slightly to make the default behavior roughly axis-aligned; this is a quirk of
    ' this algorithm specifically.)
    fxAngle = fxAngle + 45#
    If (fxAngle > 360#) Then fxAngle = fxAngle - 360#
    fxAngle = -fxAngle
    
    If (Not toPreview) Then Message "Summoning M. C. Escher..."
    
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
    
    'Calculate the center of the image
    Dim midX As Double, midY As Double
    midX = CDbl(finalX - initX) * centerX
    midX = midX + initX + 0.1   'Avoid 0 which crashes log(), below
    midY = CDbl(finalY - initY) * centerY
    midY = midY + initY + 0.1
    
    'Rotation values
    Dim theta As Double, sRadius As Double
    
    'X and Y values, remapped around a center point of (0, 0)
    Dim nX As Double, nY As Double
    
    'Max radius is calculated as the distance from the center of the image to a corner
    Dim tWidth As Long, tHeight As Long
    tWidth = curDIBValues.Width
    tHeight = curDIBValues.Height
    sRadius = PDMath.Max2Int(tWidth, tHeight)
    
    'Define inner and outer radii as a ratio of the image's largest dimension
    Dim inRadius As Double, outRadius As Double
    inRadius = sRadius * (fxInnerRadius / 100#)
    outRadius = sRadius * (fxOuterRadius / 100#)
    
    'Next comes a bunch of ugly logarithmic math!  Formulas are derived from this page (link good as of August 2021):
    ' http://www.josleys.com/article_show.php?id=82
    Dim rFrac As Double
    rFrac = Math.Log(outRadius / inRadius)
    
    'Dim theta As Double
    theta = PDMath.Atan2(rFrac, PI_DOUBLE)
    
    'double f = Math.Cos(alpha);
    Dim f As Double
    f = Math.Cos(theta)
    
    'And now for the really fun stuff: COMPLEX NUMBERS.  See the ComplexNumbers module for details on
    ' how these functions are implemented, as well as additional copyright and license information.
    Dim iTheta As ComplexNumberF, beta As ComplexNumberF
    iTheta = ComplexNumbers.make_complexf(0!, theta)
    beta = ComplexNumbers.c_expf(iTheta)
    beta.c_real = beta.c_real * f
    beta.c_imag = beta.c_imag * f
    
    Dim zin As ComplexNumberF, ztemp1 As ComplexNumberF, ztemp2 As ComplexNumberF, zout As ComplexNumberF
    Dim from_x As Double, from_y As Double
    Dim rtemp As Double
    
    Dim rotatedX As Double, rotatedY As Double
    Dim angleRad As Double
    angleRad = PDMath.DegreesToRadians(fxAngle)
            
    Dim angleCos As Double, angleSin As Double
    angleCos = Math.Cos(angleRad)
    angleSin = Math.Sin(angleRad)
    
    'PD-specific support
    Dim tmpQuad As RGBQuad, newQuad As RGBQuad
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
        j = x - midX
        k = midY - y    'Notice the deliberate remap vs y - midY; this makes the result vertically correct, by default
        
        'Sample a number of source pixels corresponding to the user's supplied quality value; more quality means
        ' more samples, and much better representation in the final output.
        For sampleIndex = 0 To numSamples
        
            'Offset the pixel amount by the supersampling lookup table
            nX = j + ssX(sampleIndex)
            nY = k + ssY(sampleIndex)
            
            'Reset our color and complex number values
            zout.c_real = nX
            zout.c_imag = nY
            tmpQuad.Blue = 0: tmpQuad.Green = 0: tmpQuad.Red = 0: tmpQuad.Alpha = 0
            
            'Initial calculation
            ztemp1 = ComplexNumbers.c_divf(ComplexNumbers.c_logf(zout), beta)
            
            'Repeat until we reach a pixel inside the image.  (On the 3rd try, if no pixel is reached,
            ' we will apply the filter's current edge-wrap mode to ensure a valid pixel is used.)
            Const LOOP_MAX As Long = 2
            Dim loopFind As Long
            For loopFind = 0 To LOOP_MAX
                
                If fxSwapOrder Then
                    rtemp = PDMath.Modulo(ztemp1.c_real, rFrac) + (LOOP_MAX - loopFind) * rFrac
                Else
                    rtemp = PDMath.Modulo(ztemp1.c_real, rFrac) + loopFind * rFrac
                End If
                
                ztemp2.c_real = rtemp
                ztemp2.c_imag = ztemp1.c_imag
                
                'Inverse to bring coordinates back into real space
                zin = ComplexNumbers.c_expf(ztemp2)
                zin.c_real = zin.c_real * inRadius
                zin.c_imag = zin.c_imag * inRadius
                
                'Re-center around original origin
                from_x = zin.c_real + midX
                from_y = midY - zin.c_imag
                
                'Finally, rotate by the user-supplied angle
                rotatedX = ((from_x - midX) * angleCos - (from_y - midY) * angleSin + midX)
                rotatedY = ((from_y - midY) * angleCos - (from_x - midX) * -angleSin + midY)
                
                'On the first two passes, ignore PD's custom edge-handling capabilities and simply look
                ' for a pixel inside the image.  (Depending on parameters, this may fail with somewhat
                ' high probability.)
                If (loopFind < LOOP_MAX) Then
                    newQuad = fSupport.GetColorsFromSource_FastErase(rotatedX, rotatedY)
                    
                'If we haven't found an opaque pixel by the 3rd pass, use edge-wrapping to ensure a pixel
                ' is reached.  This avoids ugly "holes" in the output.
                Else
                    newQuad = fSupport.GetColorsFromSource(rotatedX, rotatedY, x, y)
                End If
                
                'Blend the retrieved color with our existing color tracker, using non-standard blend rules
                ' (see AddColor() for details).
                If (tmpQuad.Alpha = 0) Then
                    tmpQuad = newQuad
                Else
                    tmpQuad = AddColor(tmpQuad, newQuad)
                End If
                
                'Once an opaque pixel is reached, exit the loop immediately
                If (tmpQuad.Alpha = 255) Then Exit For
                
            Next loopFind
            
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

'Cheap additive color function for improved antialiasing along edge boundaries.  Note that this does *NOT*
' use normal blend formulas, by design.  (If you're curious, as I was, about how normal color-blending would
' work, you can easily plug it in - but the results produce worse output, especially along edges, due to the
' way the algorithm renders hard edges across much of the image.)
Private Function AddColor(ByRef origColor As RGBQuad, ByRef newColor As RGBQuad) As RGBQuad
    
    Dim addAlpha As Long, totalAlpha As Long
    addAlpha = newColor.Alpha
    totalAlpha = 255 - origColor.Alpha
    If (totalAlpha < addAlpha) Then addAlpha = totalAlpha
    totalAlpha = origColor.Alpha + addAlpha
    AddColor.Alpha = totalAlpha
    
    Dim orig_frac As Double
    orig_frac = origColor.Alpha / totalAlpha
    
    Dim add_frac As Double
    add_frac = addAlpha / totalAlpha
    
    AddColor.Blue = Int(origColor.Blue * orig_frac + newColor.Blue * add_frac)
    AddColor.Green = Int(origColor.Green * orig_frac + newColor.Green * add_frac)
    AddColor.Red = Int(origColor.Red * orig_frac + newColor.Red * add_frac)
    
End Function

Private Sub chkOrder_Click()
    UpdatePreview
End Sub

'OK button
Private Sub cmdBar_OKClick()
    Process "Droste", , GetLocalParamString(), UNDO_Layer
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub cmdBar_ResetClick()
    cboEdges.ListIndex = pdeo_Clamp
    sltQuality.Value = 2
End Sub

Private Sub Form_Load()
    
    'Suspend previews until the dialog has fully loaded
    cmdBar.SetPreviewStatus False
    
    'I use a central function to populate the edge handling combo box; this way, I can add new methods and have
    ' them immediately available to all distort functions.
    PopDistortEdgeBox cboEdges, pdeo_Clamp
    
    'Apply translations and visual themes
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

Private Sub sldRadiusInner_Change()
    UpdatePreview
End Sub

Private Sub sldRadiusOuter_Change()
    UpdatePreview
End Sub

Private Sub sltAngle_Change()
    UpdatePreview
End Sub

Private Sub sltQuality_Change()
    UpdatePreview
End Sub

'Redraw the on-screen preview of the transformed image
Private Sub UpdatePreview()
    If cmdBar.PreviewsAllowed Then FxDroste GetLocalParamString(), True, pdFxPreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub pdFxPreview_ViewportChanged()
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
        .AddParam "angle", sltAngle.Value
        .AddParam "radius-inner", sldRadiusInner.Value
        .AddParam "radius-outer", sldRadiusOuter.Value
        .AddParam "edges", cboEdges.ListIndex
        .AddParam "quality", sltQuality.Value
        .AddParam "swap-order", chkOrder.Value
        .AddParam "centerx", sltXCenter.Value
        .AddParam "centery", sltYCenter.Value
    End With
    
    GetLocalParamString = cParams.GetParamString()
    
End Function
