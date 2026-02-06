VERSION 5.00
Begin VB.Form FormSunshine 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   " Sunshine"
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
   Begin PhotoDemon.pdCheckBox chkRandomize 
      Height          =   375
      Left            =   6075
      TabIndex        =   11
      Top             =   5160
      Width           =   5805
      _ExtentX        =   10239
      _ExtentY        =   661
      Caption         =   "randomize"
      FontSize        =   11
   End
   Begin PhotoDemon.pdSlider sldRadius 
      Height          =   705
      Left            =   6000
      TabIndex        =   1
      Top             =   1560
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   1244
      Caption         =   "radius"
      Min             =   1
      Max             =   100
      SigDigits       =   1
      Value           =   50
      NotchPosition   =   2
      NotchValueCustom=   50
   End
   Begin PhotoDemon.pdFxPreviewCtl pdFxPreview 
      Height          =   5625
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9922
      DisableZoomPan  =   -1  'True
      PointSelection  =   -1  'True
   End
   Begin PhotoDemon.pdSlider sldRays 
      Height          =   705
      Left            =   9000
      TabIndex        =   2
      Top             =   1560
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   1244
      Caption         =   "number of rays"
      Min             =   1
      Max             =   200
      Value           =   50
      NotchPosition   =   2
      NotchValueCustom=   50
   End
   Begin PhotoDemon.pdSlider sldXCenter 
      Height          =   405
      Left            =   6000
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
   Begin PhotoDemon.pdSlider sldYCenter 
      Height          =   405
      Left            =   9000
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
   Begin PhotoDemon.pdCommandBar cmdBar 
      Height          =   750
      Left            =   0
      TabIndex        =   7
      Top             =   5790
      Width           =   12090
      _ExtentX        =   21325
      _ExtentY        =   1323
   End
   Begin PhotoDemon.pdColorSelector clrBase 
      Height          =   810
      Left            =   6000
      TabIndex        =   5
      Top             =   4080
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   1429
      Caption         =   "color"
      curColor        =   8978431
   End
   Begin PhotoDemon.pdSlider sldColorVariance 
      Height          =   705
      Left            =   9000
      TabIndex        =   6
      Top             =   2400
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   1244
      Caption         =   "color variance"
      Max             =   100
      SigDigits       =   1
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
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   285
      Left            =   6000
      Top             =   120
      Width           =   5925
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "center position (x, y)"
      FontSize        =   12
      ForeColor       =   4210752
   End
   Begin PhotoDemon.pdSlider sldLengthVariance 
      Height          =   705
      Left            =   6000
      TabIndex        =   8
      Top             =   2400
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   1244
      Caption         =   "ray variance"
      Max             =   100
      SigDigits       =   1
   End
   Begin PhotoDemon.pdDropDown cboBlendMode 
      Height          =   735
      Left            =   9000
      TabIndex        =   9
      Top             =   4080
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   1296
      Caption         =   "blend mode"
   End
   Begin PhotoDemon.pdSlider sldStrength 
      Height          =   705
      Left            =   6000
      TabIndex        =   10
      Top             =   3240
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   1244
      Caption         =   "strength"
      Min             =   0.1
      Max             =   100
      SigDigits       =   1
      Value           =   100
      NotchPosition   =   2
      NotchValueCustom=   100
   End
   Begin PhotoDemon.pdSlider sldHDR 
      Height          =   705
      Left            =   9000
      TabIndex        =   12
      Top             =   3240
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   1244
      Caption         =   "HDR"
      Max             =   500
      SigDigits       =   1
      Value           =   100
      NotchPosition   =   2
      NotchValueCustom=   100
   End
End
Attribute VB_Name = "FormSunshine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Sunshine Effect Form
'Copyright 2017-2026 by Tanner Helland
'Created: 31/July/17
'Created: 01/August/17
'Last update: finish work on new implementation
'
'Overlay a "light-burst" effect on a given image.  The overlay is generated as a standalone 32-bpp layer,
' and once complete, it is auto-blended onto the base layer.  If you want the effect as a standalone image,
' simply apply it to a blank 32-bpp layer.
'
'This tool uses a heavily modified version of a "sparkle" algorithm originally developed by
' Jerry Huxtable of JH Labs.  Jerry's original code is licensed under an Apache 2.0 license.
' You can download his original version at the following link (good as of 31 July '17):
' http://www.jhlabs.com/ip/filters/index.html
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'To improve performance, the sunshine effect is rendered onto a blank intermediary DIB, and once it is ready,
' the final result is composited onto the image.  To improve performance during previews, we cache the overlay.
Private m_RayOverlay As pdDIB

'Persistent random seeds are supported, which allows previews and final results to appear the same
Private m_Randomize As pdRandomize

Public Sub fxSunshine(ByVal effectParams As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)
    
    If (Not toPreview) Then Message "Generating light beams..."
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString effectParams
    
    Dim sunRadius As Double, numBeams As Long, baseColor As Long, colorVariance As Long
    Dim centerX As Double, centerY As Double, lengthVariance As Double, beamStrength As Double, hdrStrength As Double
    Dim overlayBlend As PD_BlendMode
    
    With cParams
        sunRadius = .GetDouble("radius", 100#)
        numBeams = .GetLong("rays", sldRays.Value)
        beamStrength = .GetDouble("strength", 100#) * 0.01
        lengthVariance = .GetDouble("lengthvariance", 0#) * 0.01
        hdrStrength = .GetDouble("hdr", 100#) * 0.001
        baseColor = .GetLong("color", clrBase.Color)
        m_Randomize.SetSeed_Float .GetDouble("seed", m_Randomize.GetRandomFloat_VB)
        colorVariance = .GetLong("colorvariance", sldColorVariance.Value)
        overlayBlend = .GetLong("blendmode", BM_Normal)
        centerX = .GetDouble("centerx", 0.5)
        centerY = .GetDouble("centery", 0.5)
    End With
    
    'Create a local array and point it at the pixel data of the current image
    Dim dstSA As SafeArray2D
    EffectPrep.PrepImageData dstSA, toPreview, dstPic, , , True
    
    'If this is a preview, we need to adjust the radius to match the size of the preview box
    sunRadius = (sunRadius * 0.005)
    If (curDIBValues.Width < curDIBValues.Height) Then
        sunRadius = sunRadius * curDIBValues.Height * 0.5
    Else
        sunRadius = sunRadius * curDIBValues.Width * 0.5
    End If
    
    'If this is *not* a preview, we need to generate a progress bar
    Dim progBarCheck As Long
    If (Not toPreview) Then
        ProgressBars.SetProgBarMax curDIBValues.Bottom
        progBarCheck = ProgressBars.FindBestProgBarValue()
    End If
    
    'Prepare a blank overlay at the same size as the image; to improve performance, we'll render the
    ' sun beams onto this blank image, then use a pdCompositor instance to merge the results.
    If (m_RayOverlay Is Nothing) Then Set m_RayOverlay = New pdDIB
    If (m_RayOverlay.GetDIBWidth <> curDIBValues.Width) Or (m_RayOverlay.GetDIBHeight <> curDIBValues.Height) Then
        m_RayOverlay.CreateBlank curDIBValues.Width, curDIBValues.Height, 32, 0, 0
    Else
        m_RayOverlay.ResetDIB 0
    End If
    
    Dim pxOverlay() As Byte, pxSA As SafeArray1D
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left
    initY = curDIBValues.Top
    finalX = curDIBValues.Right
    finalY = curDIBValues.Bottom
    
    'To improve performance on the inner loop, all light beam lengths and colors (if color variation
    ' is active) are pre-calculated.
    Dim beamLengths() As Double, beamColorR() As Double, beamColorG() As Double, beamColorB() As Double
    ReDim beamLengths(0 To numBeams - 1) As Double
    ReDim beamColorR(0 To numBeams - 1) As Double
    ReDim beamColorG(0 To numBeams - 1) As Double
    ReDim beamColorB(0 To numBeams - 1) As Double
    
    'The user's color is used as the base for our burst, but note that colors are applied using a
    ' compositor and blend mode, so their final appearance in the image may vary in non-obvious ways.
    Dim newR As Double, newG As Double, newB As Double
    newR = Colors.ExtractRed(baseColor) / 255#
    newG = Colors.ExtractGreen(baseColor) / 255#
    newB = Colors.ExtractBlue(baseColor) / 255#
    
    'If hue variations are enabled, we'll use the HSV color space to calculate variant colors
    Dim colorShiftThreshold As Double
    colorShiftThreshold = colorVariance / 200#
    
    Dim h As Double, s As Double, newH As Double, v As Double
    Dim spokeRed As Double, spokeGreen As Double, spokeBlue As Double
    
    'Calculate HSV equivalents of the target color
    fRGBtoHSV newR, newG, newB, h, s, v
    
    Dim i As Long, j As Long
    For i = 0 To numBeams - 1
    
        'To get a pseudo-normal distribution of lengths, we take the mean of several random values.
        ' (for details, see https://stackoverflow.com/questions/2325472/generate-random-numbers-following-a-normal-distribution-in-c-c)
        Dim rndSum As Double, rndIterations As Long
        rndSum = 0#
        rndIterations = 3
        
        For j = 0 To rndIterations - 1
            rndSum = rndSum + m_Randomize.GetRandomFloat_WH()
        Next j
        rndSum = rndSum * (1# / CDbl(rndIterations))
        
        'Calculate the length of this spoke by treating the user's specified radius as a maximum,
        ' and shrinking it relative to the random length generated above,
        Dim tmpRadius As Double
        tmpRadius = sunRadius + (sunRadius * lengthVariance * rndSum * 1.5) - (sunRadius * lengthVariance * 0.75)
        If (tmpRadius < 1#) Then tmpRadius = 1#
        beamLengths(i) = tmpRadius
        
        'While we're here, randomize the hue for this spoke according to a user-specified threshold.
        ' (Note that we invoke the random number generator *even if we don't use its return* - this
        ' guarantees that the spoke length, above, doesn't change, even if the user toggles the
        ' color shift setting.)
        tmpRadius = m_Randomize.GetRandomFloat_WH()
        If (colorShiftThreshold <> 0#) Then
            
            newH = h + (tmpRadius * 2# - 1#) * colorShiftThreshold
            If (newH > 1#) Then newH = newH - 1#
            If (newH < 0#) Then newH = newH + 1#
            
            Dim rFloat As Double, gFloat As Double, bFloat As Double
            fHSVtoRGB newH, s, v, rFloat, gFloat, bFloat
        
            beamColorR(i) = rFloat
            beamColorG(i) = gFloat
            beamColorB(i) = bFloat
        
        Else
            beamColorR(i) = newR
            beamColorG(i) = newG
            beamColorB(i) = newB
        End If
        
    Next i
    
    'Calculate the center of the sunshine as an absolute position
    Dim midX As Double, midY As Double
    midX = CDbl(finalX - initX) * centerX
    midX = midX + initX
    midY = CDbl(finalY - initY) * centerY
    midY = midY + initY
    
    'Because this tool operates largely in polar mode (for calculating a circular sunburst),
    ' we need the usual assortment of cartesian > polar variables.
    Dim newX As Double, newY As Double
    Dim pxDistance As Double, pxAngle As Double, pxFade As Double, pxLength As Double
    Dim dstNormalized As Double, dstInt As Long
    Dim firstIndex As Long, secondIndex As Long
    Const PI2_INV As Double = 1# / (2# * PI)
    
    'Loop through each pixel in the image, converting values as we go
    initX = initX * 4
    finalX = finalX * 4
    
    For y = initY To finalY
        m_RayOverlay.WrapArrayAroundScanline pxOverlay, pxSA, y
    For x = initX To finalX Step 4
        
        'Calculate an angle and distance relative to the center of the light
        newX = (x * 0.25) - midX
        newY = y - midY
        pxAngle = PDMath.Atan2_Fastest(newY, newX)
        
        'Find the two rays neighboring this point.  We do this by normalizing the returned angle -
        ' remember that Atan2 returns a value on the range [-PI, PI], so we must convert it to
        ' [0, 1.0].  That normalized value is then multiplied by the number of spokes, which gives us
        ' an index into our pre-calculated beam table.
        dstNormalized = (pxAngle + PI) * numBeams * PI2_INV
        dstInt = Int(dstNormalized)
        pxFade = dstNormalized - dstInt
        
        firstIndex = dstInt Mod numBeams
        secondIndex = (dstInt + 1) Mod numBeams
        
        'Interpolate between the length of the two nearest neighboring rays
        pxLength = beamLengths(secondIndex) * pxFade + beamLengths(firstIndex) * (1# - pxFade)
        
        'Perform a similar interpolation for color
        spokeRed = beamColorR(secondIndex) * pxFade + beamColorR(firstIndex) * (1# - pxFade)
        spokeGreen = beamColorG(secondIndex) * pxFade + beamColorG(firstIndex) * (1# - pxFade)
        spokeBlue = beamColorB(secondIndex) * pxFade + beamColorB(firstIndex) * (1# - pxFade)
        
        'Using the calculated length and color values, and this point's distance from the center
        ' of the image, we now want to calculate an appropriate intensity value.
        
        'Start by normalizing this pixel's length from the sunbeam's center.
        pxDistance = newX * newX + newY * newY
        dstNormalized = pxLength * pxLength / (pxDistance + 0.000000001)
        
        'Apply the strength value provided by the user, if any
        'dstNormalized = dstNormalized * beamStrength
        
        'Next, take the fade value - which is a fraction [0, 1.0] describing where this point lies
        ' on the angle between beam(firstIndex) and beam(secondIndex) - and convert it to the range
        ' [-0.5, 0.5].  This lets us center each ray over its position, which simplifies calculations.
        pxFade = pxFade - 0.5
        
        'Because we're dealing with light, square the fade distance to create natural fall-off
        pxFade = 1# - pxFade * pxFade
        
        'Finally, reduce the intensity of the pixel by its distance from the center, including the
        ' user's specified beam strength (if any)
        pxFade = pxFade * dstNormalized * beamStrength
        
        'As a failsafe, clamp the output to [0, 1]
        If (pxFade < 0#) Then pxFade = 0#
        If (pxFade > 1#) Then pxFade = 1#
        
        'pxFade represents the alpha value of this pixel, and it is now calculated completely.
        
        'As a final step, light sources in an image always look better if they're given proper
        ' HDR treatment.  This means we want to increase light intensity above the color specified
        ' by the user, in regions where the light is most intense.
        
        'Start by calculating an HDR modifier for this pixel.  This value will be added to the
        ' original pixel value, so we want it to be affected by not just HDR strength, but by
        ' the pixel's distance from the center, including any strength modifiers supplied by
        ' the user.  (Also, because it's additive, we want it on the range [0, 255].)
        dstNormalized = dstNormalized * dstNormalized * beamStrength * hdrStrength * 255#
        
        'Apply the HDR modifier, if any, to this pixel's original color value to arrive at a
        ' "final" pixel value.
        bFloat = spokeBlue * 255# + (spokeBlue * dstNormalized)
        gFloat = spokeGreen * 255# + (spokeGreen * dstNormalized)
        rFloat = spokeRed * 255# + (spokeRed * dstNormalized)
        
        'Clamp any excessive values
        If (bFloat > 255#) Then bFloat = 255#
        If (gFloat > 255#) Then gFloat = 255#
        If (rFloat > 255#) Then rFloat = 255#
        
        'Treating the previously calculated intensity value as an alpha value, apply both colors
        ' and alpha to the overlay image.  (Because this is the final step, alpha is premultiplied;
        ' this allows for very rapid blending when the overlay image is complete.)
        pxOverlay(x) = bFloat * pxFade
        pxOverlay(x + 1) = gFloat * pxFade
        pxOverlay(x + 2) = rFloat * pxFade
        pxOverlay(x + 3) = pxFade * 255#
        
    Next x
        If (Not toPreview) Then
            If (y And progBarCheck) = 0 Then
                If Interface.UserPressedESC() Then Exit For
                SetProgBarVal y
            End If
        End If
    Next y
    
    'Unwrap our scanline tracker
    m_RayOverlay.UnwrapArrayFromDIB pxOverlay
    m_RayOverlay.SetInitialAlphaPremultiplicationState True
    
    'Prepare a pdCompositor instance; it will perform the actual blend for us
    Dim cCompositor As pdCompositor
    Set cCompositor = New pdCompositor
    cCompositor.QuickMergeTwoDibsOfEqualSize workingDIB, m_RayOverlay, overlayBlend
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    EffectPrep.FinalizeImageData toPreview, dstPic, True
    
End Sub

Private Sub cboBlendMode_Click()
    UpdatePreview
End Sub

Private Sub chkRandomize_Click()
    UpdatePreview
End Sub

'OK button
Private Sub cmdBar_OKClick()
    Process "Sunshine", , GetLocalParamString(False), UNDO_Layer
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub cmdBar_ResetClick()
    clrBase.Color = RGB(255, 255, 60)
End Sub

Private Sub clrBase_ColorChanged()
    UpdatePreview
End Sub

Private Sub Form_Load()

    'Disable previewing until the form has been fully initialized
    cmdBar.SetPreviewStatus False
    
    Set m_Randomize = New pdRandomize
    m_Randomize.SetSeed_AutomaticAndRandom
    
    'Populate the blend mode drop-down
    Interface.PopulateBlendModeDropDown cboBlendMode, BM_Normal
    
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
    sldXCenter.Value = xRatio
    sldYCenter.Value = yRatio
    cmdBar.SetPreviewStatus True
    UpdatePreview
End Sub

Private Sub pdFxPreview_ViewportChanged()
    UpdatePreview
End Sub

Private Sub sldColorVariance_Change()
    UpdatePreview
End Sub

Private Sub sldHDR_Change()
    UpdatePreview
End Sub

Private Sub sldLengthVariance_Change()
    UpdatePreview
End Sub

Private Sub sldRadius_Change()
    UpdatePreview
End Sub

Private Sub sldRays_Change()
    UpdatePreview
End Sub

Private Sub sldStrength_Change()
    UpdatePreview
End Sub

Private Sub sldXCenter_Change()
    UpdatePreview
End Sub

Private Sub sldYCenter_Change()
    UpdatePreview
End Sub

Private Sub UpdatePreview()
    If cmdBar.PreviewsAllowed Then Me.fxSunshine GetLocalParamString(), True, pdFxPreview
End Sub

Private Function GetLocalParamString(Optional ByVal isPreview As Boolean = True) As String
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    
    With cParams
        
        .AddParam "radius", sldRadius.Value
        .AddParam "rays", sldRays.Value
        .AddParam "lengthvariance", sldLengthVariance.Value
        .AddParam "strength", sldStrength.Value
        
        'Randomizing is a bit weird; we only do it if the user has enabled it, *and* if it's a preview.
        ' (This allows the actual effect to match the last preview the user saw.)
        If (chkRandomize.Value And isPreview) Then m_Randomize.SetSeed_AutomaticAndRandom
        .AddParam "seed", m_Randomize.GetSeed
        
        .AddParam "color", clrBase.Color
        .AddParam "colorvariance", sldColorVariance.Value
        .AddParam "blendmode", cboBlendMode.ListIndex
        .AddParam "hdr", sldHDR.Value
        
        .AddParam "centerx", sldXCenter.Value
        .AddParam "centery", sldYCenter.Value
        
    End With
    
    GetLocalParamString = cParams.GetParamString()
    
End Function
