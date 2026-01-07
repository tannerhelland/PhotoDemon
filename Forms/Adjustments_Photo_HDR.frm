VERSION 5.00
Begin VB.Form FormHDR 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   " HDR"
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
   Begin PhotoDemon.pdSlider sltRadius 
      Height          =   705
      Left            =   6000
      TabIndex        =   2
      Top             =   1920
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   1244
      Caption         =   "quality"
      Min             =   1
      Max             =   100
      Value           =   5
      DefaultValue    =   5
   End
   Begin PhotoDemon.pdSlider sltStrength 
      Height          =   705
      Left            =   6000
      TabIndex        =   3
      Top             =   3000
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   1244
      Caption         =   "strength"
      Min             =   1
      Max             =   100
      SigDigits       =   1
      Value           =   20
      DefaultValue    =   20
   End
End
Attribute VB_Name = "FormHDR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Imitation HDR Tool
'Copyright 2014-2026 by Tanner Helland
'Created: 09/July/14
'Last updated: 20/July/17
'Last update: migrate to XML params, minor performance improvements
'
'This is a heavily optimized imitation HDR function.  HDR normally works by having a photographer take multiple shots
' of a scene (3-5, typically), each at a unique exposure.  Software then merges those photos together, selecting pixels
' from each exposure and blending them to produce an evenly exposed photo across a wide luminance range.  This not only
' produces a neat visual effect, but also reproduces detail in ways that would be impossible in a single exposure.
'
'While a merge-to-HDR function that operates in the traditional manner would be nice to eventually include in PD, the
' trouble of asking a photographer to capture multiple back-to-back photos, each at a different exposure, without
' shaking the camera, is no small feat.  The inclusion of HDR as a built-in mode on many cameras and smartphones has
' also reduced the utility such a technique in a separate piece of software.
'
'So instead, what I've done here is put together a tool that mimics the results of HDR, using a contrast-adaptive local
' equalization function.  The details are complicated, but basically the function calculates a local average around
' each pixel, using a user-supplied radius (presented in PD as "quality").  The difference between the current pixel and
' that average is then amplified and applied to each channel; this allows regions of color to stay consistent,
' without the distortion inherent to global equalization.
'
'Anyway, assuming the original photograph was exposed reasonably well, this function should produce a good result.
' Poorly exposed original photographs cannot be saved by this technique, however, especially if a smartphone camera
' or other cheap sensor was used, as the inherent noise will screw up the filter's ability to properly solve the
' partial histogram problem.  C'est la vie.  Applying a median or noise-reduction filter in advance might help to
' improve the output.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'New test approach to HDR.  Unsharp masking can produce an HDR-like image, and it can do it a hell of a lot faster
' than the CLAHE-based method we've been using.  I'm going to have some testers experiment with the new method, to see
' if they prefer it.
Public Sub ApplyImitationHDR(ByVal effectParams As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)
        
    If (Not toPreview) Then Message "Generating HDR map for image..."
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString effectParams
    
    Dim fxQuality As Double, blendStrength As Double
    fxQuality = cParams.GetDouble("radius", 5#)
    blendStrength = cParams.GetDouble("strength", 20#)
    
    'Create a local array and point it at the pixel data of the current image
    Dim dstSA As SafeArray2D
    EffectPrep.PrepImageData dstSA, toPreview, dstPic
    
    'Create a second local array.  This will contain the a copy of the current image, and we will use it as our source reference
    ' (This is necessary to prevent blurred pixel values from spreading across the image as we go.)
    Dim srcDIB As pdDIB
    Set srcDIB = New pdDIB
    srcDIB.CreateFromExistingDIB workingDIB
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left
    initY = curDIBValues.Top
    finalX = curDIBValues.Right
    finalY = curDIBValues.Bottom
    
    'fxQuality represents an HDR radius.  We actually treat this as a percentage of the current image size, using the
    ' largest dimension.  Max quality is 20% of the image.
    Dim largestDimension As Long
    If (finalX - initX) > (finalY - initY) Then largestDimension = (finalX - initX) Else largestDimension = (finalY - initY)
    
    Dim hdrRadius As Long
    hdrRadius = ((fxQuality / 100#) * largestDimension) * 0.2
    
    'Strength is used as an analog for multiple parameters.  Here, we use it to calculate a saturation modifier,
    ' which is applied linearly to the final RGB values, as a way to further pop colors.
    Dim satBoost As Double
    satBoost = 1# + (blendStrength / 100#) * 0.3
    
    'Strength is presented to the user on a [1, 100] scale, but we actually boost this to a literal value of [1, 200]
    blendStrength = (blendStrength * 2#) / 100#
    
    'If this is a preview, we need to adjust the kernel radius to match the size of the preview box
    If (hdrRadius = 0) Then hdrRadius = 1
    
    'I almost always recommend quality over speed for PD tools, but in this case, the fast option is SO much faster,
    ' and the results so indistinguishable (3% different according to the Central Limit Theorem:
    ' https://www.khanacademy.org/math/probability/statistics-inferential/sampling_distribution/v/central-limit-theorem?playlist=Statistics
    ' ), that I use the faster method instead.
    Dim gaussBlurSuccess As Long
    gaussBlurSuccess = 0
    
    Dim progBarCalculation As Long
    progBarCalculation = finalY * 3 + finalX * 3
    If (Not toPreview) Then ProgressBars.SetProgBarMax progBarCalculation + finalY
    gaussBlurSuccess = CreateApproximateGaussianBlurDIB(hdrRadius, workingDIB, srcDIB, 3, toPreview, progBarCalculation + finalY)
    
    'Assuming the blur was created successfully, proceed with the masking portion of the filter.
    If (gaussBlurSuccess <> 0) Then
    
        'Now that we have a gaussian DIB created in workingDIB, we can point arrays toward it and the source DIB
        Dim dstImageData() As Byte, dstSA1D As SafeArray1D
        Dim srcImageData() As Byte, srcSA1D As SafeArray1D
        
        Dim xStride As Long
        
        'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
        ' based on the size of the area to be processed.
        Dim progBarCheck As Long
        progBarCheck = ProgressBars.FindBestProgBarValue()
        
        'ScaleFactor is used to apply the unsharp mask.  Maximum strength can be any value, but PhotoDemon locks it at 10.
        Dim scaleFactor As Double, invScaleFactor As Double
        scaleFactor = blendStrength + 1#
        invScaleFactor = 1# - scaleFactor
    
        Dim blendVal As Double
        
        'More color variables - in this case, sums for each color component
        Dim r As Long, g As Long, b As Long
        Dim r2 As Long, g2 As Long, b2 As Long
        Dim newR As Long, newG As Long, newB As Long
        Dim h As Double, s As Double, l As Double
        Dim tLumDelta As Long
        
        Const ONE_DIV_255 As Double = 1# / 255#
        
        'The final step of the smart blur function is to find edges, and replace them with the blurred data as necessary
        For y = initY To finalY
            srcDIB.WrapArrayAroundScanline srcImageData, srcSA1D, y
            workingDIB.WrapArrayAroundScanline dstImageData, dstSA1D, y
        For x = initX To finalX
        
            xStride = x * 4
            
            'Retrieve the original image's pixels
            b = dstImageData(xStride)
            g = dstImageData(xStride + 1)
            r = dstImageData(xStride + 2)
            
            'Now, retrieve the gaussian pixels
            b2 = srcImageData(xStride)
            g2 = srcImageData(xStride + 1)
            r2 = srcImageData(xStride + 2)
            
            tLumDelta = Abs(GetLuminance(r, g, b) - GetLuminance(r2, g2, b2))
            
            newR = (scaleFactor * r) + (invScaleFactor * r2)
            newG = (scaleFactor * g) + (invScaleFactor * g2)
            newB = (scaleFactor * b) + (invScaleFactor * b2)
            
            If (newR > 255) Then newR = 255
            If (newR < 0) Then newR = 0
            If (newG > 255) Then newG = 255
            If (newG < 0) Then newG = 0
            If (newB > 255) Then newB = 255
            If (newB < 0) Then newB = 0
            
            blendVal = tLumDelta * ONE_DIV_255
            
            'Standard blend formula
            newR = ((1# - blendVal) * newR) + (blendVal * r)
            newG = ((1# - blendVal) * newG) + (blendVal * g)
            newB = ((1# - blendVal) * newB) + (blendVal * b)
            
            'Finally, apply a saturation boost proportional to the final calculated strength
            Colors.ImpreciseRGBtoHSL newR, newG, newB, h, s, l
            s = s * satBoost
            If (s > 1#) Then s = 1#
            Colors.ImpreciseHSLtoRGB h, s, l, newR, newG, newB
            
            dstImageData(xStride) = newB
            dstImageData(xStride + 1) = newG
            dstImageData(xStride + 2) = newR
                                    
        Next x
            If (Not toPreview) Then
                If (y And progBarCheck) = 0 Then
                    If Interface.UserPressedESC() Then Exit For
                    SetProgBarVal progBarCalculation + y
                End If
            End If
        Next y
        
        workingDIB.UnwrapArrayFromDIB dstImageData
        srcDIB.UnwrapArrayFromDIB srcImageData
        Set srcDIB = Nothing
        
    End If
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    EffectPrep.FinalizeImageData toPreview, dstPic
        
End Sub

'OK button
Private Sub cmdBar_OKClick()
    Process "HDR", , GetLocalParamString(), UNDO_Layer
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

Private Sub sltRadius_Change()
    UpdatePreview
End Sub

Private Sub UpdatePreview()
    If cmdBar.PreviewsAllowed Then ApplyImitationHDR GetLocalParamString(), True, pdFxPreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub pdFxPreview_ViewportChanged()
    UpdatePreview
End Sub

Private Sub sltStrength_Change()
    UpdatePreview
End Sub

Private Function GetLocalParamString() As String
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    
    With cParams
        .AddParam "radius", sltRadius.Value
        .AddParam "strength", sltStrength.Value
    End With
    
    GetLocalParamString = cParams.GetParamString()
    
End Function
