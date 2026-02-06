VERSION 5.00
Begin VB.Form FormDustAndScratches 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   " Dust and scratches"
   ClientHeight    =   6540
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12030
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
   ScaleWidth      =   802
   Begin PhotoDemon.pdCheckBox chkLuminance 
      Height          =   375
      Left            =   6120
      TabIndex        =   3
      Top             =   3720
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   661
      Caption         =   "luminance only"
   End
   Begin PhotoDemon.pdCommandBar cmdBar 
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5790
      Width           =   12030
      _ExtentX        =   21220
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
      Top             =   1800
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
      Caption         =   "radius"
      Min             =   1
      Max             =   500
      ScaleStyle      =   1
      Value           =   5
      DefaultValue    =   1
   End
   Begin PhotoDemon.pdSlider sldThreshold 
      Height          =   705
      Left            =   6000
      TabIndex        =   4
      Top             =   2760
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
      Caption         =   "threshold"
      Max             =   100
      Value           =   5
      NotchPosition   =   2
      NotchValueCustom=   50
   End
End
Attribute VB_Name = "FormDustAndScratches"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Dust and scratches tool
'Copyright 2020-2026 by Tanner Helland
'Created: 17/September/20
'Last updated: 17/September/20
'Last update: spin off from normal Median tool; minor modifications to a median blur achieves
'             identical results to Photoshop's "Dust and Scratches" tool
'
'Thanks to the excellent detective work of "laviewpbt" (link good as of Sep 2020):
'
' https://www.cnblogs.com/Imageshop/p/11087804.html
'
'...we finally have "official" confirmation that Photoshop's Dust and Scratches tool is just a
' median filter with a threshold parameter slapped on top.  This confirmation makes it trivial
' for me to add an identical tool to PhotoDemon.
'
'For details on PD's median filter implementation, please refer to the median effect dialog;
' this tool is just a minor modification of work already done there.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Apply a dust-and-scratches filter to the image
'Input parameters:
' - radius of the median [1, any arbitrary max]
' - threshold percentage [0, 100] - it will be scaled to [0, 255] for 32-bit RGBA pixels
' - luminance-only [boolean] - only comparing luminance is faster and results are nearly identical at small radii
Public Sub ApplyDustAndScratchesFilter(ByVal parameterList As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)
    
    'Parse out the parameter list
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString parameterList
    
    Dim mRadius As Long, mThreshold As Long, useLab As Boolean
    mRadius = cParams.GetLong("radius", 1&)
    mThreshold = Int(CDbl(cParams.GetLong("threshold", 50&)) * 2.55 + 0.5)
    useLab = cParams.GetBool("luminance-only", False, True)
    
    If (Not toPreview) Then Message "Fixing dust and scratches..."
    
    'Create a local array and point it at the pixel data of the current image
    Dim dstSA As SafeArray2D
    EffectPrep.PrepImageData dstSA, toPreview, dstPic
    
    'Before proceeding any further, make a copy of the original, unmodified image.
    ' We need this because we're going to compare a median-filtered copy of the image
    ' against it to produce the final "dust-and-scratches-fixed" result.
    Dim origCopy As pdDIB
    Set origCopy = New pdDIB
    origCopy.CreateFromExistingDIB workingDIB
    
    'If this is a preview, we need to adjust the kernel radius to match the size of the preview box
    If toPreview Then
        mRadius = mRadius * curDIBValues.previewModifier
        If (mRadius < 1) Then mRadius = 1
    End If
    
    'Median filters are incredibly slow.  To improve performance at very large radii, we use a
    ' wavelet approach.  Figure out how far we can shrink the image while still maintaining an
    ' acceptable level of quality.
    Dim waveletsUsed As Boolean, waveletRatio As Double, waveletWidth As Long, waveletHeight As Long
    waveletsUsed = False
    
    If (mRadius > 1) Then
        
        'First, if the image is sufficiently small, don't optimize with wavelets as it's a
        ' waste of time.
        If (workingDIB.GetDIBWidth > mRadius) Or (workingDIB.GetDIBHeight > mRadius) Then
        
            'If the image is larger than the underlying radius, use a wavelet approximation.
            
            'Wavelets (https://en.wikipedia.org/wiki/Wavelet) are an old-school signal-processing
            ' approach.  Basically, you can represent any wave by breaking it down into a series of
            ' waves with different wavelengths and amplitudes, which - when added together -
            ' precisely reproduce the original wave.  This is very similar in concept to Fourier
            ' transforms (which are specifically sine waves) and DCTs (cosine waves).
            
            'In image processing, wavelets are typically copies of the image at different sizes,
            ' with each size representing a different level of detail.  You can modify different
            ' "wavelets" to isolate higher- or lower-frequency noise.
            
            'Per it's name, the median filter affects all frequencies with increasing strength
            ' as the radius of the effect increases.  This means that there's little point in
            ' wasting energy on high-frequency noise at large radii, as the whole point of a
            ' median filter is to *ignore* high-frequency noise.  So as median filter radius
            ' increases, we can operate on a smaller copy of the image with minimal data loss,
            ' provided we...
            ' 1) scale the wavelet size at an appropriate rate for median radius changes, and...
            ' 2) use appropriate sampling techniques to ensure the wavelet approach doesn't
            '    inadvertently introduce new noise.
            
            '(2) is easily covered by using smart resampling settings already built in to PD.
            ' (1) was tackled by good old-fashioned trial-and-error.  I sampled a half-dozen
            ' images at various sizes and median radii, then tested how aggressively I could
            ' downsample them while still retaining roughly 99.9% accuracy.  I fit the resulting
            ' settings to a curve and tweaked it against expected image sizes (e.g. PD can only
            ' handle images so large, and median radius is limited to 500px for practical
            ' reasons).  This is the end result.
            waveletRatio = 100# - (170# * CDbl(mRadius)) ^ 0.4
            
            'Convert the ratio to the range [0, 1], and invert it so that it represents
            ' a multiplication factor.
            If (waveletRatio > 100#) Then waveletRatio = 100#
            If (waveletRatio < 1#) Then waveletRatio = 1#
            waveletRatio = waveletRatio / 100#
            
            'If we produced a valid value, activate wavelet mode!
            waveletsUsed = (waveletRatio < 1#)
            
        End If
    
    End If
    
    Dim srcDIB As pdDIB, dstDIB As pdDIB
    Set srcDIB = New pdDIB
    
    'If wavelets are a valid option, create a copy of the source image at a smaller size and
    ' modify the median radius accordingly.
    If waveletsUsed Then
    
        waveletWidth = Int(CDbl(workingDIB.GetDIBWidth) * waveletRatio + 0.5)
        If (waveletWidth < 1) Then waveletWidth = 1
        
        waveletHeight = Int(CDbl(workingDIB.GetDIBHeight) * waveletRatio + 0.5)
        If (waveletHeight < 1) Then waveletHeight = 1
        
        mRadius = Int(CDbl(mRadius) * waveletRatio + 0.5)
        If (mRadius < 1) Then mRadius = 1
        
        srcDIB.CreateBlank waveletWidth, waveletHeight, 32, 0, 0
        GDI_Plus.GDIPlus_StretchBlt srcDIB, 0, 0, waveletWidth, waveletHeight, workingDIB, 0, 0, workingDIB.GetDIBWidth, workingDIB.GetDIBHeight, interpolationType:=GP_IM_HighQualityBicubic, isZoomedIn:=True, dstCopyIsOkay:=True
        
        Set dstDIB = New pdDIB
        dstDIB.CreateBlank waveletWidth, waveletHeight, 32, 0, 0
        
    'If the radius isn't small enough to warrant a wavelet approach, simply mirror the
    ' existing image as-is.  Note also that we can also cheat and "paint" the finished result
    ' directly into the working DIB provided us by PD's effect engine.
    Else
        srcDIB.CreateFromExistingDIB workingDIB
        Set dstDIB = workingDIB
    End If
    
    'Luminance-only mode uses a CIELab transform via LCMS
    If useLab Then
        
        'Create an RGBA to ALAB transform
        Dim cRGB As pdLCMSProfile
        Set cRGB = New pdLCMSProfile
        cRGB.CreateSRGBProfile True
        
        Dim cLAB As pdLCMSProfile
        Set cLAB = New pdLCMSProfile
        cLAB.CreateLabProfile True
        
        Dim cTransform As pdLCMSTransform
        Set cTransform = New pdLCMSTransform
        cTransform.CreateTwoProfileTransform cRGB, cLAB, TYPE_BGRA_8, TYPE_ALab_8, INTENT_PERCEPTUAL
        
        'Create an intermediary copy of the image and store the ALAB-transformed copy there
        Dim labCopy() As Byte
        ReDim labCopy(0 To srcDIB.GetDIBStride * srcDIB.GetDIBHeight - 1) As Byte
        cTransform.ApplyTransformToScanline srcDIB.GetDIBPointer, VarPtr(labCopy(0)), srcDIB.GetDIBWidth * srcDIB.GetDIBHeight
        
        'Next, copy the L channel only into a dedicated array; this is critical for improving
        ' CPU cache performance
        Dim labLOnly() As Byte, dstLabLOnly() As Byte
        ReDim labLOnly(0 To srcDIB.GetDIBWidth - 1, 0 To srcDIB.GetDIBHeight - 1) As Byte
        
        Dim x As Long, y As Long, imgOffset As Long
        
        For y = 0 To srcDIB.GetDIBHeight - 1
            imgOffset = y * srcDIB.GetDIBStride
        For x = 0 To srcDIB.GetDIBWidth - 1
            labLOnly(x, y) = labCopy(imgOffset + x * 4 + 3)
        Next x
        Next y
        
        'Apply the median filter
        ReDim dstLabLOnly(0 To srcDIB.GetDIBWidth - 1, 0 To srcDIB.GetDIBHeight - 1) As Byte
        Filters_ByteArray.Median_ByteArray mRadius, 50#, PDPRS_Rectangle, labLOnly, dstLabLOnly, srcDIB.GetDIBWidth, srcDIB.GetDIBHeight, toPreview
        
        'Copy the new L values back into the ALAB-transformed copy
        For y = 0 To srcDIB.GetDIBHeight - 1
            imgOffset = y * srcDIB.GetDIBStride
        For x = 0 To srcDIB.GetDIBWidth - 1
            labCopy(imgOffset + x * 4 + 3) = dstLabLOnly(x, y)
        Next x
        Next y
        
        'Translate the new ALAB copy back into the original source DIB
        cTransform.CreateTwoProfileTransform cLAB, cRGB, TYPE_ALab_8, TYPE_BGRA_8, INTENT_PERCEPTUAL
        cTransform.ApplyTransformToScanline VarPtr(labCopy(0)), dstDIB.GetDIBPointer, srcDIB.GetDIBWidth * srcDIB.GetDIBHeight
        
    Else
        
        'The median function does *not* copy over alpha values; as such, we have to fill those
        ' in manually before running the filter.
        dstDIB.CreateFromExistingDIB srcDIB
        Filters_Layers.CreateMedianDIB mRadius, 50#, PDPRS_Rectangle, srcDIB, dstDIB, toPreview
        
    End If
    
    'Regardless of how we produced the median effect, we're finished with the source DIB;
    ' free it as it may be quite large.
    Set srcDIB = Nothing
    
    'If wavelets were used, we now need to sample the processed result back into workingDIB.
    ' Importantly, note that different settings from the original downsample are used - this is
    ' intentional and critical for avoiding fringing while accurately mimicking the "soft edge"
    ' look produced by the full-radius median filter we're attempting to approximate.
    If waveletsUsed Then
        workingDIB.ResetDIB 0
        GDI_Plus.GDIPlus_StretchBlt workingDIB, 0, 0, workingDIB.GetDIBWidth, workingDIB.GetDIBHeight, dstDIB, 0, 0, dstDIB.GetDIBWidth, dstDIB.GetDIBHeight, interpolationType:=GP_IM_HighQualityBilinear, dstCopyIsOkay:=True
        Set dstDIB = Nothing
    End If
    
    'workingDIB now contains a median-filtered copy of the original image.
    ' We now need to do a (fast) pixel-by-pixel comparison, and replace pixels in the original
    ' image with their median equivalent if the difference between the two exceeds the user's
    ' (arbitrary!) threshold.
    Dim origPixels() As Byte, origPx1D As SafeArray1D
    Dim medianPixels() As Byte, medianPx1D As SafeArray1D
    
    Dim maxLen As Long
    maxLen = (workingDIB.GetDIBStride * workingDIB.GetDIBHeight) - 1
    
    Dim r1 As Long, g1 As Long, b1 As Long, lum1 As Long
    Dim r2 As Long, g2 As Long, b2 As Long, lum2 As Long
    
    origCopy.WrapArrayAroundDIB_1D origPixels, origPx1D
    workingDIB.WrapArrayAroundDIB_1D medianPixels, medianPx1D
    For x = 0 To maxLen Step 4
        b1 = origPixels(x)
        g1 = origPixels(x + 1)
        r1 = origPixels(x + 2)
        lum1 = (218 * r1 + 732 * g1 + 74 * b1) \ 1024
        b2 = medianPixels(x)
        g2 = medianPixels(x + 1)
        r2 = medianPixels(x + 2)
        lum2 = (218 * r2 + 732 * g2 + 74 * b2) \ 1024
        If (Abs(lum2 - lum1) > mThreshold) Then
            origPixels(x) = b2
            origPixels(x + 1) = g2
            origPixels(x + 2) = r2
            origPixels(x + 3) = medianPixels(x + 3)
        End If
    Next x
    
    origCopy.UnwrapArrayFromDIB origPixels
    workingDIB.UnwrapArrayFromDIB medianPixels
    
    'Finally, swap references (because the final result is inside the original image copy,
    ' *not* workingDIB)
    Set workingDIB = origCopy
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering using the
    ' data inside workingDIB
    EffectPrep.FinalizeImageData toPreview, dstPic

End Sub

Private Sub chkLuminance_Click()
    UpdatePreview
End Sub

Private Sub cmdBar_OKClick()
    Process "Dust and scratches", , GetLocalParamString(), UNDO_Layer
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub cmdBar_ResetClick()
    sldThreshold.Value = 50
End Sub

Private Sub Form_Load()
    
    'Disable previews while we get everything initialized
    cmdBar.SetPreviewStatus False
    
    'Apply translations and visual themes
    ApplyThemeAndTranslations Me, True, True
    cmdBar.SetPreviewStatus True
    UpdatePreview
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub sldThreshold_Change()
    UpdatePreview
End Sub

Private Sub sltRadius_Change()
    UpdatePreview
End Sub

Private Sub UpdatePreview()
    If cmdBar.PreviewsAllowed Then ApplyDustAndScratchesFilter GetLocalParamString(), True, pdFxPreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub pdFxPreview_ViewportChanged()
    UpdatePreview
End Sub

Private Function GetLocalParamString() As String
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    
    With cParams
        .AddParam "radius", sltRadius.Value, True
        .AddParam "threshold", sldThreshold.Value, True
        .AddParam "luminance-only", chkLuminance.Value, True
    End With
    
    GetLocalParamString = cParams.GetParamString()
    
End Function
