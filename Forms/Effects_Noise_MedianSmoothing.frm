VERSION 5.00
Begin VB.Form FormMedian 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   " Median"
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
   Begin PhotoDemon.pdCheckBox chkLuminance 
      Height          =   375
      Left            =   6000
      TabIndex        =   5
      Top             =   2280
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   661
      Caption         =   "luminance only"
   End
   Begin PhotoDemon.pdCommandBar cmdBar 
      Align           =   2  'Align Bottom
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
      Top             =   1440
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   1244
      Caption         =   "radius"
      Min             =   1
      Max             =   500
      ScaleStyle      =   1
      Value           =   5
      DefaultValue    =   1
   End
   Begin PhotoDemon.pdSlider sltPercent 
      Height          =   705
      Left            =   6000
      TabIndex        =   3
      Top             =   3960
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   1244
      Caption         =   "percentile"
      Min             =   1
      Max             =   100
      Value           =   50
      NotchPosition   =   2
      NotchValueCustom=   50
   End
   Begin PhotoDemon.pdButtonStrip btsKernelShape 
      Height          =   1095
      Left            =   6000
      TabIndex        =   4
      Top             =   2760
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   1931
      Caption         =   "kernel shape"
   End
End
Attribute VB_Name = "FormMedian"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Median Filter Tool
'Copyright 2013-2026 by Tanner Helland
'Created: 08/Feb/13
'Last updated: 23/June/20
'Last update: as median radius increases, use an increasingly strong wavelet approximation.
'              This provides large (100x+) improvements to performance at large (100px+) radii.
'
'This is a heavily optimized median filter function.  An "accumulation" technique is used instead
' of the standard sliding window mechanism.
' (See http://web.archive.org/web/20060718054020/http://www.acm.uiuc.edu/siggraph/workshops/wjarosz_convolution_2001.pdf)
' This allows the algorithm to perform extremely well, despite being written in pure VB.
'
'That said, it is still unfortunately slow in the IDE.  I STRONGLY recommend compiling before
' applying any median filter(s).
'
'An optional percentile option is available.  At minimum value, this performs identically to an
' erode (minimum) filter.  Similarly, at max value it performs identically to a dilate (maximum)
' filter.  Default setting is 50%.
'
'To improve performance at very large radii, a wavelet approximation is automatically enabled as
' radius increases.  On e.g. an 8-megapixel photo, this reduces running time of a 100px median filter
' from ~160 seconds to ~7 seconds, and produces an image that is 99.86% identical to a "true" median
' (measured using PD's built-in RMSD tool).
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Because this tool can be used for multiple actions (median, dilate, erode), we need to track which mode is currently active.
' When the reset or randomize buttons are pressed, we will automatically adjust our behavior to match.
Private Enum MedianToolMode
    MEDIAN_DEFAULT = 0
    MEDIAN_DILATE = 1
    MEDIAN_ERODE = 2
End Enum
Private curMode As MedianToolMode

'Apply a median filter to the image (heavily optimized accumulation implementation!)
'Input: radius of the median (min 1, no real max - but the scroll bar is maxed at 200 presently)
Public Sub ApplyMedianFilter(ByVal parameterList As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)
    
    'Parse out the parameter list
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString parameterList
    
    Dim mRadius As Long, mPercent As Double, kernelShape As PD_PixelRegionShape, useLab As Boolean
    mRadius = cParams.GetLong("radius", 1&)
    mPercent = cParams.GetLong("percent", 50&)
    kernelShape = cParams.GetLong("kernelshape", PDPRS_Rectangle)
    useLab = cParams.GetBool("luminance-only", False, True)
    
    If (Not toPreview) Then
        If (mPercent = 1) Then
            Message "Applying erode (minimum rank) filter..."
        ElseIf (mPercent = 100) Then
            Message "Applying dilate (maximum rank) filter..."
        Else
            Message "Applying median filter..."
        End If
    End If
    
    'Create a local array and point it at the pixel data of the current image
    Dim dstSA As SafeArray2D
    EffectPrep.PrepImageData dstSA, toPreview, dstPic
    
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
        Filters_ByteArray.Median_ByteArray mRadius, mPercent, kernelShape, labLOnly, dstLabLOnly, srcDIB.GetDIBWidth, srcDIB.GetDIBHeight, toPreview
        
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
        Filters_Layers.CreateMedianDIB mRadius, mPercent, kernelShape, srcDIB, dstDIB, toPreview
        
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
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering using the
    ' data inside workingDIB
    EffectPrep.FinalizeImageData toPreview, dstPic

End Sub

Private Sub btsKernelShape_Click(ByVal buttonIndex As Long)
    UpdatePreview
End Sub

Private Sub chkLuminance_Click()
    UpdatePreview
End Sub

Private Sub cmdBar_OKClick()
    Process "Median", , GetLocalParamString(), UNDO_Layer
End Sub

'Because this dialog can be used for multiple tools, we need to clarify some behavior when resetting and randomizing
Private Sub cmdBar_RandomizeClick()

    Select Case curMode
    
        Case MEDIAN_DEFAULT
            
        Case MEDIAN_DILATE
            sltPercent.Value = 100
        
        Case MEDIAN_ERODE
            sltPercent.Value = 1
    
    End Select

End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub cmdBar_ResetClick()

    Select Case curMode
    
        Case MEDIAN_DEFAULT
            sltPercent.Value = 50
            
        Case MEDIAN_DILATE
            sltPercent.Value = 100
        
        Case MEDIAN_ERODE
            sltPercent.Value = 1
    
    End Select
    
End Sub

Private Sub Form_Load()
    
    'Disable previews while we get everything initialized
    cmdBar.SetPreviewStatus False
    
    'Populate the kernel shape box with whatever shapes PD currently supports
    Interface.PopKernelShapeButtonStrip btsKernelShape, PDPRS_Circle
    
    'Apply translations and visual themes
    ApplyThemeAndTranslations Me, True, True
    cmdBar.SetPreviewStatus True
    UpdatePreview
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'The median dialog is reused for several tools: minimum, median, maximum.
Public Sub SetMedianCutoff(ByVal initPercentage As Long)

    If (initPercentage = 1) Then
        If (Not g_WindowManager Is Nothing) Then g_WindowManager.SetWindowCaptionW Me.hWnd, " " & g_Language.TranslateMessage("Erode")
        sltPercent.Value = 1
        sltPercent.Visible = False
        cmdBar.SetToolName "Erode"
        curMode = MEDIAN_ERODE
        
    ElseIf (initPercentage = 100) Then
        If (Not g_WindowManager Is Nothing) Then g_WindowManager.SetWindowCaptionW Me.hWnd, " " & g_Language.TranslateMessage("Dilate")
        sltPercent.Value = 100
        sltPercent.Visible = False
        cmdBar.SetToolName "Dilate"
        curMode = MEDIAN_DILATE
        
    Else
        If (Not g_WindowManager Is Nothing) Then g_WindowManager.SetWindowCaptionW Me.hWnd, " " & g_Language.TranslateMessage("Median")
        sltPercent.Value = initPercentage
        sltPercent.Visible = True
        curMode = MEDIAN_DEFAULT
        
    End If
    
End Sub

Private Sub sltPercent_Change()
    UpdatePreview
End Sub

Private Sub sltRadius_Change()
    UpdatePreview
End Sub

Private Sub UpdatePreview()
    If cmdBar.PreviewsAllowed Then ApplyMedianFilter GetLocalParamString(), True, pdFxPreview
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
        .AddParam "percent", sltPercent.Value, True
        .AddParam "kernelshape", btsKernelShape.ListIndex, True
        .AddParam "luminance-only", chkLuminance.Value, True
    End With
    
    GetLocalParamString = cParams.GetParamString()
    
End Function
