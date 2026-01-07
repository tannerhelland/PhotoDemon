VERSION 5.00
Begin VB.Form FormContour 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   " Trace contour"
   ClientHeight    =   6540
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11655
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
   ScaleWidth      =   777
   Begin PhotoDemon.pdCommandBar cmdBar 
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5790
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   1323
   End
   Begin PhotoDemon.pdCheckBox chkBlackBackground 
      Height          =   375
      Left            =   6120
      TabIndex        =   2
      Top             =   3120
      Width           =   5340
      _ExtentX        =   9419
      _ExtentY        =   661
      Caption         =   "use black background"
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
   Begin PhotoDemon.pdCheckBox chkSmoothing 
      Height          =   375
      Left            =   6120
      TabIndex        =   3
      Top             =   3720
      Width           =   5340
      _ExtentX        =   9419
      _ExtentY        =   661
      Caption         =   "apply contour smoothing"
   End
   Begin PhotoDemon.pdSlider sltThickness 
      Height          =   705
      Left            =   6000
      TabIndex        =   4
      Top             =   2160
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   1244
      Caption         =   "thickness"
      Min             =   1
      Max             =   100
      Value           =   1
      DefaultValue    =   1
   End
End
Attribute VB_Name = "FormContour"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Trace Contour (Outline) Tool
'Copyright 2013-2026 by Tanner Helland
'Created: 15/Feb/13
'Last updated: 24/June/20
'Last update: introduce wavelet optimizations for the median step; this provides large performance improvements
'
'Contour tracing is performed by "stacking" a series of filters together:
' 1) Gaussian blur to smooth out fine details
' 2) Median to unify colors and round out edges
' 3) Edge detection
' 4) Auto white balance (as the original edge detection function is quite dark)
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

Private Sub chkSmoothing_Click()
    UpdatePreview
End Sub

'Convolve an image using a gaussian kernel (separable implementation!)
'Input: radius of the contour (min 1, no real max - but the scroll bar is maxed at 200 presently)
Public Sub TraceContour(ByVal effectParams As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)
    
    If (Not toPreview) Then Message "Tracing image contour..."
            
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString effectParams
    
    Dim cRadius As Long, useBlackBackground As Boolean, useSmoothing As Boolean
    
    With cParams
        cRadius = .GetLong("thickness", sltThickness.Value)
        useBlackBackground = .GetBool("blackbackground", True)
        useSmoothing = .GetBool("smoothing", True)
    End With
    
    'Create a local array and point it at the pixel data of the current image
    Dim dstSA As SafeArray2D
    EffectPrep.PrepImageData dstSA, toPreview, dstPic
    
    'Create a second local array.  This will contain the a copy of the current image, and we will
    ' use it as our source reference (necessary to prevent modified pixel values from spreading
    ' across the image as we go.)
    Dim srcDIB As pdDIB
    Set srcDIB = New pdDIB
    srcDIB.CreateFromExistingDIB workingDIB
    
    'If this is a preview, we need to adjust the kernel radius to match the size of the preview box
    If toPreview Then
        cRadius = cRadius * curDIBValues.previewModifier
        If (cRadius = 0) Then cRadius = 1
    End If
    
    Dim finalX As Long, finalY As Long
    finalX = workingDIB.GetDIBWidth
    finalY = workingDIB.GetDIBHeight
    
    Dim progBarMax As Long, progBarOffsets() As Long
        
    If useSmoothing Then
        
        'Calculate a progress bar maximum value.  At present, each sub-function has its own strategy for
        ' calculating progress.
        ReDim progBarOffsets(0 To 3) As Long
        
        'Blur uses 3 iterations in each direction (width/height)
        progBarOffsets(0) = 0
        progBarMax = finalX * 3 + finalY * 3
        progBarOffsets(1) = progBarMax
        
        'Median uses a wavelet approximation; its size is specific to the median filter.
        ' (For details on how this works, refer to FormMedian.ApplyMedianFilter().)
        Dim waveletsUsed As Boolean, waveletRatio As Double, waveletWidth As Long, waveletHeight As Long, waveletRadius As Long
        waveletsUsed = False
        
        If (cRadius > 1) Then
            
            If (workingDIB.GetDIBWidth > cRadius) Or (workingDIB.GetDIBHeight > cRadius) Then
            
                waveletRatio = 100# - (170# * CDbl(cRadius)) ^ 0.4
                
                'Convert the ratio to the range [0, 1], and invert it so that it represents
                ' a multiplication factor.
                If (waveletRatio > 100#) Then waveletRatio = 100#
                If (waveletRatio < 1#) Then waveletRatio = 1#
                waveletRatio = waveletRatio / 100#
                
                'If we produced a valid value, activate wavelet mode
                waveletsUsed = (waveletRatio < 1#)
                
            End If
        
        End If
        
        'If wavelets are a valid option, modify the median radius accordingly.
        If waveletsUsed Then
        
            waveletWidth = Int(CDbl(workingDIB.GetDIBWidth) * waveletRatio + 0.5)
            If (waveletWidth < 1) Then waveletWidth = 1
            
            waveletHeight = Int(CDbl(workingDIB.GetDIBHeight) * waveletRatio + 0.5)
            If (waveletHeight < 1) Then waveletHeight = 1
            
            waveletRadius = Int(CDbl(cRadius) * waveletRatio + 0.5)
            If (waveletRadius < 1) Then waveletRadius = 1
            
        Else
            waveletWidth = finalX
        End If
        
        progBarMax = progBarMax + waveletWidth * 4
        progBarOffsets(2) = progBarMax
        
        'ContourDIB uses width of the image
        progBarMax = progBarMax + finalX
        progBarOffsets(3) = progBarMax
        
        'WhiteBalance uses height
        progBarMax = progBarMax + finalY
        
        'Blur the current DIB
        If Filters_Layers.CreateApproximateGaussianBlurDIB(cRadius, srcDIB, workingDIB, 3, toPreview, progBarMax, progBarOffsets(0)) Then
        
            'Use the median filter to round out edges
            Dim medianOK As Boolean
            If waveletsUsed Then
                
                'Produce wavelet copies of the image.  (Actually, we just want a low-frequency version
                ' of the image, because median is specifically designed to remove high-frequency noise!)
                Dim wSrcDIB As pdDIB, wDstDIB As pdDIB
                
                Set wSrcDIB = New pdDIB
                wSrcDIB.CreateBlank waveletWidth, waveletHeight, 32, 0, 0
                GDI_Plus.GDIPlus_StretchBlt wSrcDIB, 0, 0, waveletWidth, waveletHeight, workingDIB, 0, 0, workingDIB.GetDIBWidth, workingDIB.GetDIBHeight, interpolationType:=GP_IM_HighQualityBicubic, isZoomedIn:=True, dstCopyIsOkay:=True
                
                Set wDstDIB = New pdDIB
                wDstDIB.CreateFromExistingDIB wSrcDIB
                
                medianOK = (Filters_Layers.CreateMedianDIB(waveletRadius, 50, PDPRS_Circle, wSrcDIB, wDstDIB, toPreview, progBarMax, progBarOffsets(1)) <> 0)
                Set wSrcDIB = Nothing
                
                'Paint the wavelet result into "srcDIB", which will be used for the next stage of processing
                srcDIB.ResetDIB 0
                GDI_Plus.GDIPlus_StretchBlt srcDIB, 0, 0, workingDIB.GetDIBWidth, workingDIB.GetDIBHeight, wDstDIB, 0, 0, wDstDIB.GetDIBWidth, wDstDIB.GetDIBHeight, interpolationType:=GP_IM_HighQualityBicubic, dstCopyIsOkay:=True
                Set wDstDIB = Nothing
                
            Else
                medianOK = (Filters_Layers.CreateMedianDIB(cRadius, 50, PDPRS_Circle, workingDIB, srcDIB, toPreview, progBarMax, progBarOffsets(1)) <> 0)
            End If
            
            If medianOK Then
        
                'Next, create a contour of the DIB
                If Filters_Layers.CreateContourDIB(useBlackBackground, srcDIB, workingDIB, toPreview, progBarMax, progBarOffsets(2)) Then
            
                    'Finally, white balance the resulting DIB
                    Filters_Layers.WhiteBalanceDIB 0.01, workingDIB, toPreview, progBarMax, progBarOffsets(3)
                    
                End If
            End If
        End If
    Else
        
        'Calculate a progress bar maximum value.  At present, each sub-function has its own strategy for
        ' calculating progress.
        ReDim progBarOffsets(0 To 2) As Long
        
        'Blur uses 3 iterations in each direction (width/height)
        progBarOffsets(0) = 0
        progBarMax = finalX * 3 + finalY * 3
        progBarOffsets(1) = progBarMax
        
        'ContourDIB uses width of the image
        progBarMax = progBarMax + finalX
        progBarOffsets(2) = progBarMax
        
        'WhiteBalance uses height
        progBarMax = progBarMax + finalY
        
        'Blur the current DIB
        If Filters_Layers.CreateApproximateGaussianBlurDIB(cRadius, workingDIB, srcDIB, 3, toPreview, progBarMax, progBarOffsets(0)) Then
        
            'Next, create a contour of the DIB
            If Filters_Layers.CreateContourDIB(useBlackBackground, srcDIB, workingDIB, toPreview, progBarMax, progBarOffsets(1)) Then
            
                'Finally, white balance the resulting DIB
                Filters_Layers.WhiteBalanceDIB 0.01, workingDIB, toPreview, progBarMax, progBarOffsets(2)
                
            End If
        End If
    End If
    
    Set srcDIB = Nothing
        
    'Pass control to finalizeImageData, which will handle the rest of the rendering using the data inside workingDIB
    EffectPrep.FinalizeImageData toPreview, dstPic

End Sub

Private Sub cmdBar_OKClick()
    Process "Trace contour", , GetLocalParamString(), UNDO_Layer
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
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

Private Sub chkBlackBackground_Click()
    UpdatePreview
End Sub

Private Sub sltThickness_Change()
    UpdatePreview
End Sub

Private Sub UpdatePreview()
    If cmdBar.PreviewsAllowed Then Me.TraceContour GetLocalParamString(), True, pdFxPreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub pdFxPreview_ViewportChanged()
    UpdatePreview
End Sub

Private Function GetLocalParamString() As String
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    
    With cParams
        .AddParam "thickness", sltThickness.Value
        .AddParam "blackbackground", chkBlackBackground.Value
        .AddParam "smoothing", chkSmoothing.Value
    End With
    
    GetLocalParamString = cParams.GetParamString()
    
End Function
