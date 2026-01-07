VERSION 5.00
Begin VB.Form FormHarmonicMean 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   " Harmonic mean"
   ClientHeight    =   6540
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11670
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
   ScaleWidth      =   778
   Begin PhotoDemon.pdCommandBar cmdBar 
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5790
      Width           =   11670
      _ExtentX        =   20585
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
      Index           =   0
      Left            =   6000
      TabIndex        =   2
      Top             =   1560
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   1244
      Caption         =   "horizontal strength"
      Min             =   1
      Max             =   500
      ScaleStyle      =   1
      Value           =   5
      DefaultValue    =   5
   End
   Begin PhotoDemon.pdSlider sltRadius 
      Height          =   705
      Index           =   1
      Left            =   6000
      TabIndex        =   3
      Top             =   2400
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   1244
      Caption         =   "vertical strength"
      Min             =   1
      Max             =   500
      ScaleStyle      =   1
      Value           =   5
      DefaultValue    =   5
   End
   Begin PhotoDemon.pdButtonStrip btsKernelShape 
      Height          =   1095
      Left            =   6000
      TabIndex        =   4
      Top             =   3840
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   1931
      Caption         =   "kernel shape"
   End
   Begin PhotoDemon.pdCheckBox chkSynchronize 
      Height          =   375
      Left            =   6120
      TabIndex        =   5
      Top             =   3240
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   661
      Caption         =   "synchronize search radius"
   End
End
Attribute VB_Name = "FormHarmonicMean"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Harmonic mean Tool
'Copyright 2013-2026 by Tanner Helland
'Created: 27/July/17
'Last updated: 25/June/20
'Last update: as radius increases, use increasingly strong wavelet approximation for
'             huge perf boost
'
'This is a heavily optimized "harmonic mean" function.  An accumulation technique is used instead of the standard sliding
' window mechanism.  (See http://web.archive.org/web/20060718054020/http://www.acm.uiuc.edu/siggraph/workshops/wjarosz_convolution_2001.pdf)
' This allows the algorithm to perform quite well, despite being written in pure VB.
'
'That said, it is still unfortunately slow in the IDE.  I STRONGLY recommend compiling the project before applying any
' filter of a large radius (> 20).
'
'Harmonic mean is an edge-preserving noise removal filter.  It calculate the harmonic mean
' (https://en.wikipedia.org/wiki/Harmonic_mean) for a region around each pixel, and sets pixel values accordingly.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

Public Sub ApplyHarmonicMean(ByVal parameterList As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)
    
    'Parse out the parameter list
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString parameterList
    
    Dim hRadius As Long, vRadius As Long, kernelShape As PD_PixelRegionShape
    hRadius = cParams.GetLong("radius-x", 1)
    vRadius = cParams.GetLong("radius-y", hRadius)
    kernelShape = cParams.GetLong("kernelshape", PDPRS_Rectangle)
    
    If (Not toPreview) Then Message "Applying harmonic mean filter..."
        
    'Create a local array and point it at the pixel data of the current image
    Dim dstImageData() As Byte, dstSA As SafeArray2D
    EffectPrep.PrepImageData dstSA, toPreview, dstPic
    
    Dim x As Long, y As Long
        
    'If this is a preview, we need to adjust the kernel radius to match the size of the preview box
    If toPreview Then
        hRadius = Int(hRadius * curDIBValues.previewModifier + 0.5)
        vRadius = Int(vRadius * curDIBValues.previewModifier + 0.5)
    End If
    
    'Limit radius to the size of the underlying image
    If (hRadius > workingDIB.GetDIBWidth) Then hRadius = workingDIB.GetDIBWidth
    If (vRadius > workingDIB.GetDIBHeight) Then vRadius = workingDIB.GetDIBHeight
    
    'Final sanity check
    If (hRadius < 1) Then hRadius = 1
    If (vRadius < 1) Then vRadius = 1
    
    'The number of pixels in the current median box are tracked dynamically.
    Dim numOfPixels As Long
    numOfPixels = 0
            
    'Harmonic mean filtering takes a lot of variables
    Dim rValues() As Long, gValues() As Long, bValues() As Long, aValues() As Long
    ReDim rValues(0 To 255) As Long
    ReDim gValues(0 To 255) As Long
    ReDim bValues(0 To 255) As Long
    ReDim aValues(0 To 255) As Long
    
    Dim startY As Long, stopY As Long, yStep As Long, i As Long
    
    Dim directionDown As Boolean
    directionDown = True
    
    'Normally we use doubles (as they're generally faster than singles in VB6 since they
    ' use the old x87 fp pathways), but this function is extremely cache constrained,
    ' so every byte saved helps.
    Dim pxSum As Single, pxCount As Long
    Dim finalR As Single, finalG As Single, finalB As Single
    
    'Prebuild a lookup table for all possible (1 / i) values.  To allow us to process the case of i=0
    ' (e.g. black pixels), we increment all values by 1.0, then subtract 1.0 in the inner loop, after the
    ' mean has been calculated.
    Dim oneDiv(0 To 255) As Single
    For i = 0 To 255
        oneDiv(i) = 1! / (i + 1)
    Next i
    
    'This filter is incredibly slow.  To improve performance at very large radii, we use a
    ' wavelet approach.  Figure out how far we can shrink the image while still maintaining an
    ' acceptable level of quality.
    Dim waveletsUsed As Boolean, waveletRatio As Double, waveletWidth As Long, waveletHeight As Long
    waveletsUsed = False
    
    'Use the smallest directional radius
    Dim mRadius As Long
    mRadius = hRadius
    If (vRadius < mRadius) Then mRadius = vRadius
    
    If (mRadius > 1) Then
        
        'First, if the image is sufficiently small, don't optimize with wavelets as it's a
        ' waste of time.
        If (workingDIB.GetDIBWidth > mRadius) Or (workingDIB.GetDIBHeight > mRadius) Then
        
            'If the image is larger than the underlying radius, use a wavelet approximation.
            ' See FormMedian.ApplyMedianEffect() for details on this formula.
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
    ' modify the effect radius accordingly.
    If waveletsUsed Then
    
        waveletWidth = Int(CDbl(workingDIB.GetDIBWidth) * waveletRatio + 0.5)
        If (waveletWidth < 1) Then waveletWidth = 1
        
        waveletHeight = Int(CDbl(workingDIB.GetDIBHeight) * waveletRatio + 0.5)
        If (waveletHeight < 1) Then waveletHeight = 1
        
        hRadius = Int(CDbl(hRadius) * waveletRatio + 0.5)
        If (hRadius < 1) Then hRadius = 1
        
        vRadius = Int(CDbl(vRadius) * waveletRatio + 0.5)
        If (vRadius < 1) Then vRadius = 1
        
        srcDIB.CreateBlank waveletWidth, waveletHeight, 32, 0, 0
        GDI_Plus.GDIPlus_StretchBlt srcDIB, 0, 0, waveletWidth, waveletHeight, workingDIB, 0, 0, workingDIB.GetDIBWidth, workingDIB.GetDIBHeight, interpolationType:=GP_IM_HighQualityBicubic, isZoomedIn:=True, dstCopyIsOkay:=True
        
        Set dstDIB = New pdDIB
        dstDIB.CreateFromExistingDIB srcDIB
        
    'If the radius isn't small enough to warrant a wavelet approach, simply mirror the
    ' existing image as-is.  Note also that we can also cheat and "paint" the finished result
    ' directly into the working DIB provided us by PD's effect engine.
    Else
        srcDIB.CreateFromExistingDIB workingDIB
        Set dstDIB = workingDIB
    End If
    
    dstDIB.WrapArrayAroundDIB dstImageData, dstSA
    
    'The x-dimension of the image has a stride of (width * 4) for 32-bit images; precalculate this,
    ' to spare us some processing time in the inner loop.
    Dim finalX As Long, finalY As Long
    finalX = (srcDIB.GetDIBWidth - 1) * 4
    finalY = (srcDIB.GetDIBHeight - 1)
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    If (Not toPreview) Then
        SetProgBarMax finalX
        progBarCheck = ProgressBars.FindBestProgBarValue()
    End If
    
    'Prep the pixel iterator
    Dim cPixelIterator As pdPixelIterator
    Set cPixelIterator = New pdPixelIterator
    
    If cPixelIterator.InitializeIterator(srcDIB, hRadius, vRadius, kernelShape) Then
    
        numOfPixels = cPixelIterator.LockTargetHistograms_RGBA(rValues, gValues, bValues, aValues, False)
        
        'Loop through each pixel in the image, applying the filter as we go
        For x = 0 To finalX Step 4
            
            'Based on the direction we're traveling, reverse the interior loop boundaries as necessary.
            If directionDown Then
                startY = 0
                stopY = finalY
                yStep = 1
            Else
                startY = finalY
                stopY = 0
                yStep = -1
            End If
            
            'Process the next column.  This step is pretty much identical to the row steps above (but in a vertical direction, obviously)
            For y = startY To stopY Step yStep
                
                'With histograms successfully calculated, we can now find the harmonic mean for this pixel.
                
                'Loop through each color component histogram, and average all non-zero pixels found
                pxSum = 0!
                pxCount = 0
                
                For i = 0 To 255
                    pxSum = pxSum + rValues(i) * oneDiv(i)
                    pxCount = pxCount + rValues(i)
                Next i
                If (pxSum > 0!) Then finalR = pxCount / pxSum Else finalR = 1!
                
                'Repeat for green and blue
                pxSum = 0!
                pxCount = 0
                
                For i = 0 To 255
                    pxSum = pxSum + gValues(i) * oneDiv(i)
                    pxCount = pxCount + gValues(i)
                Next i
                If (pxSum > 0!) Then finalG = pxCount / pxSum Else finalG = 1!
                
                pxSum = 0!
                pxCount = 0
                
                For i = 0 To 255
                    pxSum = pxSum + bValues(i) * oneDiv(i)
                    pxCount = pxCount + bValues(i)
                Next i
                
                If (pxSum > 0!) Then finalB = pxCount / pxSum Else finalB = 1!
                
                'Subtract one from the calculated average (which is how we compensate for black pixels),
                ' then perform a failsafe upper-bound check.  (Lower bound is guaranteed safe.)
                finalR = finalR - 1!
                finalG = finalG - 1!
                finalB = finalB - 1!
                If (finalR > 255!) Then finalR = 255!
                If (finalR < 0!) Then finalR = 0!
                If (finalG > 255!) Then finalG = 255!
                If (finalG < 0!) Then finalG = 0!
                If (finalB > 255!) Then finalB = 255!
                If (finalB < 0!) Then finalB = 0!
                
                'Update the pixel data in the destination image with our final result(s)
                dstImageData(x, y) = finalB
                dstImageData(x + 1, y) = finalG
                dstImageData(x + 2, y) = finalR
                
                'Move the iterator in the correct direction
                If directionDown Then
                    If (y < finalY) Then numOfPixels = cPixelIterator.MoveYDown
                Else
                    If (y > 0) Then numOfPixels = cPixelIterator.MoveYUp
                End If
                
            Next y
            
            'Reverse y-directionality on each pass
            directionDown = Not directionDown
            If (x < finalX) Then numOfPixels = cPixelIterator.MoveXRight
            
            'Update the progress bar every (progBarCheck) lines
            If (Not toPreview) Then
                If (x And progBarCheck) = 0 Then
                    If Interface.UserPressedESC() Then Exit For
                    SetProgBarVal x
                End If
            End If
                
        Next x
        
        'Release the pixel iterator
        cPixelIterator.ReleaseTargetHistograms_RGBA rValues, gValues, bValues, aValues
        
        'Release our local array that points to the target DIB
        dstDIB.UnwrapArrayFromDIB dstImageData
        
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
        
        'Pass control to finalizeImageData, which will handle the rest of the rendering using the data inside workingDIB
        EffectPrep.FinalizeImageData toPreview, dstPic
        
    End If

End Sub

Private Sub btsKernelShape_Click(ByVal buttonIndex As Long)
    UpdatePreview
End Sub

Private Sub chkSynchronize_Click()
    If chkSynchronize.Value Then sltRadius(1).Value = sltRadius(0).Value
End Sub

'OK button
Private Sub cmdBar_OKClick()
    Process "Harmonic mean", , GetLocalParamString(), UNDO_Layer
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub Form_Load()

    'Disable previews while we initialize everything
    cmdBar.SetPreviewStatus False
    
    'Populate the kernel shape box with whatever shapes PD currently supports
    Interface.PopKernelShapeButtonStrip btsKernelShape, PDPRS_Rectangle
    
    'Apply translations and visual themes
    ApplyThemeAndTranslations Me, True, True
    cmdBar.SetPreviewStatus True
    UpdatePreview
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub UpdatePreview()
    If cmdBar.PreviewsAllowed Then ApplyHarmonicMean GetLocalParamString(), True, pdFxPreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub pdFxPreview_ViewportChanged()
    UpdatePreview
End Sub

Private Sub sltRadius_Change(Index As Integer)
    
    If chkSynchronize.Value Then
        If (sltRadius(Abs(Index - 1)).Value <> sltRadius(Index).Value) Then
            cmdBar.SetPreviewStatus False
            sltRadius(Abs(Index - 1)).Value = sltRadius(Index).Value
            cmdBar.SetPreviewStatus True
        End If
    End If
    
    UpdatePreview
    
End Sub

Private Function GetLocalParamString() As String
    GetLocalParamString = BuildParamList("radius-x", sltRadius(0).Value, "radius-y", sltRadius(1).Value, "kernelshape", btsKernelShape.ListIndex)
End Function
