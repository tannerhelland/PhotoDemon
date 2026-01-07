VERSION 5.00
Begin VB.Form FormModernArt 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   " Modern art"
   ClientHeight    =   6540
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11535
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
   ScaleWidth      =   769
   Begin PhotoDemon.pdCommandBar cmdBar 
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5790
      Width           =   11535
      _ExtentX        =   20346
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
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   1244
      Caption         =   "horizontal strength"
      Min             =   1
      Max             =   200
      Value           =   5
      DefaultValue    =   5
   End
   Begin PhotoDemon.pdSlider sltRadius 
      Height          =   705
      Index           =   1
      Left            =   6000
      TabIndex        =   3
      Top             =   2400
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   1244
      Caption         =   "vertical strength"
      Min             =   1
      Max             =   200
      Value           =   5
      DefaultValue    =   5
   End
   Begin PhotoDemon.pdButtonStrip btsKernelShape 
      Height          =   1095
      Left            =   6000
      TabIndex        =   4
      Top             =   3240
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   1931
      Caption         =   "kernel shape"
   End
End
Attribute VB_Name = "FormModernArt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Modern Art Tool
'Copyright 2013-2026 by Tanner Helland
'Created: 09/Feb/13
'Last updated: 23/November/15
'Last update: convert to XML parameter list
'
'This is a heavily optimized "extreme rank" function.  An accumulation technique is used
' instead of the standard sliding window mechanism.
' (See http://web.archive.org/web/20060718054020/http://www.acm.uiuc.edu/siggraph/workshops/wjarosz_convolution_2001.pdf)
' This allows the algorithm to perform well, despite being written in pure VB.
'
'That said, it is still unfortunately slow in the IDE.  I STRONGLY recommend compiling the
' project before using this tool.
'
'This function works by performing both a minimum and a maximum rank calculation,
' then setting the target pixel to whichever value is further from the current one.
' This leads to a unique cut-out (or stencil?) look.  I'm not sure how much utility it
' provides, but it's fun so I've left it in the project.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Apply a "modern art" filter to the active layer (basically a max/min rank algorithm,
' with some tweaks like shape-specific search regions)
'Input: radius of the median (min 1, no real max - but the scroll bar is maxed at 200 presently)
Public Sub ApplyModernArt(ByVal parameterList As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)
    
    'Parse out the parameter list
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString parameterList
    
    Dim hRadius As Double, vRadius As Double, kernelShape As PD_PixelRegionShape
    hRadius = cParams.GetDouble("radius-x", 1#)
    vRadius = cParams.GetDouble("radius-y", hRadius)
    kernelShape = cParams.GetLong("kernelshape", PDPRS_Rectangle)
    
    If (Not toPreview) Then Message "Applying modern art techniques..."
        
    'Create a local array and point it at the pixel data of the current image
    Dim dstImageData() As Byte, dstSA As SafeArray2D
    EffectPrep.PrepImageData dstSA, toPreview, dstPic
    
    Dim x As Long, y As Long, finalX As Long, finalY As Long
    
    'If this is a preview, we need to adjust the kernel radius to match the size of
    ' the preview box
    If toPreview Then
        hRadius = hRadius * curDIBValues.previewModifier
        vRadius = vRadius * curDIBValues.previewModifier
    End If
    
    'Range-check the radius.  (During previews, the line of code above may cause the radius
    ' to drop to zero.)
    If (hRadius > workingDIB.GetDIBWidth) Then hRadius = workingDIB.GetDIBWidth
    If (vRadius > workingDIB.GetDIBHeight) Then vRadius = workingDIB.GetDIBHeight
    If (hRadius < 1) Then hRadius = 1
    If (vRadius < 1) Then vRadius = 1
    
    'Next, let's see if we can use wavelets to accelerate the filter.  For details on how
    ' this approach works, see FormMedian.ApplyMedianFilter() (which this function is based on).
    Dim minRadius As Long
    minRadius = hRadius
    If (vRadius < minRadius) Then minRadius = vRadius
    
    Dim waveletsUsed As Boolean, waveletRatio As Double, waveletWidth As Long, waveletHeight As Long
    waveletsUsed = False
    
    If (minRadius > 1) Then
        
        'First, if the image is sufficiently small, don't optimize with wavelets as it's a
        ' waste of time.
        If (workingDIB.GetDIBWidth > minRadius) Or (workingDIB.GetDIBHeight > minRadius) Then
        
            'If the image is larger than the underlying radius, use a wavelet approximation.
            waveletRatio = 100# - (170# * CDbl(minRadius)) ^ 0.4
            
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
    
    'The x-dimension of the image has a stride of (width * 4) for 32-bit images; precalculate this, to spare us some
    ' processing time in the inner loop.
    finalX = (srcDIB.GetDIBWidth - 1) * 4
    finalY = srcDIB.GetDIBHeight - 1
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    If (Not toPreview) Then
        SetProgBarMax finalX
        progBarCheck = ProgressBars.FindBestProgBarValue()
    End If
    
    'The number of pixels in the current median box are tracked dynamically.
    Dim numOfPixels As Long
    numOfPixels = 0
            
    'Median filtering takes a lot of variables
    Dim rValues() As Long, gValues() As Long, bValues() As Long, aValues() As Long
    ReDim rValues(0 To 255) As Long
    ReDim gValues(0 To 255) As Long
    ReDim bValues(0 To 255) As Long
    ReDim aValues(0 To 255) As Long
    
    Dim cutoffTotal As Long
    Dim r As Long, g As Long, b As Long
    Dim lowR As Long, lowG As Long, lowB As Long
    Dim highR As Long, highG As Long, highB As Long
    Dim startY As Long, stopY As Long, yStep As Long, i As Long
    
    Dim directionDown As Boolean
    directionDown = True
    
    'Prep the pixel iterator
    Dim cPixelIterator As pdPixelIterator
    Set cPixelIterator = New pdPixelIterator
    
    If cPixelIterator.InitializeIterator(srcDIB, hRadius, vRadius, kernelShape) Then
    
        dstDIB.WrapArrayAroundDIB dstImageData, dstSA
    
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
                
                'With the median box successfully calculated, we can now find the actual median for this pixel.
                
                'Loop through each color component histogram, until we've passed the desired percentile of pixels
                lowR = 0
                lowG = 0
                lowB = 0
                highR = 0
                highG = 0
                highB = 0
                cutoffTotal = numOfPixels \ 100
                If (cutoffTotal < 1) Then cutoffTotal = 1
                
                For i = 0 To 255
                    lowR = lowR + rValues(i)
                    If (lowR >= cutoffTotal) Then
                        lowR = i
                        Exit For
                    End If
                Next i
                
                For i = 255 To 0 Step -1
                    highR = highR + rValues(i)
                    If (highR >= cutoffTotal) Then
                        highR = i
                        Exit For
                    End If
                Next i
                
                For i = 0 To 255
                    lowG = lowG + gValues(i)
                    If (lowG >= cutoffTotal) Then
                        lowG = i
                        Exit For
                    End If
                Next i
                
                For i = 255 To 0 Step -1
                    highG = highG + gValues(i)
                    If (highG >= cutoffTotal) Then
                        highG = i
                        Exit For
                    End If
                Next i
                
                For i = 0 To 255
                    lowB = lowB + bValues(i)
                    If (lowB >= cutoffTotal) Then
                        lowB = i
                        Exit For
                    End If
                Next i
                
                For i = 255 To 0 Step -1
                    highB = highB + bValues(i)
                    If (highB >= cutoffTotal) Then
                        highB = i
                        Exit For
                    End If
                Next i
                
                'Retrieve the original pixel data, and replace it with the processed result
                b = dstImageData(x, y)
                If ((b - lowB) > (highB - b)) Then highB = lowB
                dstImageData(x, y) = highB
                
                g = dstImageData(x + 1, y)
                If ((g - lowG) > (highG - g)) Then highG = lowG
                dstImageData(x + 1, y) = highG
                
                r = dstImageData(x + 2, y)
                If ((r - lowR) > (highR - r)) Then highR = lowR
                dstImageData(x + 2, y) = highR
                
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
        
        'If wavelets were used, we now need to sample the processed result back into workingDIB.
        If waveletsUsed Then
            workingDIB.ResetDIB 0
            GDI_Plus.GDIPlus_StretchBlt workingDIB, 0, 0, workingDIB.GetDIBWidth, workingDIB.GetDIBHeight, dstDIB, 0, 0, dstDIB.GetDIBWidth, dstDIB.GetDIBHeight, interpolationType:=GP_IM_NearestNeighbor, isZoomedIn:=True, dstCopyIsOkay:=True
            Set dstDIB = Nothing
        End If
        
        'Erase our temporary DIB
        Set srcDIB = Nothing
        
        'Pass control to finalizeImageData, which will handle the rest of the rendering using the data inside workingDIB
        EffectPrep.FinalizeImageData toPreview, dstPic
        
    'If an unforseen error occurs, free our unsafe DIB wrapper
    Else
        workingDIB.UnwrapArrayFromDIB dstImageData
    End If

End Sub

Private Sub btsKernelShape_Click(ByVal buttonIndex As Long)
    UpdatePreview
End Sub

'OK button
Private Sub cmdBar_OKClick()
    Process "Modern art", , GetLocalParamString(), UNDO_Layer
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
    If cmdBar.PreviewsAllowed Then ApplyModernArt GetLocalParamString(), True, pdFxPreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub pdFxPreview_ViewportChanged()
    UpdatePreview
End Sub

Private Sub sltRadius_Change(Index As Integer)
    UpdatePreview
End Sub

Private Function GetLocalParamString() As String
    GetLocalParamString = BuildParamList("radius-x", sltRadius(0).Value, "radius-y", sltRadius(1).Value, "kernelshape", btsKernelShape.ListIndex)
End Function
