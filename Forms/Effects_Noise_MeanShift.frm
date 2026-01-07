VERSION 5.00
Begin VB.Form FormMeanShift 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   " Mean shift"
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
      Top             =   1560
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   1244
      Caption         =   "radius"
      Min             =   1
      Max             =   50
      Value           =   5
      DefaultValue    =   1
   End
   Begin PhotoDemon.pdSlider sltThreshold 
      Height          =   705
      Left            =   6000
      TabIndex        =   3
      Top             =   2400
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   1244
      Caption         =   "threshold"
      Min             =   1
      Max             =   50
      Value           =   15
      DefaultValue    =   15
   End
   Begin PhotoDemon.pdButtonStrip btsKernelShape 
      Height          =   1095
      Left            =   6000
      TabIndex        =   4
      Top             =   3240
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   1931
      Caption         =   "kernel shape"
   End
End
Attribute VB_Name = "FormMeanShift"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Mean Shift Effect Tool
'Copyright 2015-2026 by Tanner Helland
'Created: 02/October/15
'Last updated: 25/June/20
'Last update: minor perf improvements
'
'Mean shift filter, heavily optimized.  Wiki has a nice summary of this technique:
' https://en.wikipedia.org/wiki/Mean_shift
'
'Note that an "accumulation" technique is used instead of the standard sliding window mechanism.  The pdPixelIterator
' class handles this portion of the code, so all this filter has to do is process the resulting histograms.
'
'As with all area-based filters, this function will be slow inside the IDE.  I STRONGLY recommend compiling before using.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

Public Sub ApplyMeanShiftFilter(ByVal parameterList As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)
    
    'Parse out the parameter list
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString parameterList
    
    Dim mRadius As Long, mThreshold As Long, kernelShape As PD_PixelRegionShape
    mRadius = cParams.GetLong("radius", 1&)
    mThreshold = cParams.GetLong("threshold", 0&)
    kernelShape = cParams.GetLong("kernelshape", PDPRS_Rectangle)
    
    If (Not toPreview) Then Message "Applying mean shift filter..."
    
    'Create a local array and point it at the pixel data of the current image
    Dim dstImageData() As Byte, dstSA As SafeArray2D
    EffectPrep.PrepImageData dstSA, toPreview, dstPic
    
    'Create a second copy of the target DIB.
    ' (This is necessary to prevent processed pixel values from spreading across the image as we go.)
    Dim srcDIB As pdDIB
    Set srcDIB = New pdDIB
    srcDIB.CreateFromExistingDIB workingDIB
    
    'If this is a preview, we need to adjust the kernel radius to match the size of the preview box
    If toPreview Then
        mRadius = mRadius * curDIBValues.previewModifier
        If (mRadius < 1) Then mRadius = 1
    End If
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = srcDIB.GetDIBWidth - 1
    finalY = srcDIB.GetDIBHeight - 1
    
    'Just to be safe, make sure the radius isn't larger than the image itself
    If (finalY - initY) < (finalX - initX) Then
        If mRadius > (finalY - initY) Then mRadius = finalY - initY
    Else
        If mRadius > (finalX - initX) Then mRadius = finalX - initX
    End If
    
    'The x-dimension of the image has a stride of (width * 4) for 32-bit images; precalculate this, to spare us some
    ' processing time in the inner loop.
    initX = initX * 4
    finalX = finalX * 4
    
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
    
    'We use an optimized histogram technique for calculating means, which means a lot of intermediate values are required
    Dim rValues(0 To 255) As Long, gValues(0 To 255) As Long, bValues(0 To 255) As Long, aValues(0 To 255) As Long
    Dim r As Long, g As Long, b As Long, a As Long
    Dim lColor As Long, hColor As Long, cCount As Long
    Dim startY As Long, stopY As Long, yStep As Long, i As Long
    
    'Build luts for low- and high-color values.  (These are fixed based on the user-supplied
    ' threshold value.)
    Dim lowLUT(0 To 255) As Byte, highLUT(0 To 255) As Byte
    For x = 0 To 255
        y = x - mThreshold
        If (y < 0) Then y = 0
        lowLUT(x) = y
        y = x + mThreshold
        If (y > 255) Then y = 255
        highLUT(x) = y
    Next x
    
    Dim directionDown As Boolean
    directionDown = True
    
    'Prep the pixel iterator
    Dim cPixelIterator As pdPixelIterator
    Set cPixelIterator = New pdPixelIterator
    
    If cPixelIterator.InitializeIterator(srcDIB, mRadius, mRadius, kernelShape) Then
        
        numOfPixels = cPixelIterator.LockTargetHistograms_RGBA(rValues, gValues, bValues, aValues, False)
        workingDIB.WrapArrayAroundDIB dstImageData, dstSA
        
        'Loop through each pixel in the image, applying the filter as we go
        For x = initX To finalX Step 4
            
            'Based on the direction we're traveling, reverse the interior loop boundaries as necessary.
            If directionDown Then
                startY = initY
                stopY = finalY
                yStep = 1
            Else
                startY = finalY
                stopY = initY
                yStep = -1
            End If
            
            'Process the next column.  This step is pretty much identical to the row steps above (but in a vertical direction, obviously)
            For y = startY To stopY Step yStep
            
                'With a local histogram successfully built for the area surrounding this pixel, we now need to find the
                ' shifted mean value.  We do this by averaging only the pixels whose color is within the caller-specified
                ' threshold of this pixel.  This limits the average to similar pixels only.
        
                'Blue
                
                'Start by finding low/high bounds
                lColor = dstImageData(x, y)
                hColor = highLUT(lColor)
                lColor = lowLUT(lColor)
                
                'Search from low to high, tallying colors as we go
                b = 0: cCount = 0
                For i = lColor To hColor
                    a = bValues(i)
                    b = b + i * a
                    cCount = cCount + a
                Next i
                
                'Take the mean of this range of values
                If (cCount > 0) Then b = b \ cCount Else b = 255
                
                'Repeat for green
                lColor = dstImageData(x + 1, y)
                hColor = highLUT(lColor)
                lColor = lowLUT(lColor)
                
                g = 0: cCount = 0
                For i = lColor To hColor
                    a = gValues(i)
                    g = g + i * a
                    cCount = cCount + a
                Next i
                
                If (cCount > 0) Then g = g \ cCount Else g = 255
                
                'Repeat for red
                lColor = dstImageData(x + 2, y)
                hColor = highLUT(lColor)
                lColor = lowLUT(lColor)
                
                r = 0: cCount = 0
                For i = lColor To hColor
                    a = rValues(i)
                    r = r + i * a
                    cCount = cCount + a
                Next i
                
                If (cCount > 0) Then r = r \ cCount Else r = 255
                
                'Finally, apply the results to the image.
                dstImageData(x, y) = b
                dstImageData(x + 1, y) = g
                dstImageData(x + 2, y) = r
                
                'Move the iterator in the correct direction
                If directionDown Then
                    If (y < finalY) Then numOfPixels = cPixelIterator.MoveYDown
                Else
                    If (y > initY) Then numOfPixels = cPixelIterator.MoveYUp
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
        workingDIB.UnwrapArrayFromDIB dstImageData
        
        'Erase our temporary DIB
        srcDIB.EraseDIB
        Set srcDIB = Nothing
    
        'Pass control to finalizeImageData, which will handle the rest of the rendering using the data inside workingDIB
        EffectPrep.FinalizeImageData toPreview, dstPic
        
    End If

End Sub

Private Sub btsKernelShape_Click(ByVal buttonIndex As Long)
    UpdatePreview
End Sub

'OK button
Private Sub cmdBar_OKClick()
    Process "Mean shift", , GetLocalParamString(), UNDO_Layer
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub Form_Load()
    
    'Disable previews while we initialize the dialog
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

Private Sub sltThreshold_Change()
    UpdatePreview
End Sub

Private Sub sltRadius_Change()
    UpdatePreview
End Sub

Private Sub UpdatePreview()
    If cmdBar.PreviewsAllowed Then Me.ApplyMeanShiftFilter GetLocalParamString(), True, pdFxPreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub pdFxPreview_ViewportChanged()
    UpdatePreview
End Sub

Private Function GetLocalParamString() As String
    GetLocalParamString = BuildParamList("radius", sltRadius.Value, "threshold", sltThreshold.Value, "kernelshape", btsKernelShape.ListIndex)
End Function
