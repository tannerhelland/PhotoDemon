VERSION 5.00
Begin VB.Form FormSquish 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   " Squish"
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
      DisableZoomPan  =   -1  'True
   End
   Begin PhotoDemon.pdSlider sltRatioX 
      Height          =   705
      Left            =   6000
      TabIndex        =   3
      Top             =   960
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   1244
      Caption         =   "horizontal strength"
      Min             =   -100
      Max             =   100
      SigDigits       =   1
   End
   Begin PhotoDemon.pdSlider sltRatioY 
      Height          =   705
      Left            =   6000
      TabIndex        =   4
      Top             =   1920
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   1244
      Caption         =   "vertical strength"
      Min             =   -100
      Max             =   100
      SigDigits       =   1
   End
   Begin PhotoDemon.pdSlider sltQuality 
      Height          =   705
      Left            =   6000
      TabIndex        =   5
      Top             =   2880
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   1244
      Caption         =   "quality"
      Min             =   1
      Max             =   5
      Value           =   2
      NotchPosition   =   2
      NotchValueCustom=   2
   End
   Begin PhotoDemon.pdDropDown cboEdges 
      Height          =   735
      Left            =   6000
      TabIndex        =   2
      Top             =   3840
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   1296
      Caption         =   "if pixels lie outside the image..."
   End
End
Attribute VB_Name = "FormSquish"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Squish Distortion (formerly Fixed Perspective)
'Copyright 2013-2026 by Tanner Helland
'Created: 04/April/13
'Last updated: 21/February/19
'Last update: large performance improvements
'
'This tool allows the user to apply forced perspective to an image.  The code is similar (in theory) to the
' shearing algorithm used in FormShear.  Reverse-mapping and supersampling are supported, for those who need
' a very high-quality transform.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

Private Sub cboEdges_Click()
    UpdatePreview
End Sub

'Apply horizontal and/or vertical perspective to an image by shrinking it in one or more directions
' Input: xRatio, a value from -100 to 100 that specifies the horizontal perspective
'        yRatio, same as xRatio but for vertical perspective
Public Sub SquishImage(ByVal effectParams As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)

    If (Not toPreview) Then Message "Squeezing image..."
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString effectParams
    
    Dim xRatio As Double, yRatio As Double
    Dim edgeHandling As Long, superSamplingAmount As Long
    
    With cParams
        xRatio = .GetDouble("xratio", 0#)
        yRatio = .GetDouble("yratio", 0#)
        edgeHandling = .GetLong("edges", cboEdges.ListIndex)
        superSamplingAmount = .GetLong("quality", sltQuality.Value)
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
    Dim nX As Double, nY As Double
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
    
    'Midpoints of the image in the x and y direction
    Dim midX As Double, midY As Double
    midX = CDbl(finalX - initX) / 2#
    midX = midX + initX
    midY = CDbl(finalY - initY) / 2#
    midY = midY + initY
    
    'Convert xRatio and yRatio to [0, 1] range
    xRatio = xRatio / 100#
    yRatio = yRatio / 100#
    
    'Store region width and height as floating-point
    Dim iWidth As Double, iHeight As Double
    iWidth = finalX - initX
    iHeight = finalY - initY
    
    'If the source image is very small, abandon ship!
    If (finalY < 1) Or (finalX < 1) Then
        EffectPrep.FinalizeImageData toPreview, dstPic
        Exit Sub
    End If
    
    'Build a look-up table for horizontal line size and offset
    Dim leftX() As Double, lineWidth() As Double
    ReDim leftX(initY To finalY) As Double
    ReDim lineWidth(initY To finalY) As Double
    
    For y = initY To finalY
        If (xRatio >= 0) Then
            leftX(y) = ((finalY - y) / finalY) * midX * xRatio
        Else
            leftX(y) = (y / finalY) * midX * -xRatio
        End If
        lineWidth(y) = iWidth - (leftX(y) * 2#)
        If lineWidth(y) = 0# Then lineWidth(y) = 0.000000001
        lineWidth(y) = 1# / lineWidth(y)
    Next y
    
    'Do the same for vertical line size and offset
    Dim topY() As Double, lineHeight() As Double
    ReDim topY(initX To finalX) As Double
    ReDim lineHeight(initX To finalX) As Double
    
    For x = initX To finalX
        If yRatio >= 0 Then
            topY(x) = ((finalX - x) / finalX) * midY * yRatio
        Else
            topY(x) = (x / finalX) * midY * -yRatio
        End If
        lineHeight(x) = iHeight - (topY(x) * 2#)
        If lineHeight(x) = 0# Then lineHeight(x) = 0.000000001
        lineHeight(x) = 1# / lineHeight(x)
    Next x
    
    'Source X and Y values, which may or may not be used as part of a bilinear interpolation function
    Dim srcX As Double, srcY As Double
    
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
                
        'Sample a number of source pixels corresponding to the user's supplied quality value; more quality means
        ' more samples, and much better representation in the final output.
        For sampleIndex = 0 To numSamples
            
            'Offset the pixel amount by the supersampling lookup table
            nX = x + ssX(sampleIndex)
            nY = y + ssY(sampleIndex)
            
            'Reverse-map the coordinates back onto the original image (to allow for resampling)
            srcX = ((nX - leftX(nY)) * lineWidth(nY)) * iWidth
            srcY = ((nY - topY(nX)) * lineHeight(nX)) * iHeight
            
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
    Process "Squish", , GetLocalParamString(), UNDO_Layer
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub cmdBar_ResetClick()
    cboEdges.ListIndex = pdeo_Wrap
    sltQuality = 2
End Sub

Private Sub Form_Load()

    'Suspend previews until the dialog is initialized
    cmdBar.SetPreviewStatus False
    
    'I use a central function to populate the edge handling combo box; this way, I can add new methods and have
    ' them immediately available to all distort functions.
    PopDistortEdgeBox cboEdges, pdeo_Wrap
    
    'Apply translations and visual themes
    ApplyThemeAndTranslations Me, True, True
    cmdBar.SetPreviewStatus True
    UpdatePreview
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub sltQuality_Change()
    UpdatePreview
End Sub

Private Sub sltRatioX_Change()
    UpdatePreview
End Sub

Private Sub sltRatioY_Change()
    UpdatePreview
End Sub

'Redraw the on-screen preview of the transformed image
Private Sub UpdatePreview()
    If cmdBar.PreviewsAllowed Then SquishImage GetLocalParamString(), True, pdFxPreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub pdFxPreview_ViewportChanged()
    UpdatePreview
End Sub

Private Function GetLocalParamString() As String
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    
    With cParams
        .AddParam "xratio", sltRatioX.Value
        .AddParam "yratio", sltRatioY.Value
        .AddParam "edges", cboEdges.ListIndex
        .AddParam "quality", sltQuality.Value
    End With
    
    GetLocalParamString = cParams.GetParamString()
    
End Function
