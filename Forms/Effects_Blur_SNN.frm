VERSION 5.00
Begin VB.Form FormSNN 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   " Symmetric nearest-neighbor"
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
   Begin PhotoDemon.pdFxPreviewCtl pdFxPreview 
      Height          =   5625
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9922
   End
   Begin PhotoDemon.pdCommandBar cmdBar 
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5790
      Width           =   11550
      _ExtentX        =   20373
      _ExtentY        =   1323
   End
   Begin PhotoDemon.pdSlider sltStrength 
      Height          =   705
      Left            =   6000
      TabIndex        =   2
      Top             =   2880
      Width           =   5400
      _ExtentX        =   9525
      _ExtentY        =   1244
      Caption         =   "strength"
      Max             =   100
      Value           =   50
      DefaultValue    =   50
   End
   Begin PhotoDemon.pdSlider sltRadius 
      Height          =   705
      Left            =   6000
      TabIndex        =   3
      Top             =   1920
      Width           =   5400
      _ExtentX        =   9525
      _ExtentY        =   1244
      Caption         =   "radius"
      Min             =   1
      Value           =   1
      DefaultValue    =   1
   End
End
Attribute VB_Name = "FormSNN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Symmetric Nearest-Neighbor dialog
'Copyright 2015-2026 by Tanner Helland
'Created: 15/December/15
'Last updated: 26/July/17
'Last update: performance improvements, migrate to XML params
'
'Reference paper by Shahcheraghi, See, and Halin, which served as the basis for this implementation:
' http://www.academia.edu/9883426/Image_Abstraction_Using_Anisotropic_Diffusion_Symmetric_Nearest_Neighbor_Filter
'
'I consider the symmetric nearest-neighbor algorithm to fall into the same class of filters as Kuwahara filtering.
' The algorithm is simple: pixels in some radius [r] around each pixel are compared to their symmetric neighbors
' (e.g. the pixel with same radius [r], but rotated 90 degrees around the source pixel), and the pixel closest in
' color to the source pixel is added to a running total.  After all symmetric pairs have been analyzed, the center
' pixel is replaced with the average of the selected pixels.  (Or in PD's case, the center pixel is blended with the
' new value at some user-specified strength.)
'
'As far as image analysis goes, SNN has some tradeoffs.  It's very good at producing a Kuwahara-like image without
' Kuwahara's telltale "blockiness" due to the square shape of the analysis subregions.  However, it's not a great
' noise reducer, and in fact, it may inadvertently enhance specific types of noise (e.g. salt and pepper pixels).
'
'That said, I think it's a good addition to PD's overall image analysis toolbox.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Apply a standard SNN filter to an image.  At present, the only supported parameters are "radius" and "strength".
Public Sub ApplySymmetricNearestNeighbor(ByVal parameterList As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)
    
    'Create a local array and point it at the destination pixel data
    Dim dstImageData() As Byte, tmpSA As SafeArray2D
    EffectPrep.PrepImageData tmpSA, toPreview, dstPic
    workingDIB.WrapArrayAroundDIB dstImageData, tmpSA
    
    'Parse out the parameter list
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString parameterList
    
    Dim snnRadius As Long
    snnRadius = cParams.GetLong("radius", 1&)
    
    'During previews, the radius shrinks proportional to the viewport difference
    If toPreview Then
        snnRadius = snnRadius * curDIBValues.previewModifier
        If (snnRadius < 1) Then snnRadius = 1
    End If
    
    Dim snnStrength As Single
    snnStrength = cParams.GetSingle("strength", 100!)
    snnStrength = snnStrength / 100!
    
    'Create a second copy of the target DIB.
    ' (This is necessary to prevent processed pixel values from spreading across the image as we go.)
    Dim srcDIB As pdDIB
    Set srcDIB = New pdDIB
    srcDIB.CreateFromExistingDIB workingDIB
    
    Dim srcImageData() As Byte, srcSA As SafeArray2D
    srcDIB.WrapArrayAroundDIB srcImageData, srcSA
    
    'At present, we ignore edge pixels to simplify the filter's implementation; this may be dealt with in the future.
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left + 1
    initY = curDIBValues.Top + 1
    finalX = curDIBValues.Right - 1
    finalY = curDIBValues.Bottom - 1
    
    Dim xOffset As Long, xOffsetInner1 As Long, xOffsetInner2 As Long
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long, progBarOffset As Long
    
    If (Not toPreview) Then
        SetProgBarMax finalY
        progBarCheck = ProgressBars.FindBestProgBarValue()
        progBarOffset = 0
    End If
    
    'Lots of random calculation variables are required for this
    Dim rDst As Long, gDst As Long, bDst As Long, aDst As Long
    Dim rSrc1 As Long, gSrc1 As Long, bSrc1 As Long, aSrc1 As Long
    Dim rSrc2 As Long, gSrc2 As Long, bSrc2 As Long, aSrc2 As Long
    Dim rNew As Long, gNew As Long, bNew As Long, aNew As Long
    Dim rSum As Double, gSum As Double, bSum As Double, aSum As Double
    Dim xInner As Long, xInnerStart As Long, xInnerFinal As Long
    Dim yInner As Long, yInnerStart As Long, yInnerFinal As Long
    Dim xInnerSym As Long, yInnerSym As Long
    Dim snnDist1 As Long, snnDist2 As Long
    Dim numOfPixels As Long
    Dim pxDivisor As Double
    
    'Precalculate all x-offsets so we don't have to calculate them in the inner loop
    Dim xOffsetsPrecalc() As PointLong
    ReDim xOffsetsPrecalc(initX To finalX) As PointLong
    
    For x = initX To finalX
        
        xInnerStart = x - snnRadius
        xInnerFinal = x + snnRadius
        
        If (xInnerStart < 0) Then
            xInnerFinal = xInnerFinal + xInnerStart
            xInnerStart = 0
        ElseIf (xInnerFinal > finalX) Then
            xInnerStart = xInnerStart + (xInnerFinal - finalX)
            xInnerFinal = finalX
        End If
        
        xOffsetsPrecalc(x).x = xInnerStart
        xOffsetsPrecalc(x).y = xInnerFinal
        
    Next x
    
    If (Not toPreview) Then Message "Generating symmetric pixel pairs..."
        
    'Loop through each pixel in the image, converting values as we go
    For y = initY To finalY
        
        'Calculate inner loop bounds
        yInnerStart = y - snnRadius
        yInnerFinal = y + snnRadius
        If (yInnerStart < 0) Then
            yInnerFinal = yInnerFinal + yInnerStart
            yInnerStart = 0
        ElseIf (yInnerFinal > finalY) Then
            yInnerStart = yInnerStart + (yInnerFinal - finalY)
            yInnerFinal = finalY
        End If
        
    For x = initX To finalX
        
        xOffset = x * 4
        
        'Grab a copy of the original pixel values; these form the basis of all subsequent comparisons
        bDst = dstImageData(xOffset, y)
        gDst = dstImageData(xOffset + 1, y)
        rDst = dstImageData(xOffset + 2, y)
        aDst = dstImageData(xOffset + 3, y)
        
        'Reset all comparison values
        rSum = 0
        gSum = 0
        bSum = 0
        aSum = 0
        numOfPixels = 0
        
        'Pull inner loop bounds from our precalculated table
        xInnerStart = xOffsetsPrecalc(x).x
        xInnerFinal = xOffsetsPrecalc(x).y
        
        'SNN can technically be computed in a few different ways; for example, pixels can be analyzed as pairs
        ' (that sit 180 degrees apart), or they can be analyzed as quads (90 degrees part).  For now,
        ' PD takes the pair approach.
        
        'First, we apply a custom set of tests along the x-axis.  This simplifies our next loop a bit.
        For xInner = xOffset + 4 To xInnerFinal * 4 Step 4
            
            'Grab symmetric source pixels
            bSrc1 = srcImageData(xInner, y)
            gSrc1 = srcImageData(xInner + 1, y)
            rSrc1 = srcImageData(xInner + 2, y)
            aSrc1 = srcImageData(xInner + 3, y)
            
            xInnerSym = xOffset + (xOffset - xInner)
            bSrc2 = srcImageData(xInnerSym, y)
            gSrc2 = srcImageData(xInnerSym + 1, y)
            rSrc2 = srcImageData(xInnerSym + 2, y)
            aSrc2 = srcImageData(xInnerSym + 3, y)
            
            'Calculate "similarity" between each pixel and the source pixel
            snnDist1 = Abs(rDst - rSrc1) + Abs(gDst - gSrc1) + Abs(bDst - bSrc1) + Abs(aDst - aSrc1)
            snnDist2 = Abs(rDst - rSrc2) + Abs(gDst - gSrc2) + Abs(bDst - bSrc2) + Abs(aDst - aSrc2)
            
            'Store the closest pixel of the pair
            If (snnDist1 < snnDist2) Then
                rSum = rSum + rSrc1
                gSum = gSum + gSrc1
                bSum = bSum + bSrc1
                aSum = aSum + aSrc1
            Else
                rSum = rSum + rSrc2
                gSum = gSum + gSrc2
                bSum = bSum + bSrc2
                aSum = aSum + aSrc2
            End If
            
            numOfPixels = numOfPixels + 1
            
        Next xInner
        
        'With the x-axis on the same line as the source pixel handled successfully, we can now move to a general-purpose
        ' inner loop that compares symmetrical pixels in both the X and Y direction.
        For yInner = yInnerStart To yInnerFinal
            yInnerSym = y - (yInner - y)
        For xInner = xInnerStart To xInnerFinal
        
            'Calculate symmetry positions
            xInnerSym = x - (xInner - x)
            
            'Grab RGBA values
            xOffsetInner1 = xInner * 4
            bSrc1 = srcImageData(xOffsetInner1, yInner)
            gSrc1 = srcImageData(xOffsetInner1 + 1, yInner)
            rSrc1 = srcImageData(xOffsetInner1 + 2, yInner)
            aSrc1 = srcImageData(xOffsetInner1 + 3, yInner)
            
            xOffsetInner2 = xInnerSym * 4
            bSrc2 = srcImageData(xOffsetInner2, yInnerSym)
            gSrc2 = srcImageData(xOffsetInner2 + 1, yInnerSym)
            rSrc2 = srcImageData(xOffsetInner2 + 2, yInnerSym)
            aSrc2 = srcImageData(xOffsetInner2 + 3, yInnerSym)
            
            'Calculate "similarity" between each pixel and the source pixel
            snnDist1 = Abs(rDst - rSrc1) + Abs(gDst - gSrc1) + Abs(bDst - bSrc1) + Abs(aDst - aSrc1)
            snnDist2 = Abs(rDst - rSrc2) + Abs(gDst - gSrc2) + Abs(bDst - bSrc2) + Abs(aDst - aSrc2)
            
            'Store the closest pixel of the pair
            If (snnDist1 < snnDist2) Then
                rSum = rSum + rSrc1
                gSum = gSum + gSrc1
                bSum = bSum + bSrc1
                aSum = aSum + aSrc1
            Else
                rSum = rSum + rSrc2
                gSum = gSum + gSrc2
                bSum = bSum + bSrc2
                aSum = aSum + aSrc2
            End If
            
            numOfPixels = numOfPixels + 1
            
        Next xInner
        Next yInner
        
        'We have now calculated full SNN sums for each color channel.  Take the average of each channel.
        pxDivisor = 1# / numOfPixels
        rNew = rSum * pxDivisor
        gNew = gSum * pxDivisor
        bNew = bSum * pxDivisor
        aNew = aSum * pxDivisor
        
        'Blend pixels accordingly
        If (snnStrength < 1!) Then
            rNew = BlendLongs(rDst, rNew, snnStrength)
            gNew = BlendLongs(gDst, gNew, snnStrength)
            bNew = BlendLongs(bDst, bNew, snnStrength)
            aNew = BlendLongs(aDst, aNew, snnStrength)
        End If
        
        'Store the new values
        dstImageData(xOffset, y) = bNew
        dstImageData(xOffset + 1, y) = gNew
        dstImageData(xOffset + 2, y) = rNew
        dstImageData(xOffset + 3, y) = aNew
        
    Next x
        If (Not toPreview) Then
            If (y And progBarCheck) = 0 Then
                If Interface.UserPressedESC() Then Exit For
                SetProgBarVal progBarOffset + y
            End If
        End If
    Next y
    
    'With our work complete, point all arrays away from their respective DIBs and deallocate any temp copies
    workingDIB.UnwrapArrayFromDIB dstImageData
    srcDIB.UnwrapArrayFromDIB srcImageData
    srcDIB.EraseDIB
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    EffectPrep.FinalizeImageData toPreview, dstPic

End Sub

'Blend byte1 w/ byte2 based on mixRatio. mixRatio is expected to be a value between 0 and 1.
Private Function BlendLongs(ByVal baseColor As Long, ByVal newColor As Long, ByVal mixRatio As Single) As Long
    BlendLongs = Int((1! - mixRatio) * baseColor + (mixRatio * newColor) + 0.5!)
End Function

Private Sub cmdBar_OKClick()
    Process "Symmetric nearest-neighbor", , GetLocalParamString(), UNDO_Layer
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub cmdBar_ResetClick()
    sltStrength.Value = 50
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

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub pdFxPreview_ViewportChanged()
    UpdatePreview
End Sub

Private Sub UpdatePreview()
    If cmdBar.PreviewsAllowed Then Me.ApplySymmetricNearestNeighbor GetLocalParamString(), True, pdFxPreview
End Sub

Private Sub sltRadius_Change()
    UpdatePreview
End Sub

Private Sub sltStrength_Change()
    UpdatePreview
End Sub

Private Function GetLocalParamString() As String
    GetLocalParamString = BuildParamList("radius", sltRadius.Value, "strength", sltStrength.Value)
End Function
