VERSION 5.00
Begin VB.Form FormSNN 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Symmetric nearest-neighbor"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   12030
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   436
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   802
   ShowInTaskbar   =   0   'False
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
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5790
      Width           =   12030
      _ExtentX        =   21220
      _ExtentY        =   1323
   End
   Begin PhotoDemon.pdSlider sltStrength 
      Height          =   705
      Left            =   6000
      TabIndex        =   2
      Top             =   2880
      Width           =   5880
      _ExtentX        =   10372
      _ExtentY        =   1270
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
      Width           =   5880
      _ExtentX        =   10372
      _ExtentY        =   1270
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
'Copyright 2015-2016 by Tanner Helland
'Created: 15/December/15
'Last updated: 15/December/15
'Last update: initial build
'
'Reference paper by Shahcheraghi, See, and Halin, which served as the basis for this implementation:
' http://www.academia.edu/9883426/Image_Abstraction_Using_Anisotropic_Diffusion_Symmetric_Nearest_Neighbor_Filter
'
'I consider the symmetric nearest-neighbor algorithm to fall into the same class of filters as Kuwahara filtering.
' The algorithm is simple: pixels in some radius [r] around each pixel are compared to their symmetric neighbors
' (e.g. the pixel with same radius [r], but rotated 180 degrees around the source pixel), and the pixel closest in
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
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Apply anisotropic diffusion to an image
'Input: directionality (0 = NESW only, 1 = NE/NW/SE/SW only, 2 - all eight cardinal and ordinal directions)
'       option (0 or 1; a nebulous value proposed by Perona and Malik, where 0 emphasizes high-contrast edges in its
'                       calculations, while 1 emphasizes wide similarly-colored regions over smaller ones)
'       flow ([1, 100] - controls the corresponding kappa value; higher numbers = greater propensity for color flow)
'       strength ([0, 100] - 0 = no change, 100 = fully replace target pixel with anisotropic result,
'                            1-99 = partially blend original and anisotropic result)
Public Sub ApplySymmetricNearestNeighbor(ByVal parameterList As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)
    
    'Parse out the parameter list
    Dim cParams As pdParamXML
    Set cParams = New pdParamXML
    cParams.SetParamString parameterList
    
    Dim snnRadius As Long
    snnRadius = cParams.GetLong("radius", 1&)
    
    'During previews, the radius shrinks proportional to the viewport difference
    If toPreview Then
        snnRadius = snnRadius * curDIBValues.previewModifier
        If snnRadius < 1 Then snnRadius = 1
    End If
    
    Dim snnStrength As Double
    snnStrength = cParams.GetDouble("strength", 100#)
    snnStrength = snnStrength / 100
    
    'Create a local array and point it at the destination pixel data
    Dim dstImageData() As Byte
    Dim tmpSA As SAFEARRAY2D
    PrepImageData tmpSA, toPreview, dstPic
    CopyMemory ByVal VarPtrArray(dstImageData()), VarPtr(tmpSA), 4
    
    'Create a second copy of the target DIB.
    ' (This is necessary to prevent processed pixel values from spreading across the image as we go.)
    Dim srcDIB As pdDIB
    Set srcDIB = New pdDIB
    srcDIB.CreateFromExistingDIB workingDIB
    
    Dim srcImageData() As Byte
    Dim srcSA As SAFEARRAY2D
    PrepSafeArray srcSA, srcDIB
    CopyMemory ByVal VarPtrArray(srcImageData()), VarPtr(srcSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here.
    ' (At present, we ignore edge pixels to simplify the filter's implementation; this will be dealt with in the future.)
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left + 1
    initY = curDIBValues.Top + 1
    finalX = curDIBValues.Right - 1
    finalY = curDIBValues.Bottom - 1
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim xOffset As Long, xOffsetInner1 As Long, xOffsetInner2 As Long, yOffsetInner1 As Long, yOffsetInner2 As Long
    Dim qvDepth As Long
    qvDepth = curDIBValues.BytesPerPixel
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long, progBarOffset As Long
    
    If Not toPreview Then
        SetProgBarMax finalY
        progBarCheck = FindBestProgBarValue()
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
    
    If Not toPreview Then Message "Generating symmetric pixel pairs..."
        
    'Loop through each pixel in the image, converting values as we go
    For y = initY To finalY
    For x = initX To finalX
        
        xOffset = x * qvDepth
        
        'Grab a copy of the original pixel values; these form the basis of all subsequent comparisons
        bDst = dstImageData(xOffset, y)
        gDst = dstImageData(xOffset + 1, y)
        rDst = dstImageData(xOffset + 2, y)
        If qvDepth = 4 Then aDst = dstImageData(xOffset + 3, y)
        
        'Reset all comparison values
        rSum = 0
        gSum = 0
        bSum = 0
        aSum = 0
        numOfPixels = 0
        
        'Calculate inner loop bounds
        xInnerStart = x - snnRadius
        xInnerFinal = x + snnRadius
        If xInnerStart < 0 Then
            xInnerFinal = xInnerFinal + xInnerStart
            xInnerStart = 0
        ElseIf xInnerFinal > finalX Then
            xInnerStart = xInnerStart + (xInnerFinal - finalX)
            xInnerFinal = finalX
        End If
        
        yInnerStart = y - snnRadius
        yInnerFinal = y + snnRadius
        If yInnerStart < 0 Then
            yInnerFinal = yInnerFinal + yInnerStart
            yInnerStart = 0
        ElseIf yInnerFinal > finalY Then
            yInnerStart = yInnerStart + (yInnerFinal - finalY)
            yInnerFinal = finalY
        End If
        
        'SNN can technically be computed in a few different ways; for example, pixels can be analyzed as pairs
        ' (that sit 180 degrees apart), or they can be analyzed as quads (90 degrees part).  For now,
        ' PD takes the pair approach.
        
        'First, we apply a custom set of tests along the x-axis.  This simplifies our next loop a bit.
        For xInner = xOffset + qvDepth To xInnerFinal * qvDepth Step qvDepth
            
            'Grab symmetric source pixels
            bSrc1 = srcImageData(xInner, y)
            gSrc1 = srcImageData(xInner + 1, y)
            rSrc1 = srcImageData(xInner + 2, y)
            If qvDepth = 4 Then aSrc1 = srcImageData(xInner + 3, y)
            
            xInnerSym = xOffset + (xOffset - xInner)
            bSrc2 = srcImageData(xInnerSym, y)
            gSrc2 = srcImageData(xInnerSym + 1, y)
            rSrc2 = srcImageData(xInnerSym + 2, y)
            If qvDepth = 4 Then aSrc2 = srcImageData(xInnerSym + 3, y)
            
            'Calculate "similarity" between each pixel and the source pixel
            snnDist1 = Abs(rDst - rSrc1) + Abs(gDst - gSrc1) + Abs(bDst - bSrc1) + Abs(aDst - aSrc1)
            snnDist2 = Abs(rDst - rSrc2) + Abs(gDst - gSrc2) + Abs(bDst - bSrc2) + Abs(aDst - aSrc2)
            
            'Store the closest pixel of the pair
            If snnDist1 < snnDist2 Then
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
        For xInner = xInnerStart To xInnerFinal
        
            'Calculate symmetry positions
            xInnerSym = x - (xInner - x)
            yInnerSym = y - (yInner - y)
            
            xOffsetInner1 = xInner * qvDepth
            xOffsetInner2 = xInnerSym * qvDepth
            
            'Grab RGBA values
            bSrc1 = srcImageData(xOffsetInner1, yInner)
            gSrc1 = srcImageData(xOffsetInner1 + 1, yInner)
            rSrc1 = srcImageData(xOffsetInner1 + 2, yInner)
            If qvDepth = 4 Then aSrc1 = srcImageData(xOffsetInner1 + 3, yInner)
            
            bSrc2 = srcImageData(xOffsetInner2, yInnerSym)
            gSrc2 = srcImageData(xOffsetInner2 + 1, yInnerSym)
            rSrc2 = srcImageData(xOffsetInner2 + 2, yInnerSym)
            If qvDepth = 4 Then aSrc2 = srcImageData(xOffsetInner2 + 3, yInnerSym)
            
            'Calculate "similarity" between each pixel and the source pixel
            snnDist1 = Abs(rDst - rSrc1) + Abs(gDst - gSrc1) + Abs(bDst - bSrc1) + Abs(aDst - aSrc1)
            snnDist2 = Abs(rDst - rSrc2) + Abs(gDst - gSrc2) + Abs(bDst - bSrc2) + Abs(aDst - aSrc2)
            
            'Store the closest pixel of the pair
            If snnDist1 < snnDist2 Then
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
        rNew = rSum \ numOfPixels
        gNew = gSum \ numOfPixels
        bNew = bSum \ numOfPixels
        aNew = aSum \ numOfPixels
        
        'Blend pixels accordingly
        If snnStrength < 1 Then
            rNew = BlendLongs(rDst, rNew, snnStrength)
            gNew = BlendLongs(gDst, gNew, snnStrength)
            bNew = BlendLongs(bDst, bNew, snnStrength)
            aNew = BlendLongs(aDst, aNew, snnStrength)
        End If
        
        'Store the new values
        dstImageData(xOffset, y) = bNew
        dstImageData(xOffset + 1, y) = gNew
        dstImageData(xOffset + 2, y) = rNew
        If qvDepth = 4 Then dstImageData(xOffset + 3, y) = aNew
        
    Next x
        If Not toPreview Then
            If (y And progBarCheck) = 0 Then
                If UserPressedESC() Then Exit For
                SetProgBarVal progBarOffset + y
            End If
        End If
    Next y
    
    'With our work complete, point all arrays away from their respective DIBs and deallocate any temp copies
    CopyMemory ByVal VarPtrArray(dstImageData), 0&, 4
    Erase dstImageData
    
    CopyMemory ByVal VarPtrArray(srcImageData), 0&, 4
    Erase srcImageData
    
    srcDIB.EraseDIB
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    FinalizeImageData toPreview, dstPic

End Sub

'Blend byte1 w/ byte2 based on mixRatio. mixRatio is expected to be a value between 0 and 1.
Private Function BlendLongs(ByVal baseColor As Long, ByVal newColor As Long, ByRef mixRatio As Double) As Long
    BlendLongs = ((1# - mixRatio) * CDbl(baseColor)) + (mixRatio * CDbl(newColor))
End Function

Private Sub cmdBar_OKClick()
    Process "Symmetric nearest-neighbor", , GetLocalParamString(), UNDO_LAYER
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub cmdBar_ResetClick()
    sltStrength.Value = 50
End Sub

Private Sub Form_Load()
    cmdBar.MarkPreviewStatus False
    ApplyThemeAndTranslations Me
    cmdBar.MarkPreviewStatus True
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




