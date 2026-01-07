VERSION 5.00
Begin VB.Form FormAnisotropic 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   " Anisotropic diffusion"
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
   Begin PhotoDemon.pdButtonStrip btsDirection 
      Height          =   975
      Left            =   6000
      TabIndex        =   3
      Top             =   1560
      Width           =   5895
      _ExtentX        =   10186
      _ExtentY        =   1085
      Caption         =   "directionality"
   End
   Begin PhotoDemon.pdSlider sltFlow 
      Height          =   705
      Left            =   6000
      TabIndex        =   2
      Top             =   3600
      Width           =   5880
      _ExtentX        =   10372
      _ExtentY        =   1270
      Caption         =   "gradient flow"
      Min             =   1
      Max             =   100
      Value           =   50
      DefaultValue    =   50
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
   Begin PhotoDemon.pdCommandBar cmdBar 
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
      TabIndex        =   4
      Top             =   4440
      Width           =   5880
      _ExtentX        =   10372
      _ExtentY        =   1270
      Caption         =   "strength"
      Max             =   100
      Value           =   50
      DefaultValue    =   50
   End
   Begin PhotoDemon.pdSlider sltIterations 
      Height          =   705
      Left            =   6000
      TabIndex        =   5
      Top             =   2760
      Width           =   5880
      _ExtentX        =   10372
      _ExtentY        =   1270
      Caption         =   "iterations"
      Min             =   1
      Max             =   16
      Value           =   1
      DefaultValue    =   1
   End
   Begin PhotoDemon.pdButtonStrip btsEmphasis 
      Height          =   975
      Left            =   6000
      TabIndex        =   6
      Top             =   480
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1720
      Caption         =   "model"
   End
End
Attribute VB_Name = "FormAnisotropic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Anisotropic Diffusion dialog
'Copyright 2015-2026 by Tanner Helland
'Created: 11/December/15
'Last updated: 15/May/19
'Last update: minor code clean-up
'
'Relevant wikipedia link:
' https://en.wikipedia.org/wiki/Anisotropic_diffusion
'
'Original paper by Perona and Malik, which served as the basis for this implementation:
' http://authors.library.caltech.edu/6498/1/PERieeetpami90.pdf
'
'For a nice look at potential anisotropic filtering applications, check out this lovely page:
' http://www.cs.utah.edu/~manasi/coursework/cs7960/p2/project2.html
'
'This is pretty much a reference implementation of Perona-Malik's original paper.  I've tweaked some of the input
' ranges (and the corresponding names presented to the user) to try and make the filter a bit more accessible to
' beginners, but suggestions for further improvement are obviously welcome.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
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
Public Sub ApplyAnisotropicDiffusion(ByVal parameterList As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)
    
    'Parse out the parameter list
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString parameterList
    
    Dim adDirection As Long, adIterations As Long, adOption As Long
    adDirection = cParams.GetLong("direction", 0&)
    adIterations = cParams.GetLong("iterations", 1&)
    adOption = cParams.GetLong("option", 0&)
    
    'Kappa is an input parameter specific to this function.  Values of 25-30 seem common, but in their original paper,
    ' Perona and Malik refuse to suggest a specific value (argh).
    '
    'PD currently accepts input on the range [1-100], which it then redistributes to [60, 1], which gives a wide
    ' variety of potential output.
    Dim adKappa As Double
    adKappa = cParams.GetDouble("flow", 50#)
    
    If adOption = 0 Then
        adKappa = (100 - (adKappa - 1)) * 2 + 1
    Else
        adKappa = (adKappa - 1) * 0.5
        adKappa = adKappa + 1
    End If
    
    Dim adStrength As Double
    adStrength = cParams.GetDouble("strength", 100#)
    adStrength = adStrength / 100
    
    'Based on the supported direction, create simple boolean values that determine what directions we calculate
    ' in the inner pixel loop.
    Dim adCardinal As Boolean, adOrdinal As Boolean
    adCardinal = (adDirection = 0) Or (adDirection = 2)
    adOrdinal = (adDirection = 1) Or (adDirection = 2)
    
    'Lambda is effectively the "strength" of the final calculation.  Its maximal value should not be larger than
    ' the number of pixels processed (1/4 for either 4-way filter, or 1/8 for the full 8-way filter).
    Dim lambda As Double
    If adCardinal And adOrdinal Then lambda = 1# / 8# Else lambda = 1# / 4#
    lambda = lambda * adStrength
    
    'Conduction values are constant, given a difference on the range [-255, 255], but they are calculated differently
    ' based on the "option" parameter given in the original paper.
    Dim conduction() As Single
    ReDim conduction(-255 To 255) As Single
    
    Dim i As Long, tmpFloat As Double
    For i = -255 To 255
        tmpFloat = (CDbl(i) / adKappa)
        
        If (adOption = 0) Then
            conduction(i) = -1# * (tmpFloat * tmpFloat)
        Else
            conduction(i) = 1# / (1# + tmpFloat * tmpFloat)
        End If
        
    Next i
    
    'Create a local array and point it at the destination pixel data
    Dim dstImageData() As Byte, dstSA As SafeArray2D
    EffectPrep.PrepImageData dstSA, toPreview, dstPic
    workingDIB.WrapArrayAroundDIB dstImageData, dstSA
    
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
    
    Dim xStride As Long, xStrideInner As Long, quickY As Long
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long, progBarOffset As Long
    
    If (Not toPreview) Then
        SetProgBarMax finalX * adIterations
        progBarCheck = ProgressBars.FindBestProgBarValue()
        progBarOffset = 0
    End If
    
    'Lots of random calculation variables are required for this
    Dim rDst As Long, gDst As Long, bDst As Long, aDst As Long
    Dim rSrc As Long, gSrc As Long, bSrc As Long, aSrc As Long
    Dim rNew As Long, gNew As Long, bNew As Long, aNew As Long
    Dim rNabla As Long, gNabla As Long, bNabla As Long, aNabla As Long
    Dim rSum As Double, gSum As Double, bSum As Double, aSum As Double
    
    'Because this filter uses iterations, we may be performing the filter a bunch of times in immediate succession
    For i = 1 To adIterations
        
        If (Not toPreview) Then Message "Calculating energy gradients (pass %1 of %2)...", i, adIterations
        
        'Loop through each pixel in the image, converting values as we go
        For x = initX To finalX
            xStride = x * 4
        For y = initY To finalY
        
            'Grab a copy of the original pixel values; these form the basis of all subsequent comparisons
            bDst = dstImageData(xStride, y)
            gDst = dstImageData(xStride + 1, y)
            rDst = dstImageData(xStride + 2, y)
            aDst = dstImageData(xStride + 3, y)
            
            'Reset all comparison values
            rSum = 0
            gSum = 0
            bSum = 0
            aSum = 0
            
            'If cardinal directionality is requested, calculate the minimum out of NESW distances now
            If adCardinal Then
                
                'North
                quickY = y - 1
                bSrc = srcImageData(xStride, quickY)
                gSrc = srcImageData(xStride + 1, quickY)
                rSrc = srcImageData(xStride + 2, quickY)
                aSrc = srcImageData(xStride + 3, quickY)
                
                'Calculate nabla (gradient) and conduction, and add them to the running total
                rNabla = (rSrc - rDst)
                rSum = rSum + rNabla * conduction(rNabla)
                gNabla = (gSrc - gDst)
                gSum = gSum + gNabla * conduction(gNabla)
                bNabla = (bSrc - bDst)
                bSum = bSum + bNabla * conduction(bNabla)
                aNabla = (aSrc - aDst)
                aSum = aSum + aNabla * conduction(aNabla)
                
                'South
                quickY = y + 1
                bSrc = srcImageData(xStride, quickY)
                gSrc = srcImageData(xStride + 1, quickY)
                rSrc = srcImageData(xStride + 2, quickY)
                aSrc = srcImageData(xStride + 3, quickY)
                
                rNabla = (rSrc - rDst)
                rSum = rSum + rNabla * conduction(rNabla)
                gNabla = (gSrc - gDst)
                gSum = gSum + gNabla * conduction(gNabla)
                bNabla = (bSrc - bDst)
                bSum = bSum + bNabla * conduction(bNabla)
                aNabla = (aSrc - aDst)
                aSum = aSum + aNabla * conduction(aNabla)
                
                'West
                xStrideInner = xStride - 4
                bSrc = srcImageData(xStrideInner, y)
                gSrc = srcImageData(xStrideInner + 1, y)
                rSrc = srcImageData(xStrideInner + 2, y)
                aSrc = srcImageData(xStrideInner + 3, y)
                
                rNabla = (rSrc - rDst)
                rSum = rSum + rNabla * conduction(rNabla)
                gNabla = (gSrc - gDst)
                gSum = gSum + gNabla * conduction(gNabla)
                bNabla = (bSrc - bDst)
                bSum = bSum + bNabla * conduction(bNabla)
                aNabla = (aSrc - aDst)
                aSum = aSum + aNabla * conduction(aNabla)
                
                'East
                xStrideInner = xStride + 4
                bSrc = srcImageData(xStrideInner, y)
                gSrc = srcImageData(xStrideInner + 1, y)
                rSrc = srcImageData(xStrideInner + 2, y)
                aSrc = srcImageData(xStrideInner + 3, y)
                
                rNabla = (rSrc - rDst)
                rSum = rSum + rNabla * conduction(rNabla)
                gNabla = (gSrc - gDst)
                gSum = gSum + gNabla * conduction(gNabla)
                bNabla = (bSrc - bDst)
                bSum = bSum + bNabla * conduction(bNabla)
                aNabla = (aSrc - aDst)
                aSum = aSum + aNabla * conduction(aNabla)
                
            End If
            
            'If ordinal directionality is requested, calculate the minimum out of NW/NE/SW/SE distances now
            If adOrdinal Then

                'North-west
                xStrideInner = xStride - 4
                quickY = y - 1
                bSrc = srcImageData(xStrideInner, quickY)
                gSrc = srcImageData(xStrideInner + 1, quickY)
                rSrc = srcImageData(xStrideInner + 2, quickY)
                aSrc = srcImageData(xStrideInner + 3, quickY)
                
                rNabla = (rSrc - rDst)
                rSum = rSum + rNabla * conduction(rNabla)
                gNabla = (gSrc - gDst)
                gSum = gSum + gNabla * conduction(gNabla)
                bNabla = (bSrc - bDst)
                bSum = bSum + bNabla * conduction(bNabla)
                aNabla = (aSrc - aDst)
                aSum = aSum + aNabla * conduction(aNabla)
                
                'North-east
                xStrideInner = xStride + 4
                quickY = y - 1
                bSrc = srcImageData(xStrideInner, quickY)
                gSrc = srcImageData(xStrideInner + 1, quickY)
                rSrc = srcImageData(xStrideInner + 2, quickY)
                aSrc = srcImageData(xStrideInner + 3, quickY)
                
                rNabla = (rSrc - rDst)
                rSum = rSum + rNabla * conduction(rNabla)
                gNabla = (gSrc - gDst)
                gSum = gSum + gNabla * conduction(gNabla)
                bNabla = (bSrc - bDst)
                bSum = bSum + bNabla * conduction(bNabla)
                aNabla = (aSrc - aDst)
                aSum = aSum + aNabla * conduction(aNabla)
                
                'South-west
                xStrideInner = xStride - 4
                quickY = y + 1
                bSrc = srcImageData(xStrideInner, quickY)
                gSrc = srcImageData(xStrideInner + 1, quickY)
                rSrc = srcImageData(xStrideInner + 2, quickY)
                aSrc = srcImageData(xStrideInner + 3, quickY)
                
                rNabla = (rSrc - rDst)
                rSum = rSum + rNabla * conduction(rNabla)
                gNabla = (gSrc - gDst)
                gSum = gSum + gNabla * conduction(gNabla)
                bNabla = (bSrc - bDst)
                bSum = bSum + bNabla * conduction(bNabla)
                aNabla = (aSrc - aDst)
                aSum = aSum + aNabla * conduction(aNabla)
                
                'South-east
                xStrideInner = xStride + 4
                quickY = y + 1
                bSrc = srcImageData(xStrideInner, quickY)
                gSrc = srcImageData(xStrideInner + 1, quickY)
                rSrc = srcImageData(xStrideInner + 2, quickY)
                aSrc = srcImageData(xStrideInner + 3, quickY)
                
                rNabla = (rSrc - rDst)
                rSum = rSum + rNabla * conduction(rNabla)
                gNabla = (gSrc - gDst)
                gSum = gSum + gNabla * conduction(gNabla)
                bNabla = (bSrc - bDst)
                bSum = bSum + bNabla * conduction(bNabla)
                aNabla = (aSrc - aDst)
                aSum = aSum + aNabla * conduction(aNabla)
                
            End If
            
            'We have now calculated full anistropic sums for each color channel.  Take the average of each channel,
            ' and add it to our original pixel value: this is our final diffused value.
            rNew = rDst + lambda * rSum
            gNew = gDst + lambda * gSum
            bNew = bDst + lambda * bSum
            aNew = aDst + lambda * aSum
            
            'Clamp invalid output
            If rNew < 0 Then rNew = 0
            If rNew > 255 Then rNew = 255
            If gNew < 0 Then gNew = 0
            If gNew > 255 Then gNew = 255
            If bNew < 0 Then bNew = 0
            If bNew > 255 Then bNew = 255
            If aNew < 0 Then aNew = 0
            If aNew > 255 Then aNew = 255
                
            'Store the new values
            dstImageData(xStride, y) = bNew
            dstImageData(xStride + 1, y) = gNew
            dstImageData(xStride + 2, y) = rNew
            dstImageData(xStride + 3, y) = aNew
            
        Next y
            If (Not toPreview) Then
                If (x And progBarCheck) = 0 Then
                    If Interface.UserPressedESC() Then Exit For
                    SetProgBarVal progBarOffset + x
                End If
            End If
        Next x
        
        'On each iteration, we must copy over the new bits to the source image
        If (i < adIterations) Then CopyMemoryStrict VarPtr(srcImageData(0, 0)), VarPtr(dstImageData(0, 0)), (finalX - initX + 1) * (finalY - initY + 1) * 4
        If (Not toPreview) Then progBarOffset = finalX * i
        
    Next i
    
    'With our work complete, point all arrays away from their respective DIBs and deallocate any temp copies
    workingDIB.UnwrapArrayFromDIB dstImageData
    srcDIB.UnwrapArrayFromDIB srcImageData
    srcDIB.EraseDIB
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    EffectPrep.FinalizeImageData toPreview, dstPic

End Sub

Private Sub btsDirection_Click(ByVal buttonIndex As Long)
    UpdatePreview
End Sub

Private Sub btsEmphasis_Click(ByVal buttonIndex As Long)
    UpdatePreview
End Sub

Private Sub cmdBar_OKClick()
    Process "Anisotropic diffusion", , GetLocalParamString(), UNDO_Layer
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub cmdBar_ResetClick()
    sltFlow.Value = 50
    sltStrength.Value = 50
    sltIterations.Value = 1
    btsDirection.ListIndex = 0
End Sub

Private Sub Form_Load()
    
    'Disable previews while we initialize the dialog
    cmdBar.SetPreviewStatus False
    
    btsDirection.AddItem "4-way cardinal", 0
    btsDirection.AddItem "4-way ordinal", 1
    btsDirection.AddItem "8-way", 2
    btsDirection.ListIndex = 0
    
    btsEmphasis.AddItem "sharpen", 0
    btsEmphasis.AddItem "smooth", 1
    btsEmphasis.ListIndex = 0
    
    'Apply translations and visual themes
    ApplyThemeAndTranslations Me, True, True
    
    'Draw a preview of the effect
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
    If cmdBar.PreviewsAllowed Then Me.ApplyAnisotropicDiffusion GetLocalParamString(), True, pdFxPreview
End Sub

Private Sub sltFlow_Change()
    UpdatePreview
End Sub

Private Sub sltIterations_Change()
    UpdatePreview
End Sub

Private Sub sltStrength_Change()
    UpdatePreview
End Sub

Private Function GetLocalParamString() As String
    GetLocalParamString = BuildParamList("direction", btsDirection.ListIndex, "option", btsEmphasis.ListIndex, "iterations", sltIterations.Value, "flow", sltFlow.Value, "strength", sltStrength.Value)
End Function
