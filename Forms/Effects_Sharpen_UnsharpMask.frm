VERSION 5.00
Begin VB.Form FormUnsharpMask 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   " Unsharp mask"
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
   Begin PhotoDemon.pdFxPreviewCtl pdFxPreview 
      Height          =   5625
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9922
   End
   Begin PhotoDemon.pdSlider sltThreshold 
      Height          =   705
      Left            =   6000
      TabIndex        =   2
      Top             =   2880
      Width           =   5445
      _ExtentX        =   9604
      _ExtentY        =   1244
      Caption         =   "threshold"
      Max             =   255
   End
   Begin PhotoDemon.pdSlider sltAmount 
      Height          =   705
      Left            =   6000
      TabIndex        =   3
      Top             =   1920
      Width           =   5445
      _ExtentX        =   9604
      _ExtentY        =   1244
      Caption         =   "amount"
      Min             =   0.1
      SigDigits       =   1
      Value           =   1
      NotchPosition   =   2
      NotchValueCustom=   1
   End
   Begin PhotoDemon.pdSlider sltRadius 
      Height          =   705
      Left            =   6000
      TabIndex        =   4
      Top             =   960
      Width           =   5445
      _ExtentX        =   9604
      _ExtentY        =   1244
      Caption         =   "radius"
      Min             =   0.1
      Max             =   200
      SigDigits       =   1
      Value           =   1
      DefaultValue    =   1
   End
   Begin PhotoDemon.pdButtonStrip btsQuality 
      Height          =   1080
      Left            =   6000
      TabIndex        =   5
      Top             =   3840
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   1905
      Caption         =   "mode"
   End
End
Attribute VB_Name = "FormUnsharpMask"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Unsharp Masking Tool
'Copyright 2001-2026 by Tanner Helland
'Created: 03/March/01
'Last updated: 27/July/17
'Last update: performance improvements, migrate to XML params
'
'To my knowledge, this tool is the first of its kind in VB6 - a variable radius Unsharp Mask filter
' that utilizes all three traditional controls (radius, amount, and threshold) and is based on a
' true Gaussian kernel.

'The use of separable kernels makes this much, much faster than a standard unsharp mask function.  The
' exact speed gain for a P x Q kernel is PQ/(P + Q) - so for a radius of 4 (which is an actual kernel
' of 9x9) the processing time is 4.5x faster.  For a radius of 100, this is 100x faster than a
' traditional method.
'
'Despite this, it's still quite slow in the IDE.  I STRONGLY recommend compiling the project before
' applying any action at a large radius.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Convolve an image using a gaussian kernel (separable implementation!)
'Input: radius of the blur (min 1, no real max - but the scroll bar is maxed at 200 presently)
Public Sub UnsharpMask(ByVal effectParams As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)
        
    If (Not toPreview) Then Message "Applying unsharp mask (step %1 of %2)...", 1, 2
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString effectParams
    
    Dim umRadius As Double, umAmount As Double, umThreshold As Long, gaussQuality As Long
    
    With cParams
        umRadius = .GetDouble("radius", sltRadius.Value)
        umAmount = .GetDouble("amount", sltAmount.Value)
        umThreshold = .GetLong("threshold", sltThreshold.Value)
        gaussQuality = .GetLong("quality", btsQuality.ListIndex)
    End With
    
    'Threshold is presented on the range [0, 255] but we don't need that level of resolution on the inner loop.
    ' Shrink it accordingly.
    umThreshold = umThreshold \ 5
    
    'Create a local array and point it at the pixel data of the current image
    Dim dstSA As SafeArray2D
    EffectPrep.PrepImageData dstSA, toPreview, dstPic
    
    'Create a second local array.  This will contain the a copy of the current image, and we will use it as our source reference
    ' (This is necessary to prevent blurred pixel values from spreading across the image as we go.)
    Dim srcDIB As pdDIB
    Set srcDIB = New pdDIB
    srcDIB.CreateFromExistingDIB workingDIB
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left
    initY = curDIBValues.Top
    finalX = curDIBValues.Right
    finalY = curDIBValues.Bottom
    
    'If this is a preview, we need to adjust the kernel radius to match the size of the preview box
    If toPreview Then
        umRadius = umRadius * curDIBValues.previewModifier
        If (umRadius < 0.1) Then umRadius = 0.1
    End If
    
    'I almost always recommend quality over speed for PD tools, but in this case, the fast option is SO much faster,
    ' and the results so indistinguishable (3% different according to the Central Limit Theorem:
    ' https://www.khanacademy.org/math/probability/statistics-inferential/sampling_distribution/v/central-limit-theorem?playlist=Statistics
    ' ), that I recommend the faster methods instead.
    Dim gaussBlurSuccess As Long
    gaussBlurSuccess = 0
    
    Dim progBarCalculation As Long
    progBarCalculation = 0
    
    'Previous versions of this filter supported an extremely slow (but exact) gaussian blur routine; we now substitute
    ' IIR filtering for that approach, as the output is nearly identical but many times faster.
    If (gaussQuality > 1) Then gaussQuality = 1
    
    Select Case gaussQuality
    
        '3 iteration box blur
        Case 0
            progBarCalculation = finalY * 3 + finalX * 3
            gaussBlurSuccess = CreateApproximateGaussianBlurDIB(umRadius, workingDIB, srcDIB, 3, toPreview, progBarCalculation + finalY)
        
        'IIR Gaussian estimation
        Case Else
            progBarCalculation = finalY + finalX
            gaussBlurSuccess = Filters_Area.GaussianBlur_Deriche(srcDIB, umRadius, 3, toPreview, progBarCalculation + finalY)
        
    End Select
    
    'Assuming the blur was created successfully, proceed with the masking portion of the filter.
    If (gaussBlurSuccess <> 0) Then
    
        'Now that we have a gaussian DIB created in workingDIB, we can point arrays toward it and the source DIB
        Dim dstImageData() As Byte, dstSA1D As SafeArray1D
        Dim srcImageData() As Byte, srcSA1D As SafeArray1D
        
        'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
        ' based on the size of the area to be processed.
        Dim progBarCheck As Long
        progBarCheck = ProgressBars.FindBestProgBarValue()
            
        If (Not toPreview) Then Message "Applying unsharp mask (step %1 of %2)...", 2, 2
            
        'ScaleFactor is used to apply the unsharp mask.  Maximum strength can be any value, but PhotoDemon locks it at 10.
        Dim scaleFactor As Double, invScaleFactor As Double
        scaleFactor = umAmount + 1#
        invScaleFactor = 1# - scaleFactor
    
        Dim blendVal As Double
        
        'More color variables - in this case, sums for each color component
        Dim r As Long, g As Long, b As Long, a As Long
        Dim r2 As Long, g2 As Long, b2 As Long, a2 As Long
        Dim newR As Long, newG As Long, newB As Long, newA As Long
        Dim tLumDelta As Long
        Const ONE_DIV_255 As Double = 1# / 255#
        
        'Wrap 1D arrays around the source and destination images
        workingDIB.WrapArrayAroundScanline dstImageData, dstSA1D, 0
        srcDIB.WrapArrayAroundScanline srcImageData, srcSA1D, 0
        
        Dim dstDibPointer As Long, dstDibStride As Long
        dstDibPointer = dstSA1D.pvData
        dstDibStride = dstSA1D.cElements
        
        Dim srcDibPointer As Long, srcDibStride As Long
        srcDibPointer = srcSA1D.pvData
        srcDibStride = srcSA1D.cElements
        
        initX = initX * 4
        finalX = finalX * 4
        
        'The final step of the smart blur function is to find edges, and replace them with the blurred data as necessary
        For y = initY To finalY
            dstSA1D.pvData = dstDibPointer + dstDibStride * y
            srcSA1D.pvData = srcDibPointer + srcDibStride * y
        For x = initX To finalX Step 4
            
            'Retrieve the original image's pixels
            b = dstImageData(x)
            g = dstImageData(x + 1)
            r = dstImageData(x + 2)
            a = dstImageData(x + 3)
            
            'Now, retrieve the gaussian pixels
            b2 = srcImageData(x)
            g2 = srcImageData(x + 1)
            r2 = srcImageData(x + 2)
            a2 = srcImageData(x + 3)
            
            'Calculate a delta for the threshold comparison
            tLumDelta = Abs(Colors.GetHQLuminance(r, g, b) - Colors.GetHQLuminance(r2, g2, b2))
                            
            'If the delta is below the specified threshold, sharpen it
            If (tLumDelta > umThreshold) Then
                            
                newR = (scaleFactor * r) + (invScaleFactor * r2)
                If (newR > 255) Then newR = 255
                If (newR < 0) Then newR = 0
                    
                newG = (scaleFactor * g) + (invScaleFactor * g2)
                If (newG > 255) Then newG = 255
                If (newG < 0) Then newG = 0
                    
                newB = (scaleFactor * b) + (invScaleFactor * b2)
                If (newB > 255) Then newB = 255
                If (newB < 0) Then newB = 0
                
                newA = (scaleFactor * a) + (invScaleFactor * a2)
                If (newA > 255) Then newA = 255
                If (newA < 0) Then newA = 0
                    
                blendVal = tLumDelta * ONE_DIV_255
                
                'Feather the results
                dstImageData(x) = Colors.BlendColors(newB, b, blendVal)
                dstImageData(x + 1) = Colors.BlendColors(newG, g, blendVal)
                dstImageData(x + 2) = Colors.BlendColors(newR, r, blendVal)
                dstImageData(x + 3) = BlendColors(newA, a, blendVal)
                
            End If
                    
        Next x
            If (Not toPreview) Then
                If (y And progBarCheck) = 0 Then
                    If Interface.UserPressedESC() Then Exit For
                    SetProgBarVal progBarCalculation + y
                End If
            End If
        Next y
        
        PutMem4 VarPtrArray(srcImageData), 0&
        PutMem4 VarPtrArray(dstImageData), 0&
        
        Set srcDIB = Nothing
        
    End If
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    EffectPrep.FinalizeImageData toPreview, dstPic
        
End Sub

Private Sub btsQuality_Click(ByVal buttonIndex As Long)
    UpdatePreview
End Sub

Private Sub cmdBar_OKClick()
    Process "Unsharp mask", , GetLocalParamString(), UNDO_Layer
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub Form_Load()
    
    'Suspend previews until the dialog is fully loaded
    cmdBar.SetPreviewStatus False
    
    'Populate the quality selector
    btsQuality.AddItem "fast", 0
    btsQuality.AddItem "precise", 1
    btsQuality.ListIndex = 0
    
    'Apply visual themes to the form
    ApplyThemeAndTranslations Me, True, True
    cmdBar.SetPreviewStatus True
    UpdatePreview
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub sltAmount_Change()
    UpdatePreview
End Sub

Private Sub sltRadius_Change()
    UpdatePreview
End Sub

Private Sub sltThreshold_Change()
    UpdatePreview
End Sub

Private Sub UpdatePreview()
    If cmdBar.PreviewsAllowed Then UnsharpMask GetLocalParamString(), True, pdFxPreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub pdFxPreview_ViewportChanged()
    UpdatePreview
End Sub

Private Function GetLocalParamString() As String
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    
    With cParams
        .AddParam "radius", sltRadius.Value
        .AddParam "amount", sltAmount.Value
        .AddParam "threshold", sltThreshold.Value
        .AddParam "quality", btsQuality.ListIndex
    End With
    
    GetLocalParamString = cParams.GetParamString()
    
End Function
