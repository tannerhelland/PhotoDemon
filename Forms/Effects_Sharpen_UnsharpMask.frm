VERSION 5.00
Begin VB.Form FormUnsharpMask 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Unsharp masking"
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
      Width           =   5925
      _ExtentX        =   10451
      _ExtentY        =   1270
      Caption         =   "threshold"
      Max             =   255
   End
   Begin PhotoDemon.pdSlider sltAmount 
      Height          =   705
      Left            =   6000
      TabIndex        =   3
      Top             =   1920
      Width           =   5925
      _ExtentX        =   10451
      _ExtentY        =   1270
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
      Width           =   5925
      _ExtentX        =   10451
      _ExtentY        =   1270
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
      Width           =   5910
      _ExtentX        =   10425
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
'Copyright 2001-2017 by Tanner Helland
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
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Convolve an image using a gaussian kernel (separable implementation!)
'Input: radius of the blur (min 1, no real max - but the scroll bar is maxed at 200 presently)
Public Sub UnsharpMask(ByVal umRadius As Double, ByVal umAmount As Double, ByVal umThreshold As Long, Optional ByVal gaussQuality As Long = 1, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)
        
    If (Not toPreview) Then Message "Applying unsharp mask (step %1 of %2)...", 1, 2
        
    'Create a local array and point it at the pixel data of the current image
    Dim dstSA As SAFEARRAY2D
    EffectPrep.PrepImageData dstSA, toPreview, dstPic
    
    'Create a second local array.  This will contain the a copy of the current image, and we will use it as our source reference
    ' (This is necessary to prevent blurred pixel values from spreading across the image as we go.)
    Dim srcDIB As pdDIB
    Set srcDIB = New pdDIB
    srcDIB.CreateFromExistingDIB workingDIB
            
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left
    initY = curDIBValues.Top
    finalX = curDIBValues.Right
    finalY = curDIBValues.Bottom
    
    'If the quality is set to 1 ("better" quality), and the radius is under 30, simply use quality 0.  There is no reason
    ' to distinguish between them at that level, as differences really aren't noticeable until much larger amounts.
    'If (gaussQuality = 1) And (umRadius < 30) Then gaussQuality = 0
    
    'If this is a preview, we need to adjust the kernel radius to match the size of the preview box
    If toPreview Then
        umRadius = umRadius * curDIBValues.previewModifier
        If (umRadius = 0#) Then umRadius = 0.1
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
            gaussBlurSuccess = CreateApproximateGaussianBlurDIB(umRadius, workingDIB, srcDIB, 3, toPreview, progBarCalculation + finalX)
        
        'IIR Gaussian estimation
        Case Else
            progBarCalculation = finalY + finalX
            gaussBlurSuccess = Filters_Area.GaussianBlur_IIRImplementation(srcDIB, umRadius, 3, toPreview, progBarCalculation + finalX)
        
    End Select
    
    'Assuming the blur was created successfully, proceed with the masking portion of the filter.
    If (gaussBlurSuccess <> 0) Then
    
        'Now that we have a gaussian DIB created in workingDIB, we can point arrays toward it and the source DIB
        Dim dstImageData() As Byte
        CopyMemory ByVal VarPtrArray(dstImageData()), VarPtr(dstSA), 4
        
        Dim srcImageData() As Byte
        Dim srcSA As SAFEARRAY2D
        PrepSafeArray srcSA, srcDIB
        CopyMemory ByVal VarPtrArray(srcImageData()), VarPtr(srcSA), 4
        
        'These values will help us access locations in the array more quickly.
        ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
        Dim quickVal As Long, qvDepth As Long
        qvDepth = curDIBValues.BytesPerPixel
        
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
        
        umThreshold = umThreshold \ 5
        
        'The final step of the smart blur function is to find edges, and replace them with the blurred data as necessary
        For x = initX To finalX
            quickVal = x * qvDepth
        For y = initY To finalY
            
            'Retrieve the original image's pixels
            r = dstImageData(quickVal + 2, y)
            g = dstImageData(quickVal + 1, y)
            b = dstImageData(quickVal, y)
            
            'Now, retrieve the gaussian pixels
            r2 = srcImageData(quickVal + 2, y)
            g2 = srcImageData(quickVal + 1, y)
            b2 = srcImageData(quickVal, y)
            
            tLumDelta = Abs(GetLuminance(r, g, b) - GetLuminance(r2, g2, b2))
                            
            'If the delta is below the specified threshold, sharpen it
            If tLumDelta > umThreshold Then
                            
                newR = (scaleFactor * r) + (invScaleFactor * r2)
                If (newR > 255) Then newR = 255
                If (newR < 0) Then newR = 0
                    
                newG = (scaleFactor * g) + (invScaleFactor * g2)
                If (newG > 255) Then newG = 255
                If (newG < 0) Then newG = 0
                    
                newB = (scaleFactor * b) + (invScaleFactor * b2)
                If (newB > 255) Then newB = 255
                If (newB < 0) Then newB = 0
                
                blendVal = tLumDelta / 255
                
                newR = BlendColors(newR, r, blendVal)
                newG = BlendColors(newG, g, blendVal)
                newB = BlendColors(newB, b, blendVal)
                
                dstImageData(quickVal + 2, y) = newR
                dstImageData(quickVal + 1, y) = newG
                dstImageData(quickVal, y) = newB
                
                If qvDepth = 4 Then
                    a2 = srcImageData(quickVal + 3, y)
                    a = dstImageData(quickVal + 3, y)
                    newA = (scaleFactor * a) + (invScaleFactor * a2)
                    If newA > 255 Then newA = 255
                    If newA < 0 Then newA = 0
                    dstImageData(quickVal + 3, y) = BlendColors(newA, a, blendVal)
                End If
                
            End If
                    
        Next y
            If (Not toPreview) Then
                If (x And progBarCheck) = 0 Then
                    If Interface.UserPressedESC() Then Exit For
                    SetProgBarVal progBarCalculation + x
                End If
            End If
        Next x
        
        CopyMemory ByVal VarPtrArray(srcImageData), 0&, 4
        CopyMemory ByVal VarPtrArray(dstImageData), 0&, 4
        
        Set srcDIB = Nothing
        
    End If
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    EffectPrep.FinalizeImageData toPreview, dstPic
        
End Sub

Private Sub btsQuality_Click(ByVal buttonIndex As Long)
    UpdatePreview
End Sub

Private Sub cmdBar_OKClick()
    Process "Unsharp mask", , BuildParams(sltRadius, sltAmount, sltThreshold, btsQuality.ListIndex), UNDO_LAYER
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub Form_Load()
    
    'Suspend previews until the dialog is fully loaded
    cmdBar.MarkPreviewStatus False
    
    'Populate the quality selector
    btsQuality.AddItem "fast", 0
    btsQuality.AddItem "precise", 1
    btsQuality.ListIndex = 0
    
    'Apply visual themes to the form
    ApplyThemeAndTranslations Me
    cmdBar.MarkPreviewStatus True
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
    If cmdBar.PreviewsAllowed Then UnsharpMask sltRadius, sltAmount, sltThreshold, btsQuality.ListIndex, True, pdFxPreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub pdFxPreview_ViewportChanged()
    UpdatePreview
End Sub

Private Function GetLocalParamString() As String
    
    Dim cParams As pdParamXML
    Set cParams = New pdParamXML
    
    With cParams
    
    End With
    
    GetLocalParamString = cParams.GetParamString()
    
End Function
