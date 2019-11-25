VERSION 5.00
Begin VB.Form FormSurfaceBlur 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Surface blur"
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
   Begin PhotoDemon.pdSlider sltRadius 
      Height          =   705
      Left            =   6000
      TabIndex        =   3
      Top             =   1920
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
      Caption         =   "radius"
      Min             =   0.1
      Max             =   200
      SigDigits       =   1
      ScaleStyle      =   1
      Value           =   5
      DefaultValue    =   5
   End
   Begin PhotoDemon.pdSlider sltThreshold 
      Height          =   705
      Left            =   6000
      TabIndex        =   4
      Top             =   2880
      Width           =   5925
      _ExtentX        =   10451
      _ExtentY        =   1270
      Caption         =   "threshold"
      Max             =   255
      Value           =   50
      DefaultValue    =   50
   End
   Begin PhotoDemon.pdButtonStrip btsQuality 
      Height          =   1080
      Left            =   6000
      TabIndex        =   5
      Top             =   3720
      Width           =   5910
      _ExtentX        =   10425
      _ExtentY        =   1058
      Caption         =   "mode"
   End
   Begin PhotoDemon.pdButtonStrip btsArea 
      Height          =   1080
      Left            =   6000
      TabIndex        =   2
      Top             =   600
      Width           =   5910
      _ExtentX        =   10425
      _ExtentY        =   1058
      Caption         =   "target"
   End
End
Attribute VB_Name = "FormSurfaceBlur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Surface Blur Tool (formerly "Smart Blur")
'Copyright 2013-2019 by Tanner Helland
'Created: 17/January/13
'Last updated: 25/November/19
'Last update: performance improvements
'
'There are many different ways to implement a surface/selective blur tool.  At present, PD's operates as
' a selective gaussian, with the underlying gaussian data serving as the guideline for edges vs smooth
' regions.  (If a pixel from the source matches its blurred value, it's likely a smooth region.)
'
'This is very fast, but not the most accurate way to handle a surface blur.  A better version could
' use a separate edge detection map - and in fact, this may be a fix that I add in the near-future.
'
'Similarly, for *really* good edge handling, you could use the Noise > Bilateral Filter tool.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Convolve an image using a selective gaussian kernel (separable implementation!)
' Inputs:
'   - Radius of the blur (min 1, no real max - but processing speed obviously drops as the radius increases)
'   - Threshold (controls edge/surface distinction)
'   - Smooth Edges (for traditional surface blur (false) vs PD's edge-only softening (true))
'   - Blur quality (0, 1, 2 for iterative box blur, IIR blur, or true Gaussian, respectively)
Public Sub SurfaceBlurFilter(ByVal effectParams As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)
    
    If (Not toPreview) Then Message "Analyzing image in preparation for surface blur..."
    
    Dim cParams As pdParamXML
    Set cParams = New pdParamXML
    cParams.SetParamString effectParams
    
    Dim gRadius As Double, gThreshold As Long, smoothEdges As Boolean, sbQuality As Long
    
    With cParams
        gRadius = .GetDouble("radius", sltRadius.Value)
        gThreshold = .GetLong("threshold", sltThreshold.Value)
        smoothEdges = .GetBool("type", False)
        sbQuality = .GetLong("quality", btsQuality.ListIndex)
    End With
    
    'More color variables - in this case, sums for each color component
    Dim r As Long, g As Long, b As Long
    Dim r2 As Long, g2 As Long, b2 As Long
    Dim tDelta As Long
    
    'Create a local array and point it at the pixel data of the current image
    Dim dstSA As SafeArray2D
    EffectPrep.PrepImageData dstSA, toPreview, dstPic
    
    'Create a second local array.  This will contain the a copy of the current image, and we will use it as our source reference
    ' (This is necessary to prevent blurred pixel values from spreading across the image as we go.)
    Dim gaussDIB As pdDIB
    Set gaussDIB = New pdDIB
    gaussDIB.CreateFromExistingDIB workingDIB
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = (curDIBValues.Width - 1) * 4
    finalY = curDIBValues.Bottom
    
    'If this is a preview, we need to adjust the kernel radius to match the size of the preview box
    If toPreview Then
        gRadius = gRadius * curDIBValues.previewModifier
        If (gRadius = 0#) Then gRadius = 0.01
    End If
    
    'I almost always recommend quality over speed for PD tools, but in this case, the fast option is SO much faster,
    ' and the results so indistinguishable (3% different according to the Central Limit Theorem:
    ' https://www.khanacademy.org/math/probability/statistics-inferential/sampling_distribution/v/central-limit-theorem?playlist=Statistics
    ' ), that I recommend the faster methods instead.
    Dim gaussBlurSuccess As Long
    gaussBlurSuccess = 0
    
    Dim progBarCalculation As Long
    progBarCalculation = 0
    
    Select Case sbQuality
    
        '3 iteration box blur
        Case 0
            progBarCalculation = finalY * 3 + finalX * 3
            gaussBlurSuccess = CreateApproximateGaussianBlurDIB(gRadius, workingDIB, gaussDIB, 3, toPreview, progBarCalculation + finalX)
        
        'IIR Gaussian estimation
        Case Else
            progBarCalculation = finalY + finalX
            gaussBlurSuccess = Filters_Area.GaussianBlur_Deriche(gaussDIB, gRadius, 3, toPreview, progBarCalculation + finalX)
        
    End Select
    
    'Make sure our blur DIB created successfully before continuing
    If gaussBlurSuccess Then
        
        Dim srcDIB As pdDIB
        Set srcDIB = New pdDIB
        srcDIB.CreateFromExistingDIB workingDIB
        
        'Now that we have a gaussian DIB created in gaussDIB, we can point arrays toward it and the source DIB
        Dim dstImageData() As Byte, dstSA1D As SafeArray1D
        Dim srcImageData() As Byte, srcSA1D As SafeArray1D
        Dim gaussImageData() As Byte, gaussSA1D As SafeArray1D
        
        If (Not toPreview) Then
            Message "Applying surface blur..."
            ProgressBars.SetProgBarMax finalY
        End If
        
        'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
        ' based on the size of the area to be processed.
        Dim progBarCheck As Long
        progBarCheck = ProgressBars.FindBestProgBarValue()
            
        Dim blendVal As Double
        
        'The final step of the smart blur function is to find edges, and replace them with the blurred data as necessary
        For y = initY To finalY
            
            'Wrap arrays around the relevant scanline in each image (source, gauss, dest)
            workingDIB.WrapArrayAroundScanline dstImageData, dstSA1D, y
            srcDIB.WrapArrayAroundScanline srcImageData, srcSA1D, y
            gaussDIB.WrapArrayAroundScanline gaussImageData, gaussSA1D, y
            
        For x = initX To finalX Step 4
            
            'Retrieve the original image's pixels and
            b = srcImageData(x)
            g = srcImageData(x + 1)
            r = srcImageData(x + 2)
            
            tDelta = (218 * r + 732 * g + 74 * b) \ 1024
            
            'Now, retrieve the gaussian pixels
            b2 = gaussImageData(x)
            g2 = gaussImageData(x + 1)
            r2 = gaussImageData(x + 2)
            
            'Calculate a delta between the two
            tDelta = tDelta - ((218 * r2 + 732 * g2 + 74 * b2) \ 1024)
            If (tDelta < 0) Then tDelta = -tDelta
            
            'If the delta is below the specified threshold, replace it with the blurred data.
            If smoothEdges Then
            
                If (tDelta > gThreshold) Then
                    If (tDelta <> 0) Then blendVal = 1# - (gThreshold / tDelta) Else blendVal = 0#
                    dstImageData(x) = BlendColors(srcImageData(x), gaussImageData(x), blendVal)
                    dstImageData(x + 1) = BlendColors(srcImageData(x + 1), gaussImageData(x + 1), blendVal)
                    dstImageData(x + 2) = BlendColors(srcImageData(x + 2), gaussImageData(x + 2), blendVal)
                    dstImageData(x + 3) = BlendColors(srcImageData(x + 3), gaussImageData(x + 3), blendVal)
                End If
            
            Else
            
                If (tDelta <= gThreshold) Then
                    If (gThreshold <> 0) Then blendVal = 1# - (tDelta / gThreshold) Else blendVal = 1#
                    dstImageData(x) = BlendColors(srcImageData(x), gaussImageData(x), blendVal)
                    dstImageData(x + 1) = BlendColors(srcImageData(x + 1), gaussImageData(x + 1), blendVal)
                    dstImageData(x + 2) = BlendColors(srcImageData(x + 2), gaussImageData(x + 2), blendVal)
                    dstImageData(x + 3) = BlendColors(srcImageData(x + 3), gaussImageData(x + 3), blendVal)
                End If
        
            End If
            
        Next x
            If (Not toPreview) Then
                If (y And progBarCheck) = 0 Then
                    If Interface.UserPressedESC() Then Exit For
                    SetProgBarVal y + progBarCalculation
                End If
            End If
        Next y
            
        'With our work complete, release all arrays and temporary image copies
        gaussDIB.UnwrapArrayFromDIB gaussImageData
        srcDIB.UnwrapArrayFromDIB srcImageData
        workingDIB.UnwrapArrayFromDIB dstImageData
        
        Set gaussDIB = Nothing
        Set srcDIB = Nothing
        
    End If
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    EffectPrep.FinalizeImageData toPreview, dstPic
    
End Sub

Private Sub btsArea_Click(ByVal buttonIndex As Long)
    UpdatePreview
End Sub

Private Sub btsQuality_Click(ByVal buttonIndex As Long)
    UpdatePreview
End Sub

Private Sub cmdBar_OKClick()
    Process "Surface blur", , GetLocalParamString(), UNDO_Layer
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub cmdBar_ResetClick()
    sltThreshold.Value = 50
End Sub

Private Sub Form_Load()
    
    'Disable previews until the dialog is fully loaded
    cmdBar.SetPreviewStatus False
    
    'Apply button strip captions
    btsArea.AddItem "smooth areas", 0
    btsArea.AddItem "edges", 1
    btsArea.ListIndex = 0
    
    btsQuality.AddItem "fast", 0
    btsQuality.AddItem "precise", 1
    btsQuality.ListIndex = 0
    
    'Apply visual themes
    ApplyThemeAndTranslations Me
    
    cmdBar.SetPreviewStatus True
    UpdatePreview
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub sltRadius_Change()
    UpdatePreview
End Sub

Private Sub sltThreshold_Change()
    UpdatePreview
End Sub

'Render a new effect preview
Private Sub UpdatePreview()
    If cmdBar.PreviewsAllowed Then SurfaceBlurFilter GetLocalParamString(), True, pdFxPreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub pdFxPreview_ViewportChanged()
    UpdatePreview
End Sub

Private Function GetLocalParamString() As String
    
    Dim cParams As pdParamXML
    Set cParams = New pdParamXML
    
    With cParams
        .AddParam "radius", sltRadius.Value
        .AddParam "threshold", sltThreshold.Value
        .AddParam "type", (btsArea.ListIndex = 1)
        .AddParam "quality", btsQuality.ListIndex
    End With
    
    GetLocalParamString = cParams.GetParamString()
    
End Function
