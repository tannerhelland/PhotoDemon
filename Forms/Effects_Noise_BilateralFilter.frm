VERSION 5.00
Begin VB.Form FormBilateral 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Bilateral smoothing"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   12090
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
   ScaleWidth      =   806
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.pdCommandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5790
      Width           =   12090
      _ExtentX        =   21325
      _ExtentY        =   1323
   End
   Begin PhotoDemon.pdSlider sltRadius 
      Height          =   705
      Left            =   6000
      TabIndex        =   2
      Top             =   720
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
      Caption         =   "radius"
      Min             =   3
      Max             =   25
      Value           =   3
      DefaultValue    =   3
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
   Begin PhotoDemon.pdSlider sltSpatialFactor 
      Height          =   705
      Left            =   6000
      TabIndex        =   3
      Top             =   1680
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
      Caption         =   "edge strength"
      Min             =   1
      Max             =   100
      SigDigits       =   1
      Value           =   10
      DefaultValue    =   10
   End
   Begin PhotoDemon.pdSlider sltSpatialPower 
      Height          =   705
      Left            =   6000
      TabIndex        =   4
      Top             =   5250
      Visible         =   0   'False
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
      Caption         =   "spatial power (currently hidden)"
      Min             =   1
      SigDigits       =   2
      Value           =   2
      DefaultValue    =   2
   End
   Begin PhotoDemon.pdSlider sltColorFactor 
      Height          =   705
      Left            =   6000
      TabIndex        =   5
      Top             =   2640
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
      Caption         =   "color strength"
      Min             =   1
      Max             =   100
      SigDigits       =   1
      Value           =   10
      DefaultValue    =   10
   End
   Begin PhotoDemon.pdSlider sltColorPower 
      Height          =   705
      Left            =   6000
      TabIndex        =   6
      Top             =   3600
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
      Caption         =   "color preservation"
      Min             =   1
      SigDigits       =   2
      Value           =   2
      DefaultValue    =   2
   End
   Begin PhotoDemon.pdCheckBox chkSeparable 
      Height          =   330
      Left            =   6000
      TabIndex        =   7
      Top             =   4560
      Width           =   5820
      _ExtentX        =   10266
      _ExtentY        =   582
      Caption         =   "use estimation to improve performance"
   End
End
Attribute VB_Name = "FormBilateral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Bilateral Smoothing Form
'Copyright 2014 by Audioglider
'Created: 19/June/14
'Last updated: 23/July/14
'Last update: add a quasi-separable implementation that's about 20x faster than the naive one, at a minimal cost
'              to quality (in the Y-dimension; X-dimension quality should be roughly identical to the naive result).
'
'This filter performs selective gaussian smoothing of continuous areas of same color (domains), which removes noise
' and contrast artifacts while perserving sharp edges.
'
'The two major parameters "spatial factor" and "color factor" define the primary results of the filter. By modifying
' these parameters, users can achieve anything from light noise reduction with little change to the overall image,
' to a silky smooth cartoon-like effect across wide swaths of the image.
'
'More details on the algorithm can be found at:
' http://www.cs.duke.edu/~tomasi/papers/tomasi/tomasiIccv98.pdf
'
'In July '14, a quasi-separable variant of the function was added.  I call it "quasi-separable" because we use some
' modifications to compensate for bilateral smoothing not actually being mathematically separable.  (The spatial
' domain parameter is, but the color one is not.)  This provides a huge performance boost at a slight quality
' trade-off, so I've left the original implementation available via toggle.
'
'For details on the separable approach, see:
' http://homepage.tudelft.nl/e3q6n/publications/2005/ICME2005_TPLV.pdf
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

Private Const COLOR_COUNT As Long = 256
Private m_spatialFunc() As Double
Private m_colorFunc() As Double

'To improve performance, we only re-calculate kernel lookup tables when their relevant values change
Private m_LastKernelSize As Long, m_LastSpatialFactor As Double, m_LastSpatialPower As Double
Private m_LastColorFactor As Double, m_LastColorPower As Double

Private Sub InitSpatialFunc(ByVal kernelSize As Long, ByVal spatialFactor As Double, ByVal spatialPower As Double)
    
    If (kernelSize <> m_LastKernelSize) Or (spatialFactor <> m_LastSpatialFactor) Or (spatialPower <> m_LastSpatialPower) Then
    
        Dim i As Long, k As Long
        
        ReDim spatialFunc(-kernelSize To kernelSize, -kernelSize To kernelSize) As Double
        
        For i = -kernelSize To kernelSize
            For k = -kernelSize To kernelSize
                m_spatialFunc(i, k) = Exp(-0.5 * (Sqr(i * i + k * k) / spatialFactor) ^ spatialPower)
            Next k
        Next i
    
        m_LastKernelSize = kernelSize
        m_LastSpatialFactor = spatialFactor
        m_LastSpatialPower = spatialPower
    
    End If
    
End Sub

Private Sub InitColorFunc(ByVal colorFactor As Double, ByVal colorPower As Double)
    
    If (colorFactor <> m_LastColorFactor) Or (colorPower <> m_LastColorPower) Then
        
        Dim i As Long, k As Long
        
        ReDim m_colorFunc(0 To COLOR_COUNT - 1, 0 To COLOR_COUNT - 1) As Double
        
        For i = 0 To COLOR_COUNT - 1
            For k = 0 To COLOR_COUNT - 1
                m_colorFunc(i, k) = Exp(-0.5 * ((Abs(i - k) / colorFactor) ^ colorPower))
            Next k
        Next i
    
        m_LastColorFactor = colorFactor
        m_LastColorPower = colorPower
        
    End If
    
End Sub

'Parameters: * kernelRadius [size of square for limiting surrounding pixels that take part in calculation.
' NOTE: Small values < 9 on high-res images do not provide significant results.]
' * spatialFactor [determines smoothing power within a color domain (neighborhood pixels of similar color]
' * spatialPower [exponent power, used in spatial function calculation]
' * colorFactor [determines the variance of color for a color domain]
' * colorPower [exponent power, used in color function calculation]
Public Sub BilateralSmoothing(ByVal effectParams As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)
    
    'Extract function parameters
    Dim kernelRadius As Long
    Dim spatialFactor As Double, spatialPower As Double, colorFactor As Double, colorPower As Double
    
    Dim cParams As pdParamXML
    Set cParams = New pdParamXML
    cParams.SetParamString effectParams
    
    With cParams
        kernelRadius = .GetLong("Bilateral_KernelRadius", 1)
        spatialFactor = .GetDouble("Bilateral_SpatialFactor", 10#)
        spatialPower = .GetDouble("Bilateral_SpatialPower", 2#)
        colorFactor = .GetDouble("Bilateral_ColorFactor", 10#)
        colorPower = .GetDouble("Bilateral_ColorPower", 2#)
    End With
    
    'As a convenience to the user, we display spatial and color factors with a [0, 100].  The color factor can
    ' actually be bumped a bit, to [0, 255], so apply that now.
    colorFactor = colorFactor * 2.55
    
    'Spatial factor is left on a [0, 100] scale as a convenience to the user, but any value larger than about 10
    ' tends to produce meaningless results.  As such, shrink the input by a factor of 10.
    spatialFactor = spatialFactor / 10
    If (spatialFactor < 1#) Then spatialFactor = 1#
    
    'Spatial power is currently hidden from the user.  As such, default it to value 2.
    spatialPower = 2#
    
    If (Not toPreview) Then Message "Applying bilateral smoothing..."
    
    'Create a local array and point it at the pixel data of the current image
    Dim dstImageData() As Byte
    Dim dstSA As SAFEARRAY2D
    PrepImageData dstSA, toPreview, dstPic
    CopyMemory ByVal VarPtrArray(dstImageData()), VarPtr(dstSA), 4
    
    'Create a second local array. This will contain the a copy of the current image, and we will use it as our source reference
    ' (This is necessary to prevent already-processed pixels from affecting the results of later pixels.)
    Dim srcImageData() As Byte
    Dim srcSA As SAFEARRAY2D
    
    'If this is a preview, we need to adjust the kernal
    If toPreview Then kernelRadius = kernelRadius * curDIBValues.previewModifier
    If (kernelRadius < 1) Then kernelRadius = 1
    
    'To simplify the edge-handling required by this function, we're actually going to resize the source DIB with
    ' clamped pixel edges.  This removes the need for any edge handling whatsoever.
    Dim srcDIB As pdDIB
    Set srcDIB = New pdDIB
    PadDIBClampedPixels kernelRadius, kernelRadius, workingDIB, srcDIB
    
    PrepSafeArray srcSA, srcDIB
    CopyMemory ByVal VarPtrArray(srcImageData()), VarPtr(srcSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left
    initY = curDIBValues.Top
    finalX = curDIBValues.Right
    finalY = curDIBValues.Bottom
    
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim quickValDst As Long, quickValSrc As Long, QuickYSrc As Long, qvDepth As Long
    qvDepth = curDIBValues.BytesPerPixel
    
    'To keep processing quick, only update the progress bar when absolutely necessary. This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = FindBestProgBarValue()
        
    'Color variables
    Dim srcR As Long, srcG As Long, srcB As Long
    Dim newR As Long, newG As Long, newB As Long
    Dim srcR0 As Long, srcG0 As Long, srcB0 As Long
    
    Dim sCoefR As Double, sCoefG As Double, sCoefB As Double
    Dim sMembR As Double, sMembG As Double, sMembB As Double
    Dim coefR As Double, coefG As Double, coefB As Double
    Dim xOffset As Long, yOffset As Long, xMax As Long, yMax As Long, xMin As Long, yMin As Long
    Dim spacialFuncCache As Double
    Dim srcPixelX As Long, srcPixelY As Long
    
    'For performance improvements, color and spatial functions are precalculated prior to starting filter.
    InitSpatialFunc kernelRadius, spatialFactor, spatialPower
    InitColorFunc colorFactor, colorPower
    
    'Loop through each pixel in the image, converting values as we go
    For x = initX To finalX
        quickValDst = x * qvDepth
        quickValSrc = (x + kernelRadius) * qvDepth
    For y = initY To finalY
    
        sCoefR = 0
        sCoefG = 0
        sCoefB = 0
        sMembR = 0
        sMembG = 0
        sMembB = 0
        
        QuickYSrc = y + kernelRadius
        
        srcB0 = srcImageData(quickValSrc, QuickYSrc)
        srcG0 = srcImageData(quickValSrc + 1, QuickYSrc)
        srcR0 = srcImageData(quickValSrc + 2, QuickYSrc)
        
        'Cache y-loop boundaries so that they do not have to be re-calculated on the interior loop.  (X boundaries
        ' don't matter, but since we're doing it, for y, mirror it to x.)
        xMax = x + kernelRadius
        yMax = y + kernelRadius
        xMin = x - kernelRadius
        yMin = y - kernelRadius
        
        For xOffset = xMin To xMax
            For yOffset = yMin To yMax
                
                'Cache the source pixel's x and y locations
                srcPixelX = (xOffset + kernelRadius) * qvDepth
                srcPixelY = (yOffset + kernelRadius)
                
                srcB = srcImageData(srcPixelX, srcPixelY)
                srcG = srcImageData(srcPixelX + 1, srcPixelY)
                srcR = srcImageData(srcPixelX + 2, srcPixelY)
                
                spacialFuncCache = m_spatialFunc(xOffset - x, yOffset - y)
                
                'As a general rule, when convolving data against a kernel, any kernel value below 3-sigma can effectively
                ' be ignored (as its contribution to the convolution total is not statistically meaningful). Rather than
                ' calculating an actual sigma against a standard deviation for this kernel, we can approximate a threshold
                ' because we know that our source data - RGB colors - will only ever be on a [0, 255] range.  As such,
                ' let's assume that any spatial value below 1 / 255 (roughly 0.0039) is unlikely to have a meaningful
                ' impact on the final image; by simply ignoring values below that limit, we can save ourselves additional
                ' processing time when the incoming spatial parameters are low (as is common for the cartoon-like effect).
                If (spacialFuncCache > 0.0039) Then
                    
                    coefR = spacialFuncCache * m_colorFunc(srcR, srcR0)
                    coefG = spacialFuncCache * m_colorFunc(srcG, srcG0)
                    coefB = spacialFuncCache * m_colorFunc(srcB, srcB0)
                    
                    'We could perform an additional 3-sigma check here to account for meaningless colorFunc values, but
                    ' because we'd have to perform the check for each of R, G, and B, we risk inadvertently increasing
                    ' processing time when the color modifiers are consistently high.  As such, I think it's best to
                    ' limit our check to just the spatial modifier at present.
                    sCoefR = sCoefR + coefR
                    sCoefG = sCoefG + coefG
                    sCoefB = sCoefB + coefB
                    
                    sMembR = sMembR + coefR * srcR
                    sMembG = sMembG + coefG * srcG
                    sMembB = sMembB + coefB * srcB
                    
                End If
                        
            Next yOffset
        Next xOffset
        
        If (sCoefR <> 0) Then newR = sMembR / sCoefR
        If (sCoefG <> 0) Then newG = sMembG / sCoefG
        If (sCoefB <> 0) Then newB = sMembB / sCoefB
                
        'Assign the new values to each color channel
        dstImageData(quickValDst, y) = newB
        dstImageData(quickValDst + 1, y) = newG
        dstImageData(quickValDst + 2, y) = newR
        
    Next y
        If (Not toPreview) Then
            If (x And progBarCheck) = 0 Then
                If UserPressedESC() Then Exit For
                SetProgBarVal x
            End If
        End If
    Next x
    
    'With our work complete, point both ImageData() arrays away from their DIBs and deallocate them
    CopyMemory ByVal VarPtrArray(srcImageData), 0&, 4
    Erase srcImageData
    
    CopyMemory ByVal VarPtrArray(dstImageData), 0&, 4
    Erase dstImageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    FinalizeImageData toPreview, dstPic
    
End Sub

'Approximately the same function as BilateralSmoothing, above, but using a separable implementation to hugely boost performance.
' There is a quality trade-off, as the spatial parameter is separable but the color one is not, but we use some tricks to
' mitigate this.  All told, the separable function roughly adheres to the expected PxQ / (P+Q) performance boost, and my own
' testing shows a 10 megapixel photo at radius 25 to take just 5% of the time that a naive convolution does
' (naive: 302 seconds, separable: 14 seconds).
Public Sub BilateralSmoothingSeparable(ByVal effectParams As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)
    
    If (Not toPreview) Then Message "Applying bilateral smoothing..."
    
    'Extract function parameters
    Dim kernelRadius As Long
    Dim spatialFactor As Double, spatialPower As Double, colorFactor As Double, colorPower As Double
    
    Dim cParams As pdParamXML
    Set cParams = New pdParamXML
    cParams.SetParamString effectParams
    
    With cParams
        kernelRadius = .GetLong("Bilateral_KernelRadius", 1)
        spatialFactor = .GetDouble("Bilateral_SpatialFactor", 10#)
        spatialPower = .GetDouble("Bilateral_SpatialPower", 2#)
        colorFactor = .GetDouble("Bilateral_ColorFactor", 10#)
        colorPower = .GetDouble("Bilateral_ColorPower", 2#)
    End With
    
    'PrepImageData generates a working copy of the current filter target
    Dim dstSA As SAFEARRAY2D
    PrepImageData dstSA, toPreview, dstPic
    
    'If this is a preview, we need to adjust kernel size to match
    If toPreview Then kernelRadius = kernelRadius * curDIBValues.previewModifier
    
    'The kernel must be at least 1.0 pixels in either direction; otherwise, we'll get range errors
    If (kernelRadius < 1) Then kernelRadius = 1
    
    'PD now provides a dedicated function for separable bilateral processing
    CreateBilateralDIB workingDIB, kernelRadius, spatialFactor, spatialPower, colorFactor, colorPower, toPreview
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering using the data inside workingDIB
    FinalizeImageData toPreview, dstPic
    
End Sub

Public Sub BilateralWrapper(ByVal effectParams As String)
    
    Dim cParams As pdParamXML
    Set cParams = New pdParamXML
    cParams.SetParamString effectParams
    
    If cParams.GetBool("Bilateral_UseSeparable", True) Then
        Me.BilateralSmoothingSeparable effectParams
    Else
        Me.BilateralSmoothing effectParams
    End If
    
End Sub

Private Sub chkSeparable_Click()
    UpdatePreview
End Sub

Private Sub cmdBar_OKClick()
    Process "Bilateral smoothing", , GetLocalParamString(), UNDO_LAYER
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub Form_Load()
    
    cmdBar.MarkPreviewStatus False
    chkSeparable.AssignTooltip "Bilateral filtering is a complex task, and on large images it can take a very long time to process.  PhotoDemon can estimate certain parameters, providing a large speed boost at the cost of slightly lower quality."
    
    ApplyThemeAndTranslations Me
    cmdBar.MarkPreviewStatus True
    UpdatePreview
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub sltColorFactor_Change()
    UpdatePreview
End Sub

Private Sub sltColorPower_Change()
    UpdatePreview
End Sub

Private Sub sltRadius_Change()
    UpdatePreview
End Sub

Private Sub sltSpatialPower_Change()
    UpdatePreview
End Sub

Private Sub sltSpatialFactor_Change()
    UpdatePreview
End Sub

Private Sub UpdatePreview()

    If cmdBar.PreviewsAllowed Then
        
        'Separable kernels require a totally different function
        If CBool(chkSeparable) Then
            BilateralSmoothingSeparable GetLocalParamString(), True, pdFxPreview
        Else
            BilateralSmoothing GetLocalParamString(), True, pdFxPreview
        End If
        
    End If
    
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub pdFxPreview_ViewportChanged()
    UpdatePreview
End Sub

Private Function GetLocalParamString() As String
    
    Dim cParams As pdParamXML
    Set cParams = New pdParamXML
    
    With cParams
        .AddParam "Bilateral_KernelRadius", sltRadius.Value
        .AddParam "Bilateral_SpatialFactor", sltSpatialFactor.Value
        .AddParam "Bilateral_SpatialPower", sltSpatialPower.Value
        .AddParam "Bilateral_ColorFactor", sltColorFactor.Value
        .AddParam "Bilateral_ColorPower", sltColorPower.Value
        .AddParam "Bilateral_UseSeparable", CBool(chkSeparable)
    End With
    
    GetLocalParamString = cParams.GetParamString()
    
End Function
