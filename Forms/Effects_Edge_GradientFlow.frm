VERSION 5.00
Begin VB.Form FormGradientFlow 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   " Gradient flow"
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
   Begin PhotoDemon.pdSlider sldOpacityBack 
      Height          =   450
      Left            =   9240
      TabIndex        =   7
      Top             =   5220
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   794
      Max             =   100
      Value           =   100
      NotchPosition   =   2
      NotchValueCustom=   100
   End
   Begin PhotoDemon.pdColorSelector csBackground 
      Height          =   450
      Left            =   6360
      TabIndex        =   6
      Top             =   5220
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   794
   End
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   300
      Index           =   0
      Left            =   6240
      Top             =   4830
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   529
      Caption         =   "background color and opacity"
      FontSize        =   12
   End
   Begin PhotoDemon.pdButtonStrip btsTarget 
      Height          =   975
      Left            =   6000
      TabIndex        =   4
      Top             =   1800
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1931
      Caption         =   "render"
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
   Begin PhotoDemon.pdFxPreviewCtl pdFxPreview 
      Height          =   5625
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9922
   End
   Begin PhotoDemon.pdSlider sldSmoothing 
      Height          =   705
      Left            =   6000
      TabIndex        =   2
      Top             =   120
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
      Caption         =   "smoothing"
      Max             =   100
      ScaleStyle      =   1
   End
   Begin PhotoDemon.pdSlider sldBoost 
      Height          =   705
      Left            =   6000
      TabIndex        =   3
      Top             =   960
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
      Caption         =   "boost"
      Max             =   100
      ScaleStyle      =   1
   End
   Begin PhotoDemon.pdSlider sldSampleSize 
      Height          =   705
      Left            =   6240
      TabIndex        =   5
      Top             =   2880
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   1244
      Caption         =   "sample radius"
      Min             =   5
      Max             =   100
      ScaleStyle      =   1
      Value           =   5
      DefaultValue    =   5
   End
   Begin PhotoDemon.pdButtonStrip btsForeground 
      Height          =   975
      Left            =   6240
      TabIndex        =   8
      Top             =   3720
      Width           =   5655
      _ExtentX        =   10398
      _ExtentY        =   1931
      Caption         =   "foreground color"
   End
End
Attribute VB_Name = "FormGradientFlow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Gradient flow tool
'Copyright 2020-2026 by Tanner Helland
'Created: 26/September/20
'Last updated: 28/September/20
'Last update: wrap up initial build
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Calculate gradient flow for the underlying image, with some tweaks to improve visualization
Public Sub ApplyGradientFlowFx(ByVal parameterList As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)
    
    'Parse out the parameter list
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString parameterList
    
    Dim smoothingRadius As Long, boostMultiplier As Single, renderMagnitude As Boolean
    smoothingRadius = cParams.GetLong("smoothing", 0, True)
    boostMultiplier = (CSng(cParams.GetLong("boost", 0, True)) / 10!) + 1!
    renderMagnitude = Strings.StringsEqual(cParams.GetString("target", "target", True), "magnitude", True)
    
    Dim sampleRadius As Long, fxBackcolor As Long, backgroundOpacity As Long, dynamicForeground As Boolean
    sampleRadius = cParams.GetLong("sample-radius", 5, True)
    fxBackcolor = cParams.GetLong("background-color", vbWhite, True)
    backgroundOpacity = cParams.GetLong("background-opacity", 100, True)
    dynamicForeground = cParams.GetBool("dynamic-foreground", True, True)
    
    If (Not toPreview) Then Message "Calculating gradient flow..."
    
    'Generate workingDIB
    Dim dstSA As SafeArray2D
    EffectPrep.PrepImageData dstSA, toPreview, dstPic, doNotUnPremultiplyAlpha:=True
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left
    initY = curDIBValues.Top
    finalX = curDIBValues.Right
    finalY = curDIBValues.Bottom
    
    Dim gradLineWidth As Single
    gradLineWidth = 1.6
    
    'On non-previews, render a progress bar
    Dim progBarCheck As Long, progBarOffset As Long
    If toPreview Then
        smoothingRadius = smoothingRadius * curDIBValues.previewModifier
        sampleRadius = sampleRadius * curDIBValues.previewModifier
        gradLineWidth = 1!
    Else
        
        progBarOffset = 0
        progBarCheck = 0
        If (smoothingRadius > 0) Then
            progBarOffset = progBarOffset + (finalY + finalX) * 2
            progBarCheck = progBarCheck + progBarOffset
        End If
        
        If renderMagnitude Then
            progBarCheck = progBarCheck + finalY
        Else
            progBarCheck = progBarCheck + Int(workingDIB.GetDIBHeight \ sampleRadius) + 1
        End If
        
        ProgressBars.SetProgBarMax progBarCheck
        progBarCheck = ProgressBars.FindBestProgBarValue()
        
    End If
    
    'Enforce minimum sample radius
    If (sampleRadius < 3) Then sampleRadius = 3
    
    'Conditionally apply smoothing
    Dim tmpDIB As pdDIB
    If (smoothingRadius > 0) Then
        If (tmpDIB Is Nothing) Then Set tmpDIB = New pdDIB
        tmpDIB.CreateFromExistingDIB workingDIB
        Filters_Layers.CreateApproximateGaussianBlurDIB smoothingRadius, tmpDIB, workingDIB, 2, toPreview, ProgressBars.GetProgBarMax(), 0
    End If
    
    'Next, retrieve gradient and magnitude from the image
    Dim imgGrad() As Byte, imgMag() As Byte
    ReDim imgGrad(0 To workingDIB.GetDIBWidth - 1, 0 To workingDIB.GetDIBHeight - 1) As Byte
    ReDim imgMag(0 To workingDIB.GetDIBWidth - 1, 0 To workingDIB.GetDIBHeight - 1) As Byte
    Filters_Scientific.GetImageGradAndMag workingDIB, imgGrad, imgMag
    
    Dim r As Long, g As Long, b As Long, a As Long
    Dim newColor As Long
    
    'Build a lookup table of possible gray values
    Dim vLookup(0 To 255) As Long, tmpQuad As RGBQuad
    For x = 0 To 255
        newColor = Int(x * boostMultiplier + 0.5)
        If (newColor > 255) Then newColor = 255
        With tmpQuad
            .Alpha = 255
            .Red = newColor
            .Green = newColor
            .Blue = newColor
        End With
        GetMem4 VarPtr(tmpQuad), vLookup(x)
    Next x
    
    Dim dstImageData() As Long, tmpSA1D As SafeArray1D ', tmpGrad As Long
    
    'Rendering magnitude is much easier than direction
    If renderMagnitude Then
        
        For y = initY To finalY
            workingDIB.WrapLongArrayAroundScanline dstImageData, tmpSA1D, y
            For x = initX To finalX
                dstImageData(x) = vLookup(imgMag(x, y))
            Next x
            If (y And progBarCheck) = 0 Then
                If Interface.UserPressedESC() Then Exit For
                ProgressBars.SetProgBarVal progBarOffset + y
            End If
        Next y
        
        workingDIB.UnwrapLongArrayFromDIB dstImageData
        
    'Rendering angle is much more convoluted
    Else
        
        'Wrap a single 2D-array around the entire source image
        Dim srcPixels() As Byte
        workingDIB.WrapArrayAroundDIB srcPixels, dstSA
        If (tmpDIB Is Nothing) Then Set tmpDIB = New pdDIB
        tmpDIB.CreateFromExistingDIB workingDIB
        tmpDIB.FillWithColor fxBackcolor, backgroundOpacity
        If (backgroundOpacity <> 100) Then tmpDIB.SetAlphaPremultiplication True
        
        Dim forceForegroundColor As Long
        If (Not dynamicForeground) Then
            r = 255 - Colors.ExtractRed(fxBackcolor)
            g = 255 - Colors.ExtractGreen(fxBackcolor)
            b = 255 - Colors.ExtractBlue(fxBackcolor)
            forceForegroundColor = RGB(r, g, b)
        End If
        
        'We now need to add all vectors in each sample block of the image (according to the
        ' user-specified sampling radius).  This requires an inevitable cast back to
        ' cartesian space, alas.
        Dim xStride As Long
        
        'Calculate how many "tiles" we need to process in either direction
        Dim xLoop As Long, yLoop As Long
        xLoop = Int(workingDIB.GetDIBWidth \ sampleRadius) + 1
        yLoop = Int(workingDIB.GetDIBHeight \ sampleRadius) + 1
        
        'A number of other variables are required for the nested For...Next loops
        Dim dstXLoop As Long, dstYLoop As Long
        Dim initXLoop As Long, initYLoop As Long
        Dim i As Long, j As Long
        
        Const ANG_UN_NORMALIZE As Double = 6.28318530717959 / 255#
        Const MAG_UN_NORMALIZE = 1# / 255#
        Dim sinLookup(0 To 255) As Single, cosLookup(0 To 255) As Single
        Dim tmpMag As Single, tmpGrad As Single
        For x = 0 To 255
            tmpGrad = (CDbl(x) * ANG_UN_NORMALIZE) - PI
            sinLookup(x) = Sin(tmpGrad)
            cosLookup(x) = Cos(tmpGrad)
        Next x
        
        'We also need to count how many pixels must be averaged in each tile;
        ' (these are used to color directionality indicators).
        Dim numOfPixels As Long, pxDivisor As Double
        Dim xSum As Single, ySum As Single
        Dim pt1 As PointFloat, pt2 As PointFloat
        Dim sinAngle As Single, cosAngle As Single
        
        Dim cPen As pd2DPen
        Set cPen = New pd2DPen
        cPen.SetPenStartCap P2_LC_Flat
        cPen.SetPenEndCap P2_LC_ArrowAnchor
        cPen.SetPenWidth gradLineWidth
        cPen.SetPenOpacity 100!
        cPen.SetPenColor forceForegroundColor
        
        Dim cSurface As pd2DSurface
        Drawing2D.QuickCreateSurfaceFromDIB cSurface, tmpDIB, True
        cSurface.SetSurfacePixelOffset P2_PO_Half
        
        Dim maxRadius As Single, halfRadius As Single
        halfRadius = Sqr(sampleRadius) * boostMultiplier
        maxRadius = halfRadius * boostMultiplier
        
        Const GRAD_DIRECTION_ADJUST As Single = PI / 4!
        
        For y = 0 To yLoop
        For x = 0 To xLoop
            
            'Nested inner loops gather all data for the current mosaic tile
            initXLoop = x * sampleRadius
            initYLoop = y * sampleRadius
            dstXLoop = (x + 1) * sampleRadius - 1
            dstYLoop = (y + 1) * sampleRadius - 1
            
            For j = initYLoop To dstYLoop
            For i = initXLoop To dstXLoop
                
                'If this particular pixel is off of the image, don't bother counting it
                If (i < finalX) And (j < finalY) Then
                    
                    'Keep a running total of colors; we'll need to average these
                    xStride = i * 4
                    b = b + srcPixels(xStride, j)
                    g = g + srcPixels(xStride + 1, j)
                    r = r + srcPixels(xStride + 2, j)
                    a = a + srcPixels(xStride + 3, j)
                    
                    'We also need to tally horizontal and vertical components of this pixel's
                    ' gradient flow.  (That is the purpose of this tool, after all.)
                    
                    'Convert the original magnitude and angle values into separate x/y components
                    tmpMag = imgMag(i, j) * MAG_UN_NORMALIZE
                    xSum = xSum + tmpMag * cosLookup(imgGrad(i, j))
                    ySum = ySum + tmpMag * sinLookup(imgGrad(i, j))
                    
                    'Count this as a valid pixel
                    numOfPixels = numOfPixels + 1
                    
                End If
                
            Next i
            Next j
            
            'If this tile is completely off of the image, don't worry about it and go to the next one
            If (numOfPixels > 0) Then
                
                'Take the average red, green, and blue values of all the pixels within this tile
                pxDivisor = 1# / numOfPixels
                r = r * pxDivisor
                g = g * pxDivisor
                b = b * pxDivisor
                a = a * pxDivisor
                
                'Draw a line using the average color, with a length proportional to this block's
                ' magnitude, pointing in the direction of this pixel's net gradient.
                tmpMag = Sqr(xSum * xSum + ySum * ySum) * halfRadius
                If (tmpMag > maxRadius) Then tmpMag = maxRadius
                
                'Ignore regions with near-zero magnitude
                If (tmpMag > 0.1) Then
                    
                    'Calculate angle as the *perpendicular*, which is more intuitive for visualizing
                    tmpGrad = PDMath.Atan2(ySum, xSum) - GRAD_DIRECTION_ADJUST
                    If dynamicForeground Then cPen.SetPenColor RGB(r, g, b)
                    
                    'This calculation for a rotated line could be optimized further, but honestly
                    ' its perf is not a huge bother given the esoteric nature of this effect
                    xSum = (initXLoop + dstXLoop + 1) / 2
                    ySum = (initYLoop + dstYLoop + 1) / 2
                    sinAngle = Sin(tmpGrad)
                    cosAngle = Cos(tmpGrad)
                    pt1.x = xSum + (cosAngle * tmpMag - sinAngle * tmpMag)
                    pt1.y = ySum + (cosAngle * tmpMag + sinAngle * tmpMag)
                    pt2.x = xSum - (cosAngle * tmpMag - sinAngle * tmpMag)
                    pt2.y = ySum - (cosAngle * tmpMag + sinAngle * tmpMag)
                    
                    PD2D.DrawLineF cSurface, cPen, pt1.x, pt1.y, pt2.x, pt2.y
                    
                End If
                
            End If
    
            'Clear all trackers before processing the next block
            r = 0
            g = 0
            b = 0
            a = 0
            numOfPixels = 0
            
            xSum = 0!
            ySum = 0!
            
        Next x
            If (Not toPreview) Then
                If (y And progBarCheck) = 0 Then
                    If Interface.UserPressedESC() Then Exit For
                    ProgressBars.SetProgBarVal progBarOffset + y
                End If
            End If
        Next y
        
        workingDIB.UnwrapArrayFromDIB srcPixels
        workingDIB.CreateFromExistingDIB tmpDIB
    
    End If
    
    'Relinquish control over workingDIB
    EffectPrep.FinalizeImageData toPreview, dstPic, True
    
End Sub

Private Sub btsForeground_Click(ByVal buttonIndex As Long)
    UpdatePreview
End Sub

Private Sub btsTarget_Click(ByVal buttonIndex As Long)
    ReflowInterface
    UpdatePreview
End Sub

'OK button
Private Sub cmdBar_OKClick()
    Process "Gradient flow", , GetLocalParamString(), UNDO_Layer
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub csBackground_ColorChanged()
    UpdatePreview
End Sub

Private Sub Form_Load()

    'Disable previews while we initialize everything
    cmdBar.SetPreviewStatus False
    
    btsTarget.AddItem "magnitude", 0
    btsTarget.AddItem "direction", 1
    btsTarget.ListIndex = 0
    
    btsForeground.AddItem "dynamic", 0
    btsForeground.AddItem "static", 1
    btsForeground.ListIndex = 0
    
    'Apply translations and visual themes
    ApplyThemeAndTranslations Me, True, True
    cmdBar.SetPreviewStatus True
    
    ReflowInterface
    UpdatePreview
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub UpdatePreview()
    If cmdBar.PreviewsAllowed Then Me.ApplyGradientFlowFx GetLocalParamString(), True, pdFxPreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub pdFxPreview_ViewportChanged()
    UpdatePreview
End Sub

Private Sub sldBoost_Change()
    UpdatePreview
End Sub

Private Sub sldOpacityBack_Change()
    UpdatePreview
End Sub

Private Sub sldSampleSize_Change()
    UpdatePreview
End Sub

Private Sub sldSmoothing_Change()
    UpdatePreview
End Sub

Private Sub ReflowInterface()
    
    Dim gradModeActive As Boolean
    gradModeActive = (btsTarget.ListIndex = 1)
    
    sldSampleSize.Visible = gradModeActive
    lblTitle(0).Visible = gradModeActive
    csBackground.Visible = gradModeActive
    sldOpacityBack.Visible = gradModeActive
    btsForeground.Visible = gradModeActive

End Sub

Private Function GetLocalParamString() As String
    
    Dim cParams As New pdSerialize
    Set cParams = New pdSerialize
    
    With cParams
        
        .AddParam "smoothing", sldSmoothing.Value, True
        .AddParam "boost", sldBoost.Value, True
        
        If (btsTarget.ListIndex = 0) Then
            .AddParam "target", "magnitude"
        Else
            .AddParam "target", "direction"
        End If
        
        .AddParam "sample-radius", sldSampleSize.Value
        .AddParam "background-color", csBackground.Color
        .AddParam "background-opacity", sldOpacityBack.Value
        .AddParam "dynamic-foreground", (btsForeground.ListIndex = 0)
        
    End With
    
    GetLocalParamString = cParams.GetParamString()
    
End Function
