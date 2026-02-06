VERSION 5.00
Begin VB.Form FormBumpMap 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   " Bump map"
   ClientHeight    =   6540
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11925
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
   ScaleWidth      =   795
   Begin PhotoDemon.pdCommandBar cmdBar 
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5790
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   1323
   End
   Begin PhotoDemon.pdColorSelector csLight 
      Height          =   975
      Left            =   5880
      TabIndex        =   2
      Top             =   4080
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1720
      Caption         =   "color"
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
      PointSelection  =   -1  'True
   End
   Begin PhotoDemon.pdSlider sldIntensity 
      Height          =   705
      Left            =   8880
      TabIndex        =   3
      Top             =   1560
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   1244
      Caption         =   "intensity"
      Min             =   1
      Max             =   500
      SigDigits       =   1
      ScaleStyle      =   1
      Value           =   100
      NotchPosition   =   2
      NotchValueCustom=   100
   End
   Begin PhotoDemon.pdSlider sldDepth 
      Height          =   705
      Left            =   5880
      TabIndex        =   4
      Top             =   2400
      Width           =   2895
      _ExtentX        =   9551
      _ExtentY        =   1244
      Caption         =   "depth"
      Min             =   0.1
      SigDigits       =   2
      Value           =   2
      NotchPosition   =   2
      NotchValueCustom=   2
   End
   Begin PhotoDemon.pdSlider sldXCenter 
      Height          =   405
      Left            =   5880
      TabIndex        =   5
      Top             =   480
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   873
      Max             =   1
      SigDigits       =   2
      Value           =   0.5
      NotchPosition   =   2
      NotchValueCustom=   0.5
   End
   Begin PhotoDemon.pdSlider sldYCenter 
      Height          =   405
      Left            =   8880
      TabIndex        =   6
      Top             =   480
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   873
      Max             =   1
      SigDigits       =   2
      Value           =   0.5
      NotchPosition   =   2
      NotchValueCustom=   0.5
   End
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   330
      Index           =   0
      Left            =   5880
      Top             =   120
      Width           =   5685
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "center position (x, y)"
      FontSize        =   12
      ForeColor       =   4210752
   End
   Begin PhotoDemon.pdLabel lblExplanation 
      Height          =   435
      Index           =   0
      Left            =   6000
      Top             =   1050
      Width           =   5655
      _ExtentX        =   0
      _ExtentY        =   0
      Alignment       =   2
      Caption         =   "you can also set a center position by clicking the preview window"
      FontSize        =   9
      ForeColor       =   4210752
      Layout          =   1
   End
   Begin PhotoDemon.pdSlider sldRadius 
      Height          =   705
      Left            =   5880
      TabIndex        =   7
      Top             =   1560
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   1244
      Caption         =   "radius"
      Min             =   1
      Max             =   100
      Value           =   50
      NotchPosition   =   2
      NotchValueCustom=   50
   End
   Begin PhotoDemon.pdSlider sldAmbient 
      Height          =   705
      Left            =   8880
      TabIndex        =   8
      Top             =   2400
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   1244
      Caption         =   "ambient light"
      Max             =   100
   End
   Begin PhotoDemon.pdSlider sldOpacity 
      Height          =   720
      Left            =   5880
      TabIndex        =   9
      Top             =   3240
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   1270
      Caption         =   "opacity"
      CaptionPadding  =   2
      Max             =   100
      SigDigits       =   1
      Value           =   100
      NotchValueCustom=   25
      DefaultValue    =   100
   End
   Begin PhotoDemon.pdDropDown cboBlendMode 
      Height          =   735
      Left            =   8880
      TabIndex        =   10
      Top             =   3240
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   1296
      Caption         =   "blend mode"
   End
End
Attribute VB_Name = "FormBumpMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Bump Map Effect Dialog
'Copyright 2022-2026 by Tanner Helland
'Created: 18/April/22
'Last updated: 19/April/22
'Last update: wrap up initial build
'
'This tool was added in PhotoDemon v9.0.  It's only a rudimentary implementation of 2D bump-mapping,
' but it's (hopefully) good enough to satisfy the user that asked if I could implement something akin
' to FastStone Image Viewer's "bump map" effect.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Light map can be pre-calculated as it depends only on size of the "spotlight"
Private m_lightMap() As Single
Private m_lightWidth As Long, m_lightHeight As Long, m_lastIntensity As Double

'Height map is *not* a predetermined contour map; it's just (at present) a grayscale copy of the image.
' It's helpful to pre-cache it during preview stages since it doesn't change between renders.
Private m_heightMap() As Byte
Private m_hmWidth As Long, m_hmHeight As Long

Private Sub cboBlendMode_Click()
    UpdatePreview
End Sub

'OK button
Private Sub cmdBar_OKClick()
    Process "Bump map", , GetLocalParamString(), UNDO_Layer
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub cmdBar_ResetClick()
    sldAmbient.Value = 0
    sldDepth.Value = 2
    sldIntensity.Value = 100
    csLight.Color = RGB(255, 255, 255)
End Sub

Private Sub csLight_ColorChanged()
    UpdatePreview
End Sub

Private Sub Form_Load()
    
    cmdBar.SetPreviewStatus False
    
    'Populate the blend mode drop-down
    Interface.PopulateBlendModeDropDown cboBlendMode, BM_Normal
    
    'Apply visual themes and translations, then enable previews
    ApplyThemeAndTranslations Me, True, True
    cmdBar.SetPreviewStatus True
    UpdatePreview
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'Apply a 2D "bump map" effect to an image
Public Sub ApplyBumpMapEffect(ByVal effectParams As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)

    If (Not toPreview) Then Message "Generating bump map..."
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString effectParams
    
    Dim centerX As Double, centerY As Double
    Dim fxRadius As Double, fxAmbient As Double, fxIntensity As Double, fxDepth As Double
    Dim fxOpacity As Double, fxBlendMode As PD_BlendMode
    Dim fxColor As Long
    
    With cParams
        centerX = .GetDouble("centerx", 0.5)
        centerY = .GetDouble("centery", 0.5)
        fxAmbient = .GetDouble("ambient", 0#)
        fxRadius = .GetDouble("radius", 50#, True) / 100#
        fxIntensity = .GetDouble("intensity", 1#)
        fxDepth = .GetDouble("depth", sldDepth.Value)
        fxColor = .GetLong("color", csLight.Color)
        fxOpacity = .GetDouble("opacity", sldOpacity.Value)
        fxBlendMode = .GetLong("blendmode", cboBlendMode.ListIndex)
    End With
    
    'Create a local array and point it at the pixel data of the current image
    Dim imgPixels() As Byte, dstSA As SafeArray2D, dstSA1D As SafeArray1D
    EffectPrep.PrepImageData dstSA, toPreview, dstPic, doNotUnPremultiplyAlpha:=True
    
    'Because this effect looks artificially "bumpy" in preview mode, we want to suppress height map values
    ' during previews by a proportional amount.  This produces a better approximation of how the final effect
    ' is likely to look.
    If toPreview Then fxDepth = fxDepth * curDIBValues.previewModifier
    
    'At present, stride is always width * 4 (32-bit RGBA)
    Dim xStride As Long
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left
    initY = curDIBValues.Top
    finalX = curDIBValues.Right
    finalY = curDIBValues.Bottom
    
    'Radius is a function of the image's size.  Convert the incoming ratio to a fixed pixel size.
    Dim imgWidth As Long, imgHeight As Long
    imgWidth = curDIBValues.Width
    imgHeight = curDIBValues.Height
    
    Dim minDimension As Double
    If (imgWidth < imgHeight) Then minDimension = imgWidth Else minDimension = imgHeight
    fxRadius = fxRadius * minDimension
    
    'fxRadius is now a measurement in *pixels*.  Calculate a corresponding diameter.
    ' (This will be the size of our spotlight array.)
    Dim fxDiameter As Long
    fxDiameter = Int(fxRadius * 2 + 0.5)
    
    'Only rebuild the light map as necessary.  (If its size hasn't changed, we can reuse the map from
    ' our previous run - this results in much faster previews.)
    Dim lmChanged As Boolean: lmChanged = False
    
    If (m_lightWidth = 0) Then
        ReDim m_lightMap(0 To fxDiameter - 1, 0 To fxDiameter - 1) As Single
        m_lightWidth = fxDiameter
        m_lightHeight = fxDiameter
        lmChanged = True
    Else
        If (m_lightWidth <> fxDiameter) Or (m_lightHeight <> fxDiameter) Then
            ReDim m_lightMap(0 To fxDiameter - 1, 0 To fxDiameter - 1) As Single
            m_lightWidth = fxDiameter
            m_lightHeight = fxDiameter
            lmChanged = True
        Else
            lmChanged = False
        End If
    End If
    
    If lmChanged Or (m_lastIntensity <> fxIntensity) Then
        
        Dim invRadius As Double
        invRadius = 1! / fxRadius
        
        Dim nX As Single, nY As Single, nZ As Single
        
        'Construct the light-map (simple glowing circle for now)
        For y = 0 To fxDiameter - 1
            nY = (fxRadius - y) * invRadius
        For x = 0 To fxDiameter - 1
            nX = (fxRadius - x) * invRadius
            
            'Strength is inversely proportional to distance from center (i.e. the light is *strongest*
            ' at the center of the light, with squared drop-off as radius grows)
            nZ = 1! - Sqr(nX * nX + nY * nY)
            
            'Intensity can be applied as an artificial scaling of the precalculated value
            nZ = nZ * fxIntensity
            
            'Clip to zero and a little higher than 1 (to allow for some HDR-like "bloom"; we'll check for
            ' clipping on the inner loop, so this is OK)
            If (nZ < 0!) Then nZ = 0!
            If (nZ > 1.25!) Then nZ = 1.25!
            
            'Pre-cache all calculated values
            m_lightMap(x, y) = nZ
            
        Next x
        Next y
        
        m_lastIntensity = fxIntensity
        
    End If
    
    'Next, we want to cache a height-map of the image.  This is just a grayscale representation
    ' (where brightness corresponds to height).  It is helpful to cache this structure because it
    ' doesn't change unless the preview region changes.
    Dim hmChanged As Boolean: hmChanged = False
    
    If (m_hmWidth = 0) Then
        ReDim m_heightMap(0 To workingDIB.GetDIBWidth - 1, 0 To workingDIB.GetDIBHeight - 1) As Byte
        m_hmWidth = workingDIB.GetDIBWidth
        m_hmHeight = workingDIB.GetDIBHeight
        hmChanged = True
    Else
        If (m_hmWidth <> workingDIB.GetDIBWidth) Or (m_hmHeight <> workingDIB.GetDIBHeight) Then
            ReDim m_heightMap(0 To workingDIB.GetDIBWidth - 1, 0 To workingDIB.GetDIBHeight - 1) As Byte
            m_hmWidth = workingDIB.GetDIBWidth
            m_hmHeight = workingDIB.GetDIBHeight
            hmChanged = True
        Else
            hmChanged = False
        End If
    End If
    
    Dim r As Long, g As Long, b As Long
    If hmChanged Then
        
        For y = 0 To m_hmHeight - 1
            workingDIB.WrapArrayAroundScanline imgPixels, dstSA1D, y
        For x = 0 To m_hmWidth - 1
            
            'Retrieve source RGB values
            xStride = x * 4
            b = imgPixels(xStride)
            g = imgPixels(xStride + 1)
            r = imgPixels(xStride + 2)
            
            'Calculate a fast grayscale equivalent
            m_heightMap(x, y) = (218 * r + 732 * g + 74 * b) \ 1024
            
        Next x
        Next y
        
    End If
        
    'Height map is now complete
    workingDIB.UnwrapArrayFromDIB imgPixels
    
    'Create an effect container image; we will render the bump map into this DIB, then composite it
    ' atop the base image using the requested opacity+blendmode
    Dim fxDIB As pdDIB
    Set fxDIB = New pdDIB
    fxDIB.CreateBlank workingDIB.GetDIBWidth, workingDIB.GetDIBHeight, 32, 0, 255
    
    'Calculate the center of the light, modified by the user's custom center point inputs
    Dim midX As Double, midY As Double
    midX = CDbl(finalX - initX) * centerX + initX
    midY = CDbl(finalY - initY) * centerY + initY
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    If (Not toPreview) Then ProgressBars.SetProgBarMax finalY
    progBarCheck = ProgressBars.FindBestProgBarValue()
    
    'Color variables
    Dim rBase As Long, gBase As Long, bBase As Long
    
    'Extract the red, green, and blue values from the spotlight color.
    rBase = Colors.ExtractRed(fxColor)
    gBase = Colors.ExtractGreen(fxColor)
    bBase = Colors.ExtractBlue(fxColor)
    
    Dim srcIntensity As Double
    
    'Loop through each pixel in the image, converting values as we go
    For y = initY To finalY
        fxDIB.WrapArrayAroundScanline imgPixels, dstSA1D, y
    For x = initX To finalX
        
        'Calculate x/y gradients at this point
        Dim dx As Long, tX As Long, dy As Long, tY As Long
        If (x > initX) Then dx = m_heightMap(x - 1, y) Else dx = m_heightMap(x, y)
        If (x < finalX) Then tX = m_heightMap(x + 1, y) Else tX = m_heightMap(x, y)
        dx = dx - tX
        
        If (y > initY) Then dy = m_heightMap(x, y - 1) Else dy = m_heightMap(x, y)
        If (y < finalY) Then tY = m_heightMap(x, y + 1) Else tY = m_heightMap(x, y)
        dy = dy - tY
        
        'We now want to index this value into the pre-generated light map.  Note that we apply
        ' a few different modifiers to our calculation:
        ' 1) Obviously we need to scale against the difference between this point and the effect's center
        ' 2) Add the radius of the light to the offset to ensure the result from (1) is accurately scaled
        '    against the *center* of the lightmap (which sits at value fxRadius)
        ' 3) Artificially scale the gradient of this point against the user's passed depth value
        Dim lX As Long, lY As Long
        lX = (x - midX) + fxRadius + dx * fxDepth
        lY = (y - midY) + fxRadius + dy * fxDepth
        
        'Clip all values to the boundaries of the lightmap
        If (lX < 0) Then lX = 0
        If (lX >= fxDiameter) Then lX = fxDiameter - 1
        If (lY < 0) Then lY = 0
        If (lY >= fxDiameter) Then lY = fxDiameter - 1
        
        'Pull the corresponding light value from the table, then...
        srcIntensity = m_lightMap(lX, lY)
        
        '...add the ambient light value, if any
        srcIntensity = srcIntensity + fxAmbient
        
        'Calculate a final color as a ratio of the base color
        r = Int(rBase * srcIntensity + 0.5)
        g = Int(gBase * srcIntensity + 0.5)
        b = Int(bBase * srcIntensity + 0.5)
        
        'Because we allow intensity to exceed 1.0, we must perform a final clamp of RGB values.
        If (r < 0) Then r = 0
        If (r > 255) Then r = 255
        If (g < 0) Then g = 0
        If (g > 255) Then g = 255
        If (b < 0) Then b = 0
        If (b > 255) Then b = 255
        
        'Assign the final result and carry on!
        xStride = x * 4
        imgPixels(xStride) = b
        imgPixels(xStride + 1) = g
        imgPixels(xStride + 2) = r
        
        '(Note that we leave alpha as-is for this effect.)
        
    Next x
        If (Not toPreview) Then
            If (y And progBarCheck) = 0 Then
                If Interface.UserPressedESC() Then Exit For
                SetProgBarVal y
            End If
        End If
    Next y
    
    'Safely deallocate all image arrays
    fxDIB.UnwrapArrayFromDIB imgPixels
    fxDIB.SetInitialAlphaPremultiplicationState True
    
    'Merge down the result
    Dim cCompositor As pdCompositor
    Set cCompositor = New pdCompositor
    cCompositor.QuickMergeTwoDibsOfEqualSize workingDIB, fxDIB, fxBlendMode, fxOpacity
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    EffectPrep.FinalizeImageData toPreview, dstPic, True
 
End Sub

'Render a new preview
Private Sub UpdatePreview()
    If cmdBar.PreviewsAllowed Then ApplyBumpMapEffect GetLocalParamString(), True, pdFxPreview
End Sub

Private Sub pdFxPreview_PointSelected(xRatio As Double, yRatio As Double)
    cmdBar.SetPreviewStatus False
    sldXCenter.Value = xRatio
    sldYCenter.Value = yRatio
    cmdBar.SetPreviewStatus True
    UpdatePreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub pdFxPreview_ViewportChanged()
    UpdatePreview
End Sub

Private Sub sldAmbient_Change()
    UpdatePreview
End Sub

Private Sub sldDepth_Change()
    UpdatePreview
End Sub

Private Sub sldIntensity_Change()
    UpdatePreview
End Sub

Private Sub sldOpacity_Change()
    UpdatePreview
End Sub

Private Sub sldRadius_Change()
    UpdatePreview
End Sub

Private Sub sldXCenter_Change()
    UpdatePreview
End Sub

Private Sub sldYCenter_Change()
    UpdatePreview
End Sub

Private Function GetLocalParamString() As String
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    
    With cParams
        .AddParam "centerx", sldXCenter.Value
        .AddParam "centery", sldYCenter.Value
        .AddParam "ambient", sldAmbient.Value / 100#
        .AddParam "radius", sldRadius.Value
        .AddParam "intensity", sldIntensity.Value / 100#
        .AddParam "depth", sldDepth.Value
        .AddParam "color", csLight.Color
        .AddParam "opacity", sldOpacity.Value
        .AddParam "blendmode", cboBlendMode.ListIndex
    End With
    
    GetLocalParamString = cParams.GetParamString()
    
End Function
