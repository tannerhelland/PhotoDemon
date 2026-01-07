VERSION 5.00
Begin VB.Form FormSnow 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   " Snow"
   ClientHeight    =   6540
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12120
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
   ScaleWidth      =   808
   Begin PhotoDemon.pdSlider sldIntensity 
      Height          =   705
      Left            =   6000
      TabIndex        =   2
      Top             =   480
      Width           =   5880
      _ExtentX        =   10372
      _ExtentY        =   1270
      Caption         =   "intensity"
      Max             =   100
      SigDigits       =   1
      Value           =   10
      DefaultValue    =   10
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
      Width           =   12120
      _ExtentX        =   21378
      _ExtentY        =   1323
   End
   Begin PhotoDemon.pdSlider sldAngle 
      Height          =   705
      Left            =   6000
      TabIndex        =   3
      Top             =   2160
      Width           =   5880
      _ExtentX        =   10372
      _ExtentY        =   1270
      Caption         =   "angle"
      Min             =   -90
      Max             =   90
      SigDigits       =   1
   End
   Begin PhotoDemon.pdSlider sldWind 
      Height          =   705
      Left            =   6000
      TabIndex        =   4
      Top             =   3000
      Width           =   5880
      _ExtentX        =   10372
      _ExtentY        =   1270
      Caption         =   "wind"
      Max             =   100
      SigDigits       =   1
      Value           =   2
      DefaultValue    =   2
   End
   Begin PhotoDemon.pdSlider sldSize 
      Height          =   705
      Left            =   6000
      TabIndex        =   5
      Top             =   1320
      Width           =   5880
      _ExtentX        =   10372
      _ExtentY        =   1270
      Caption         =   "size"
      Max             =   100
      SigDigits       =   1
      Value           =   20
      DefaultValue    =   20
   End
   Begin PhotoDemon.pdSlider sldOpacity 
      Height          =   705
      Left            =   6000
      TabIndex        =   6
      Top             =   3840
      Width           =   5880
      _ExtentX        =   10372
      _ExtentY        =   1270
      Caption         =   "opacity"
      Max             =   100
      Value           =   100
      NotchPosition   =   2
      NotchValueCustom=   100
   End
   Begin PhotoDemon.pdCheckBox chkRandomize 
      Height          =   375
      Left            =   6000
      TabIndex        =   7
      Top             =   5040
      Width           =   5805
      _ExtentX        =   10239
      _ExtentY        =   661
      Caption         =   "randomize"
   End
End
Attribute VB_Name = "FormSnow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Nature > "Snow" Effect Dialog
'Copyright 2017-2026 by Tanner Helland
'Created: 03/August/17
'Last updated: 04/August/17
'Last update: wrap up initial build
'
'PD previously had a filter called "freeze", which kinda (not really) made an image look like it had
' been frozen.  The filter was unpredictable and only really worked on certain hues, but I liked the idea of
' a wintertime filter, so in 2017 I removed that low-quality function and replaced it with a "falling snow"
' generator.
'
'This function is basically just a cascading series of other, simpler functions.  Snow is generated as a
' series of randomized 3-point curved polygons.  Those polygons are then optionally blurred (contingent
' on their size), then motion-blurred to assign a "wind" direction and strength.  Finally, the entire
' result is composited against the base layer using Screen mode (which turns black transparent).  The end
' result is surprisingly good, especially if used over photos of winter scenes.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

Private m_Randomize As pdRandomize

'To improve preview performance, we reuse temporary DIBs
Private m_snowDIB As pdDIB

'Apply a hazy, cool color transformation I call an "atmospheric" transform.
Public Sub ApplySnowEffect(ByVal effectParams As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)

    If (Not toPreview) Then Message "Let it snow, let it snow, let it snow..."
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString effectParams
    
    Dim snowIntensity As Double, flakeSize As Double, snowAngle As Double, windStrength As Double
    Dim finalOpacity As Double
    
    With cParams
        snowIntensity = .GetDouble("intensity", 50#)
        flakeSize = .GetDouble("size", 5#)
        snowAngle = .GetDouble("angle", 0#) - 90#
        windStrength = .GetDouble("wind", 5#)
        finalOpacity = .GetDouble("opacity", 100#)
        m_Randomize.SetSeed_Float .GetDouble("seed", m_Randomize.GetRandomFloat_VB)
    End With
    
    'Generate a source image matching the current preview area
    Dim tmpSA As SafeArray2D
    EffectPrep.PrepImageData tmpSA, toPreview, dstPic, , , True
    
    Dim x As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left
    initY = curDIBValues.Top
    finalX = curDIBValues.Right
    finalY = curDIBValues.Bottom
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    If (Not toPreview) Then ProgressBars.SetProgBarMax 7
    progBarCheck = ProgressBars.FindBestProgBarValue()
    
    'We're going to render our snow into its own dedicated DIB, which will then be overlaid atop the
    ' original image as the final step.
    If (m_snowDIB Is Nothing) Then Set m_snowDIB = New pdDIB
    m_snowDIB.CreateBlank workingDIB.GetDIBWidth, workingDIB.GetDIBHeight, 32, 0, 255
    
    'Maximum flake size is calculated as an absolute value.  (Previous tests made this value relative,
    ' dependening on the size of the image, but it turns out that you never want very large flake sizes
    ' because they start looking terribly fake.)
    Dim maxRadius As Double
    maxRadius = flakeSize * 0.25 * curDIBValues.previewModifier
    
    'Similarly, motion blur size needs to be modified to reflect preview settings
    Dim motionBlurRadius As Double
    motionBlurRadius = windStrength * curDIBValues.previewModifier
    
    'Above a certain size, individual flakes look particularly hard-edged and unnatural.  As such, we need to
    ' manually blur the resulting image to make the flakes look "natural".  (On small flake sizes, we can just
    ' cheat and render the flakes with antialiasing.)
    Dim softenRadius As Double
    If (maxRadius > 5#) Then
        
        'Because we're using a gaussian blur, the blur radius has to be fairly high relative to flake size.
        ' (This also gives the illusion of bokeh.)
        softenRadius = maxRadius * 0.75
        
        'Above a certain point, blurring becomes redundant, especially if motion blur is applied
        ' on top of it.  As such, limit the blur past a certain point.
        If (softenRadius > 15#) Then softenRadius = 15#
        
    End If
    
    Dim blurNeeded As Boolean
    blurNeeded = (softenRadius > 0#)
    
    'If we're just gonna blur the flakes, we don't need to care about antialiasing.  If no blur is occurring,
    ' however, use antialiasing to improve output.
    Dim dstSurface As pd2DSurface
    Set dstSurface = New pd2DSurface
    dstSurface.WrapSurfaceAroundPDDIB m_snowDIB
    dstSurface.SetSurfacePixelOffset P2_PO_Half
    If (Not blurNeeded) Then dstSurface.SetSurfaceAntialiasing P2_AA_HighQuality Else dstSurface.SetSurfaceAntialiasing P2_AA_None
    
    'Start with a pure white brush at 100% opacity, but note that opacity will be randomly varied
    ' while the flakes are drawn.
    Dim cBrush As pd2DBrush
    Drawing2D.QuickCreateSolidBrush cBrush, vbWhite, 100#
    
    'The number of flakes we render corresponds linearly to the user's supplied value.
    ' (At maximum intensity, we render 40,000 unique flakes.)
    Dim numOfFlakes As Long
    numOfFlakes = snowIntensity * 400#
    
    Dim centerX As Double, centerY As Double
    Dim ptAngle As Single
    
    Dim shapeCorners() As PointFloat
    ReDim shapeCorners(0 To 3) As PointFloat
    
    If (Not toPreview) Then ProgressBars.SetProgBarVal 1
    
    'Generate each flake using a similar formula
    For x = 0 To numOfFlakes - 1
        
        'First, calculate a random center point
        centerX = m_Randomize.GetRandomFloat_WH() * finalX
        centerY = m_Randomize.GetRandomFloat_WH() * finalY
        
        'Next, generate three random points around the center, with a radius that varies up to maxRadius
        ptAngle = PI_DOUBLE * m_Randomize.GetRandomFloat_WH()
        PDMath.ConvertPolarToCartesian_Sng ptAngle, m_Randomize.GetRandomFloat_WH() * maxRadius, shapeCorners(0).x, shapeCorners(0).y, centerX, centerY
        
        ptAngle = PI_DOUBLE * m_Randomize.GetRandomFloat_WH()
        PDMath.ConvertPolarToCartesian_Sng ptAngle, m_Randomize.GetRandomFloat_WH() * maxRadius, shapeCorners(1).x, shapeCorners(1).y, centerX, centerY
        
        ptAngle = PI_DOUBLE * m_Randomize.GetRandomFloat_WH()
        PDMath.ConvertPolarToCartesian_Sng ptAngle, m_Randomize.GetRandomFloat_WH() * maxRadius, shapeCorners(2).x, shapeCorners(2).y, centerX, centerY
        
        'Randomize brush opacity between 33% and 100% to provide the illusion of depth.
        cBrush.SetBrushOpacity m_Randomize.GetRandomFloat_WH() * 67# + 33#
        
        'Draw our completed polygon, and use a "curvature" algorithm to give it a more natural shape.
        PD2D.FillPolygonF_FromPtF dstSurface, cBrush, 3, VarPtr(shapeCorners(0)), True
    
    Next x
    
    If (Not toPreview) Then ProgressBars.SetProgBarVal 2
    
    'If the flakes are very large, apply a conditional blur to improve their appearance
    Dim tmpDIB As pdDIB
    If blurNeeded Then
        Set tmpDIB = New pdDIB
        tmpDIB.CreateFromExistingDIB m_snowDIB
        Filters_Layers.CreateApproximateGaussianBlurDIB softenRadius, tmpDIB, m_snowDIB, 3, True
    End If
    
    If (Not toPreview) Then ProgressBars.SetProgBarVal 3
    
    'Motion-blur the DIB according to the user's wind settings
    If (motionBlurRadius > 0#) Then
        
        'Rotate the snow DIB into position
        Dim rotateDIB As pdDIB
        Set rotateDIB = New pdDIB
        GDI_Plus.GDIPlus_GetRotatedClampedDIB m_snowDIB, rotateDIB, snowAngle
        
        If (Not toPreview) Then ProgressBars.SetProgBarVal 4
        
        'Apply motion blur
        If (tmpDIB Is Nothing) Then Set tmpDIB = New pdDIB
        tmpDIB.CreateFromExistingDIB rotateDIB
        Dim blurSuccess As Boolean
        blurSuccess = CreateHorizontalBlurDIB(0, motionBlurRadius, tmpDIB, rotateDIB, True)
        Set tmpDIB = Nothing
        
        If (Not toPreview) Then ProgressBars.SetProgBarVal 5
        
        'Rotate the DIB back into its original position
        If blurSuccess Then GDI_Plus.GDIPlus_RotateDIBPlgStyle rotateDIB, m_snowDIB, -snowAngle, True
        Set rotateDIB = Nothing
        
        If (Not toPreview) Then ProgressBars.SetProgBarVal 6
        
    End If
    
    'Overlay the snow result onto the image, using the "Screen" blend mode to fade out black shades
    Dim cCompositor As pdCompositor
    Set cCompositor = New pdCompositor
    cCompositor.QuickMergeTwoDibsOfEqualSize workingDIB, m_snowDIB, BM_Screen, finalOpacity, AM_Normal, AM_Inherit
    
    If (Not toPreview) Then ProgressBars.SetProgBarVal 7
    
    'On non-previews, free our intermediate copy
    If (Not toPreview) Then Set m_snowDIB = Nothing
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    EffectPrep.FinalizeImageData toPreview, dstPic, True
    
End Sub

Private Sub chkRandomize_Click()
    UpdatePreview
End Sub

Private Sub cmdBar_OKClick()
    Process "Snow", , GetLocalParamString(), UNDO_Layer
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub Form_Load()
    
    cmdBar.SetPreviewStatus False
    
    Set m_Randomize = New pdRandomize
    m_Randomize.SetSeed_AutomaticAndRandom
    
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

'Update the preview whenever the combination slider/text control has its value changed
Private Sub sldIntensity_Change()
    UpdatePreview
End Sub

Private Sub sldAngle_Change()
    UpdatePreview
End Sub

Private Sub UpdatePreview()
    If cmdBar.PreviewsAllowed Then Me.ApplySnowEffect GetLocalParamString(True), True, pdFxPreview
End Sub

Private Function GetLocalParamString(Optional ByVal isPreview As Boolean = False) As String
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    
    With cParams
        
        .AddParam "intensity", sldIntensity.Value
        .AddParam "size", sldSize.Value
        .AddParam "angle", sldAngle.Value
        .AddParam "wind", sldWind.Value
        .AddParam "opacity", sldOpacity.Value
        
        'Randomizing is a bit weird; we only do it if the user has enabled it, *and* if it's a preview.
        ' (This allows the actual effect to match the last preview the user saw.)
        If (chkRandomize.Value And isPreview) Then m_Randomize.SetSeed_AutomaticAndRandom
        .AddParam "seed", m_Randomize.GetSeed
        
    End With
    
    GetLocalParamString = cParams.GetParamString()
    
End Function

Private Sub sldOpacity_Change()
    UpdatePreview
End Sub

Private Sub sldSize_Change()
    UpdatePreview
End Sub

Private Sub sldWind_Change()
    UpdatePreview
End Sub
