VERSION 5.00
Begin VB.Form FormFog 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   " Fog"
   ClientHeight    =   6555
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11790
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
   ScaleHeight     =   437
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   786
   Begin PhotoDemon.pdRandomizeUI rndSeed 
      Height          =   735
      Left            =   6000
      TabIndex        =   6
      Top             =   4800
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   1296
      Caption         =   "random seed:"
   End
   Begin PhotoDemon.pdCommandBar cmdBar 
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5805
      Width           =   11790
      _ExtentX        =   20796
      _ExtentY        =   1323
   End
   Begin PhotoDemon.pdSlider sltScale 
      Height          =   705
      Left            =   6000
      TabIndex        =   2
      Top             =   600
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   1244
      Caption         =   "scale"
      Min             =   1
      Max             =   100
      SigDigits       =   1
      Value           =   25
      NotchPosition   =   2
      NotchValueCustom=   25
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
   End
   Begin PhotoDemon.pdSlider sltContrast 
      Height          =   705
      Left            =   6000
      TabIndex        =   3
      Top             =   1680
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   1244
      Caption         =   "contrast"
      Max             =   100
      Value           =   50
      NotchPosition   =   2
      NotchValueCustom=   50
   End
   Begin PhotoDemon.pdSlider sltQuality 
      Height          =   705
      Left            =   6000
      TabIndex        =   4
      Top             =   3840
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   1244
      Caption         =   "quality"
      Min             =   1
      Max             =   6
      Value           =   5
      NotchPosition   =   2
      NotchValueCustom=   5
   End
   Begin PhotoDemon.pdSlider sltDensity 
      Height          =   705
      Left            =   6000
      TabIndex        =   5
      Top             =   2760
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   1244
      Caption         =   "density"
      Min             =   1
      Max             =   100
      Value           =   50
      NotchPosition   =   2
      NotchValueCustom=   50
   End
End
Attribute VB_Name = "FormFog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Fog Effect
'Copyright 2002-2026 by Tanner Helland
'Created: 8/April/02
'Last updated: 03/August/17
'Last update: migrate to XML params, minor performance improvements
'
'This tool allows the user to apply a layer of artificial "fog" to an image.  pdNoise is used to generate
' the fog map, using a well-known fractal generation approach to successive layers of noise
' (see http://freespace.virgin.net/hugo.elias/models/m_perlin.htm for details).
'
'A variety of options are provided to help the user find their "ideal" fog.  To simply generate clouds, without any
' trace of the original image, set the Density parameter to 100.  Also, Quality controls the number of successive
' noise planes summed together; there is arguably no visible difference once you exceed 6 (due to the range
' of RGB values involved), but maybe someone out there has sharper eyes than me, and can detect RGB differences
' of 1 or less... ;)
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

Private m_Random As pdRandomize

'To improve performance, we cache a local temporary DIB when previewing the effect
Private m_tmpFogDIB As pdDIB

Public Sub fxFog(ByVal effectParams As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)

    If (Not toPreview) Then Message "Generating artificial fog..."
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString effectParams
    
    Dim fxScale As Double, fxContrast As Double, fxRndSeed As String
    Dim fxDensity As Long, fxQuality As Long
    
    With cParams
        fxScale = .GetDouble("scale", sltScale.Value)
        fxContrast = .GetDouble("contrast", sltContrast.Value)
        fxDensity = .GetLong("density", sltDensity.Value)
        fxQuality = .GetLong("quality", sltQuality.Value)
        fxRndSeed = .GetString("rndseed")
    End With
    
    'Contrast is presented to the user on a [0, 100] scale, but the algorithm needs it on [0, 1]; convert it now
    fxContrast = fxContrast * 0.01
    
    'Quality is presented on a [1, 8] scale; convert it to [0, 7]
    fxQuality = fxQuality - 1
    
    'Create a local array and point it at the pixel data of the current image
    Dim dstImageData() As Byte, dstSA As SafeArray2D
    EffectPrep.PrepImageData dstSA, toPreview, dstPic
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left
    initY = curDIBValues.Top
    finalX = curDIBValues.Right
    finalY = curDIBValues.Bottom
    
    Dim xOffset As Long
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    If (Not toPreview) Then ProgressBars.SetProgBarMax finalY
    progBarCheck = ProgressBars.FindBestProgBarValue()
    
    'Scale is used as a fraction of the image's smallest dimension.  There's no problem with using larger
    ' values, but at some point it distorts the image beyond recognition.
    If (curDIBValues.Width > curDIBValues.Height) Then
        fxScale = (fxScale / 100#) * curDIBValues.Height
    Else
        fxScale = (fxScale / 100#) * curDIBValues.Width
    End If
    
    If (fxScale > 0#) Then fxScale = 1# / fxScale
    
    'Some values can be cached in the interior loop to speed up processing time
    Dim pNoiseCache As Double, xScaleCache As Double, yScaleCache As Double
    
    'Finally, an integer displacement will be used to actually calculate the RGB values at any point in the fog
    Dim pDisplace As Long
    Dim i As Long
    
    'The bulk of the processing time for this function occurs when we set up the initial cloud table; rather than
    ' doing this as part of the RGB assignment array, I've separated it into its own step (in hopes the compiled
    ' will be better able to optimize it!)
    Dim p2Lookup() As Single, p2InvLookup() As Single
    ReDim p2Lookup(0 To fxQuality) As Single, p2InvLookup(0 To fxQuality) As Single
    
    'The fractal noise approach we use requires successive sums of 2 ^ n and 2 ^ -n; we calculate these in advance
    ' as the POW operator is so hideously slow.
    For i = 0 To fxQuality
        p2Lookup(i) = 2 ^ i
        p2InvLookup(i) = 1# / (2 ^ i)
    Next i
    
    'The results of our fog generation will be stored to this array, in [0, 255] format to make the blending step
    ' much faster (as we can simply alpha-blend the results).
    Dim fogArray() As Byte
    ReDim fogArray(initX To finalX, initY To finalY) As Byte
    
    'A pdNoise instance handles the actual noise generation
    Dim cNoise As pdNoise
    Set cNoise = New pdNoise
    
    'To generate "random" values despite using a fixed 2D noise generator, we calculate random offsets
    ' into the "infinite grid" of possible noise values.  This yields (perceptually) random results.
    Dim rndOffsetX As Double, rndOffsetY As Double
    If (m_Random Is Nothing) Then Set m_Random = New pdRandomize
    m_Random.SetSeed_String fxRndSeed
    rndOffsetX = m_Random.GetRandomFloat_WH * 10000000# - 5000000#
    rndOffsetY = m_Random.GetRandomFloat_WH * 10000000# - 5000000#
    
    'Loop through each pixel in the image, converting values as we go
    For y = initY To finalY
    For x = initX To finalX
    
        'Calculate a displacement for this point
        xScaleCache = (CDbl(x) * fxScale)
        yScaleCache = (CDbl(y) * fxScale)
        pNoiseCache = 0#
        
        'Fractal noise works by summing successively smaller noise values taken from successively larger
        ' amplitudes of the original function.
        For i = 0 To fxQuality
            pNoiseCache = pNoiseCache + p2InvLookup(i) * cNoise.SimplexNoise2d(rndOffsetX + xScaleCache * p2Lookup(i), rndOffsetY + yScaleCache * p2Lookup(i))
        Next i
        
        'Apply contrast (e.g. stretch the calculated noise value further)
        pNoiseCache = pNoiseCache * fxContrast
        
        'Convert the calculated noise value to RGB range and cache it
        pDisplace = 127 + (pNoiseCache * 127#)
        If (pDisplace > 255) Then
            pDisplace = 255
        ElseIf (pDisplace < 0) Then
            pDisplace = 0
        End If
        
        fogArray(x, y) = pDisplace
        
    Next x
        If (Not toPreview) Then
            If (y And progBarCheck) = 0 Then
                If Interface.UserPressedESC() Then Exit For
                SetProgBarVal y
            End If
        End If
    Next y
    
    'Next, create a temporary DIB that will hold a grayscale representation of our fog data
    If (m_tmpFogDIB Is Nothing) Then Set m_tmpFogDIB = New pdDIB
    m_tmpFogDIB.CreateFromExistingDIB workingDIB
    m_tmpFogDIB.WrapArrayAroundDIB dstImageData, dstSA
    
    'Loop through each pixel in the image, converting stored fog values to RGB triplets
    For y = initY To finalY
    For x = initX To finalX
    
        pDisplace = fogArray(x, y)
        xOffset = x * 4
        dstImageData(xOffset, y) = pDisplace
        dstImageData(xOffset + 1, y) = pDisplace
        dstImageData(xOffset + 2, y) = pDisplace
        
        'Alpha raises an interesting question.  Do we leave it as-is, or forcibly set it to some new value?
        ' At present, we assume the alpha value from the base image.
        
    Next x
        If (Not toPreview) Then
            If ((y + finalY) And progBarCheck) = 0 Then
                If Interface.UserPressedESC() Then Exit For
            End If
        End If
    Next y
    
    'tmpFogDIB now contains a grayscale representation of our fog data
    m_tmpFogDIB.UnwrapArrayFromDIB dstImageData
    
    'Apply premultiplication prior to compositing
    m_tmpFogDIB.SetAlphaPremultiplication True
    workingDIB.SetAlphaPremultiplication True
    
    'Composite our custom fog image against the base layer (workingDIB) using the Normal blend mode,
    ' and adjusting opacity (taken from the Density option provided to the user).
    Dim cComposite As pdCompositor
    Set cComposite = New pdCompositor
    cComposite.QuickMergeTwoDibsOfEqualSize workingDIB, m_tmpFogDIB, BM_Normal, fxDensity
    
    If (Not toPreview) Then Set m_tmpFogDIB = Nothing
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    EffectPrep.FinalizeImageData toPreview, dstPic, True
        
End Sub

Private Sub cmdBar_OKClick()
    Process "Fog", , GetLocalParamString(), UNDO_Layer
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub cmdBar_ResetClick()
    sltScale = 25
    sltContrast = 50
    sltDensity = 50
    sltQuality = 5
End Sub

Private Sub Form_Load()

    'Disable previews
    cmdBar.SetPreviewStatus False
    
    'pdRandomize is used for all random number generation in PD
    Set m_Random = New pdRandomize
    
    'Apply visual themes and translations
    ApplyThemeAndTranslations Me, True, True
    cmdBar.SetPreviewStatus True
    UpdatePreview
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub rndSeed_Change()
    UpdatePreview
End Sub

Private Sub sltContrast_Change()
    UpdatePreview
End Sub

Private Sub sltDensity_Change()
    UpdatePreview
End Sub

Private Sub sltQuality_Change()
    UpdatePreview
End Sub

Private Sub sltScale_Change()
    UpdatePreview
End Sub

'Redraw the on-screen preview of the transformed image
Private Sub UpdatePreview()
    If cmdBar.PreviewsAllowed Then fxFog GetLocalParamString(), True, pdFxPreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub pdFxPreview_ViewportChanged()
    UpdatePreview
End Sub

Private Function GetLocalParamString() As String
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    
    With cParams
        .AddParam "scale", sltScale.Value
        .AddParam "contrast", sltContrast.Value
        .AddParam "density", sltDensity.Value
        .AddParam "quality", sltQuality.Value
        .AddParam "rndseed", rndSeed.Value
    End With
    
    GetLocalParamString = cParams.GetParamString()
    
End Function
