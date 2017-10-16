VERSION 5.00
Begin VB.Form FormFog 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Fog"
   ClientHeight    =   6555
   ClientLeft      =   -15
   ClientTop       =   225
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
   ScaleHeight     =   437
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   806
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.pdButton cmdRandomize 
      Height          =   600
      Left            =   6600
      TabIndex        =   6
      Top             =   4920
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   1058
      Caption         =   "Randomize cloud base"
   End
   Begin PhotoDemon.pdCommandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5805
      Width           =   12090
      _ExtentX        =   21325
      _ExtentY        =   1323
   End
   Begin PhotoDemon.pdSlider sltScale 
      Height          =   705
      Left            =   6000
      TabIndex        =   2
      Top             =   600
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
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
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
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
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
      Caption         =   "quality"
      Min             =   1
      Max             =   8
      Value           =   5
      NotchPosition   =   2
      NotchValueCustom=   5
   End
   Begin PhotoDemon.pdSlider sltDensity 
      Height          =   705
      Left            =   6000
      TabIndex        =   5
      Top             =   2760
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
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
'Copyright 2002-2017 by Tanner Helland
'Created: 8/April/02
'Last updated: 03/August/17
'Last update: migrate to XML params, minor performance improvements
'
'This tool allows the user to apply a layer of artificial "fog" to an image.  Perlin Noise is used to generate
' the fog map, using a well-known fractal generation approach to successive layers of noise
' (see http://freespace.virgin.net/hugo.elias/models/m_perlin.htm for details).
'
'A variety of options are provided to help the user find their "ideal" fog.  To simply generate clouds, without any
' trace of the original image, set the Density parameter to 100.  Also, Quality controls the number of successive
' Perlin Noise planes summed together; there is arguably no visible difference once you exceed 6 (due to the range
' of RGB values involved), but maybe someone out there has sharper eyes than me, and can detect RGB differences
' of 1 or less... ;)
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'This variable stores random z-location in the perlin noise generator (which allows for a unique effect each time the form is loaded)
Private m_zOffset As Double
Private m_Random As pdRandomize

'To improve performance, we cache a local temporary DIB when previewing the effect
Private m_tmpFogDIB As pdDIB

Private Sub cmbEdges_Click()
    UpdatePreview
End Sub

'Apply a "fog" effect to an image, using Perlin Noise as the base
Public Sub fxFog(ByVal effectParams As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)

    If (Not toPreview) Then Message "Generating artificial fog..."
    
    Dim cParams As pdParamXML
    Set cParams = New pdParamXML
    cParams.SetParamString effectParams
    
    Dim fxScale As Double, fxContrast As Double
    Dim fxDensity As Long, fxQuality As Long
    
    With cParams
        fxScale = .GetDouble("scale", sltScale.Value)
        fxContrast = .GetDouble("contrast", sltContrast.Value)
        fxDensity = .GetLong("density", sltDensity.Value)
        fxQuality = .GetLong("quality", sltQuality.Value)
    End With
    
    'Contrast is presented to the user on a [0, 100] scale, but the algorithm needs it on [0, 1]; convert it now
    fxContrast = fxContrast / 100#
    
    'Quality is presented on a [1, 8] scale; convert it to [0, 7]
    fxQuality = fxQuality - 1
    
    'Create a local array and point it at the pixel data of the current image
    Dim dstImageData() As Byte
    Dim dstSA As SafeArray2D
    EffectPrep.PrepImageData dstSA, toPreview, dstPic
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left
    initY = curDIBValues.Top
    finalX = curDIBValues.Right
    finalY = curDIBValues.Bottom
    
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim xOffset As Long, qvDepth As Long
    qvDepth = curDIBValues.BytesPerPixel
    
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
    
    'This effect requires a noise function to operate.  I use Steve McMahon's excellent Perlin Noise class for this.
    Dim cPerlin As cPerlin3D
    Set cPerlin = New cPerlin3D
    
    'Cache the z-value used in the Perlin Noise function.  This is faster than constantly passing
    ' it as a value.  (Note that this caching mechanism and resulting function is NOT part of
    ' Steve's initial implementation, so if it gives anyone trouble, blame me!)
    cPerlin.cacheZValue m_zOffset
    
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
    
    'Loop through each pixel in the image, converting values as we go
    For y = initY To finalY
    For x = initX To finalX
    
        'Calculate a displacement for this point, using perlin noise as the basis, but modifying it per the
        ' user's turbulence value.
        xScaleCache = x * fxScale
        yScaleCache = y * fxScale
        pNoiseCache = 0#
        
        'Fractal noise works by summing successively smaller perlin noise values taken from successively larger
        ' amplitudes of the original function.
        For i = 0 To fxQuality
            pNoiseCache = pNoiseCache + p2InvLookup(i) * cPerlin.Noise2D(p2Lookup(i) * xScaleCache, p2Lookup(i) * yScaleCache)
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
        xOffset = x * qvDepth
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
    
    'A pdCompositor class will help us selectively blend the fog results back onto the main image
    Dim cComposite As pdCompositor
    Set cComposite = New pdCompositor
    
    'Composite our custom fog image against the base layer (workingDIB) using the Normal blend mode,
    ' and adjusting opacity (taken from the Density option provided to the user).
    cComposite.QuickMergeTwoDibsOfEqualSize workingDIB, m_tmpFogDIB, BL_NORMAL, fxDensity
    
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
    
    'Calculate a random z offset for the noise function
    m_Random.SetSeed_AutomaticAndRandom
    m_zOffset = m_Random.GetRandomFloat_WH() * &HEFFFFFFF
    
End Sub

Private Sub cmdRandomize_Click()

    'Calculate a random z offset for the noise function
    m_Random.SetSeed_AutomaticAndRandom
    m_zOffset = m_Random.GetRandomFloat_WH() * &HEFFFFFFF
    
    UpdatePreview

End Sub

Private Sub Form_Load()

    'Disable previews
    cmdBar.MarkPreviewStatus False
    
    'Calculate a random z offset for the noise function
    Set m_Random = New pdRandomize
    m_Random.SetSeed_AutomaticAndRandom
    m_zOffset = m_Random.GetRandomFloat_WH() * &HEFFFFFFF
    
    'Apply visual themes and translations
    ApplyThemeAndTranslations Me
    cmdBar.MarkPreviewStatus True
    UpdatePreview
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
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
    
    Dim cParams As pdParamXML
    Set cParams = New pdParamXML
    
    With cParams
        .AddParam "scale", sltScale.Value
        .AddParam "contrast", sltContrast.Value
        .AddParam "density", sltDensity.Value
        .AddParam "quality", sltQuality.Value
    End With
    
    GetLocalParamString = cParams.GetParamString()
    
End Function
