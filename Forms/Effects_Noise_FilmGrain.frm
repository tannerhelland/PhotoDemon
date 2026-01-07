VERSION 5.00
Begin VB.Form FormFilmGrain 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   " Add film grain"
   ClientHeight    =   6540
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11415
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
   ScaleWidth      =   761
   Begin PhotoDemon.pdCommandBar cmdBar 
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5790
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   1323
   End
   Begin PhotoDemon.pdSlider sltNoise 
      Height          =   705
      Left            =   6000
      TabIndex        =   2
      Top             =   1920
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   1244
      Caption         =   "strength"
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
   Begin PhotoDemon.pdSlider sltRadius 
      Height          =   705
      Left            =   6000
      TabIndex        =   3
      Top             =   3000
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   1244
      Caption         =   "softness"
      Max             =   25
      SigDigits       =   1
      Value           =   5
      DefaultValue    =   5
   End
End
Attribute VB_Name = "FormFilmGrain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Add Film Grain Tool
'Copyright 2013-2026 by Tanner Helland
'Created: 31/January/13
'Last updated: 07/August/17
'Last update: convert to XML params, large performance improvements
'
'Tool for simulating film grain. For aesthetic reasons, film grain is restricted to monochromatic noise
' (luminance only) to better mimic traditional film grain.
'
'The separate standalone Gaussian Blur function is used for noise smoothing.
'
'All source code in this file is licensed under a modified BSD license. This means you may use the code in your own
' projects IF you provide attribution. For more information, please visit https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

Public Sub AddFilmGrain(ByVal effectParams As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)
        
    If (Not toPreview) Then Message "Generating film grain texture..."
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString effectParams
    
    Dim gStrength As Double, gSoftness As Double
    
    With cParams
        gStrength = .GetDouble("noise", sltNoise.Value)
        gSoftness = .GetDouble("radius", sltRadius.Value)
    End With
    
    'Create a local array and point it at the pixel data of the current image
    Dim dstSA As SafeArray2D
    EffectPrep.PrepImageData dstSA, toPreview, dstPic, , , True
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left
    initY = curDIBValues.Top
    finalX = curDIBValues.Right
    finalY = curDIBValues.Bottom
    
    'To keep processing quick, only update the progress bar when absolutely necessary. This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    If (Not toPreview) Then
        SetProgBarMax finalY
        progBarCheck = ProgressBars.FindBestProgBarValue()
    End If
        
    'Noise variables
    Dim nColor As Long
    
    'Although it's slow, we're stuck using random numbers for noise addition. Seed the generator with a pseudo-random value.
    Dim cRandom As pdRandomize
    Set cRandom = New pdRandomize
    cRandom.SetSeed_AutomaticAndRandom
    
    'All results are going to be placed inside a byte array, which is faster to manipulate than a DIB.
    Dim noiseBytes() As Byte
    ReDim noiseBytes(0 To finalX, 0 To finalY) As Byte
    
    'Loop through each pixel in the image, converting values as we go
    For y = initY To finalY
    For x = initX To finalX
                    
        'Generate monochromatic noise, e.g. the same amount of noise for each color component, based around RGB(127, 127, 127)
        nColor = 127 + gStrength * cRandom.GetGaussianFloat_WH()
        If (nColor < 0) Then nColor = 0
        If (nColor > 255) Then nColor = 255
        noiseBytes(x, y) = nColor
        
    Next x
        If (Not toPreview) Then
            If (y And progBarCheck) = 0 Then
                If Interface.UserPressedESC() Then Exit For
                SetProgBarVal y
            End If
        End If
    Next y
    
    If (gSoftness > 0#) And (Not g_cancelCurrentAction) Then
        
        'Next, we need to soften the noise DIB
        If (Not toPreview) And (Not g_cancelCurrentAction) Then Message "Softening film grain..."
        
        'If this is a preview, we need to adjust the softening radius to match the size of the preview box
        If toPreview Then
            gSoftness = gSoftness * curDIBValues.previewModifier
            If (gSoftness < 0.1) Then gSoftness = 0.1
        End If
        
        Filters_ByteArray.GaussianBlur_AM_ByteArray noiseBytes, workingDIB.GetDIBWidth, workingDIB.GetDIBHeight, gSoftness, 3
        
    End If
    
    If (Not g_cancelCurrentAction) Then
        
        'As our final operation, merge the noise onto the original image, using pdCompositor
        Dim noiseImage As pdDIB
        Set noiseImage = New pdDIB
        DIBs.CreateDIBFromGrayscaleMap noiseImage, noiseBytes, workingDIB.GetDIBWidth, workingDIB.GetDIBHeight
        
        Dim cCompositor As pdCompositor
        Set cCompositor = New pdCompositor
        cCompositor.QuickMergeTwoDibsOfEqualSize workingDIB, noiseImage, BM_SoftLight, , , AM_Inherit
        
    End If
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    EffectPrep.FinalizeImageData toPreview, dstPic, True
    
End Sub

Private Sub cmdBar_OKClick()
    Process "Add film grain", , GetLocalParamString(), UNDO_Layer
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub Form_Load()
    
    cmdBar.SetPreviewStatus False
    
    'Apply translations and visual themes
    ApplyThemeAndTranslations Me, True, True
    cmdBar.SetPreviewStatus True
    UpdatePreview
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub sltNoise_Change()
    UpdatePreview
End Sub

Private Sub sltRadius_Change()
    UpdatePreview
End Sub

Private Sub UpdatePreview()
    If cmdBar.PreviewsAllowed Then Me.AddFilmGrain GetLocalParamString(), True, pdFxPreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub pdFxPreview_ViewportChanged()
    UpdatePreview
End Sub

Private Function GetLocalParamString() As String
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    
    With cParams
        .AddParam "noise", sltNoise.Value
        .AddParam "radius", sltRadius.Value
    End With
    
    GetLocalParamString = cParams.GetParamString()
    
End Function
