VERSION 5.00
Begin VB.Form FormFxClouds 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   " Clouds"
   ClientHeight    =   6525
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12090
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
   ScaleHeight     =   435
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   806
   Begin PhotoDemon.pdButtonStrip btsType 
      Height          =   975
      Left            =   6000
      TabIndex        =   10
      Top             =   960
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1720
      Caption         =   "colors"
   End
   Begin PhotoDemon.pdGradientSelector grdColors 
      Height          =   855
      Left            =   6000
      TabIndex        =   9
      Top             =   2040
      Visible         =   0   'False
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1508
      Caption         =   "colors"
   End
   Begin PhotoDemon.pdCommandBar cmdBar 
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5775
      Width           =   12090
      _ExtentX        =   21325
      _ExtentY        =   1323
   End
   Begin PhotoDemon.pdSlider sldScale 
      Height          =   705
      Left            =   6000
      TabIndex        =   2
      Top             =   120
      Width           =   2895
      _ExtentX        =   5106
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
      Height          =   5505
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9710
      DisableZoomPan  =   -1  'True
   End
   Begin PhotoDemon.pdSlider sldOpacity 
      Height          =   705
      Left            =   6000
      TabIndex        =   3
      Top             =   3000
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1244
      Caption         =   "opacity"
      Max             =   100
      SigDigits       =   1
      Value           =   100
      NotchValueCustom=   25
      DefaultValue    =   100
   End
   Begin PhotoDemon.pdDropDown cboBlendMode 
      Height          =   735
      Left            =   6000
      TabIndex        =   4
      Top             =   3840
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   1296
      Caption         =   "blend mode"
   End
   Begin PhotoDemon.pdRandomizeUI rndSeed 
      Height          =   735
      Left            =   6000
      TabIndex        =   5
      Top             =   4740
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1296
      Caption         =   "random seed:"
   End
   Begin PhotoDemon.pdColorSelector cpHighlight 
      Height          =   855
      Left            =   9000
      TabIndex        =   6
      Top             =   2040
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   1508
      Caption         =   "highlight color"
      curColor        =   6262010
   End
   Begin PhotoDemon.pdColorSelector cpShadow 
      Height          =   855
      Left            =   6000
      TabIndex        =   7
      Top             =   2040
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   1508
      Caption         =   "shadow color"
      curColor        =   50
   End
   Begin PhotoDemon.pdSlider sldQuality 
      Height          =   705
      Left            =   9000
      TabIndex        =   8
      Top             =   120
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   1244
      Caption         =   "quality"
      Min             =   1
      Max             =   8
      ScaleExponent   =   0
      Value           =   5
      NotchPosition   =   2
      NotchValueCustom=   5
   End
   Begin PhotoDemon.pdDropDown cboGenerator 
      Height          =   735
      Left            =   9000
      TabIndex        =   11
      Top             =   3840
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   1296
      Caption         =   "generator"
   End
End
Attribute VB_Name = "FormFxClouds"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Render Clouds Effect
'Copyright 2019-2026 by Tanner Helland
'Created: 23/July/19
'Last updated: 23/July/19
'Last update: initial build
'
'Render clouds has been available in Photoshop for decades; about time we exposed a similar option in PD.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'To improve performance, we cache a local temporary DIB when previewing the effect
Private m_tmpDIB As pdDIB

Public Sub FxRenderClouds(ByVal effectParams As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)

    If (Not toPreview) Then Message "Rendering clouds..."
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString effectParams
    
    'At present, some parameters are hard-coded.  This is primarily to free up UI space and simplify the
    ' potential set of effect parameters.
    Dim fxScale As Double, fxOpacity As Double, fxBlendMode As PD_BlendMode, fxSeed As String
    Dim fxColorShadow As Long, fxColorHighlight As Long, fxQuality As Long
    Dim cGradient As pd2DGradient, useGradient As Boolean, fxGenerator As PD_NoiseGenerator
    
    With cParams
        
        fxScale = .GetDouble("scale", sldScale.Value)
        fxQuality = .GetLong("quality", sldQuality.Value)
        fxOpacity = .GetDouble("opacity", sldOpacity.Value)
        fxBlendMode = .GetLong("blendmode", cboBlendMode.ListIndex)
        fxSeed = .GetString("seed")
        fxGenerator = .GetLong("noise-generator", ng_Simplex)
        
        useGradient = .GetBool("use-gradient", False)
        Set cGradient = New pd2DGradient
        
        If useGradient Then
            cGradient.CreateGradientFromString .GetString("gradient", vbNullString, True)
        Else
            fxColorShadow = .GetLong("shadow-color", RGB(50, 0, 0))
            fxColorHighlight = .GetLong("highlight-color", RGB(250, 140, 95))
            cGradient.CreateTwoPointGradient fxColorShadow, fxColorHighlight
        End If
        
    End With
    
    'Random number generation is handled by pdRandomize
    Dim cRandom As pdRandomize
    Set cRandom = New pdRandomize
    cRandom.SetSeed_String fxSeed
    
    'Create a local array and point it at the pixel data of the current image
    Dim tmpSA As SafeArray2D
    EffectPrep.PrepImageData tmpSA, toPreview, dstPic, , , True
    
    If (m_tmpDIB Is Nothing) Then Set m_tmpDIB = New pdDIB
    m_tmpDIB.CreateFromExistingDIB workingDIB
    
    'Pull a lookup table from the gradient object
    Dim palColors() As Long, numPalColors As Long
    numPalColors = cGradient.GetNumOfNodes * 256
    If (numPalColors < 256) Then numPalColors = 256
    If (numPalColors > 1024) Then numPalColors = 1024
    cGradient.GetLookupTable palColors, numPalColors
    
    'The actual cloud render is handled by a dedicated function
    Filters_Render.GetCloudDIB m_tmpDIB, fxScale, VarPtr(palColors(0)), numPalColors, fxGenerator, fxQuality, cRandom.GetSeed(), toPreview
    
    'Merge the result down, then exit
    Dim cCompositor As pdCompositor
    Set cCompositor = New pdCompositor
    cCompositor.QuickMergeTwoDibsOfEqualSize workingDIB, m_tmpDIB, fxBlendMode, fxOpacity
    EffectPrep.FinalizeImageData toPreview, dstPic, True
    
End Sub

Private Sub btsType_Click(ByVal buttonIndex As Long)
    UpdateCloudTypeUI
    UpdatePreview
End Sub

Private Sub UpdateCloudTypeUI()
    cpShadow.Visible = (btsType.ListIndex = 0)
    cpHighlight.Visible = (btsType.ListIndex = 0)
    grdColors.Visible = (btsType.ListIndex = 1)
End Sub

Private Sub cboBlendMode_Click()
    UpdatePreview
End Sub

Private Sub cboGenerator_Click()
    UpdatePreview
End Sub

Private Sub cmdBar_OKClick()
    Process "Clouds", , GetLocalParamString(), UNDO_Layer
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub cmdBar_ResetClick()
    cboBlendMode.ListIndex = BM_Normal
    cpShadow.Color = RGB(0, 0, 0)
    cpHighlight.Color = RGB(255, 255, 255)
    cboGenerator.ListIndex = ng_Simplex
End Sub

Private Sub cpHighlight_ColorChanged()
    UpdatePreview
End Sub

Private Sub cpShadow_ColorChanged()
    UpdatePreview
End Sub

Private Sub Form_Load()

    'Disable previews during initialization
    cmdBar.SetPreviewStatus False
    
    'Populate the blend mode drop-down
    Interface.PopulateBlendModeDropDown cboBlendMode, BM_Normal, True
    
    'Populate button strips
    btsType.AddItem "simple", 0
    btsType.AddItem "gradient", 1
    btsType.ListIndex = 0
    UpdateCloudTypeUI
    
    cboGenerator.SetAutomaticRedraws False
    cboGenerator.AddItem "Perlin", 0
    cboGenerator.AddItem "Simplex", 1
    cboGenerator.AddItem "OpenSimplex", 2
    cboGenerator.ListIndex = 1
    cboGenerator.SetAutomaticRedraws True, True
    
    'Apply visual themes and translations, then enable previews
    ApplyThemeAndTranslations Me, True, True
    cmdBar.SetPreviewStatus True
    UpdatePreview
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub grdColors_GradientChanged()
    UpdatePreview
End Sub

Private Sub rndSeed_Change()
    UpdatePreview
End Sub

Private Sub sldOpacity_Change()
    UpdatePreview
End Sub

Private Sub sldQuality_Change()
    UpdatePreview
End Sub

Private Sub sldScale_Change()
    UpdatePreview
End Sub

'Redraw the on-screen preview of the transformed image
Private Sub UpdatePreview()
    If cmdBar.PreviewsAllowed Then FxRenderClouds GetLocalParamString(), True, pdFxPreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub pdFxPreview_ViewportChanged()
    UpdatePreview
End Sub

Private Function GetLocalParamString() As String
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    
    With cParams
        .AddParam "scale", sldScale.Value
        .AddParam "quality", sldQuality.Value
        .AddParam "opacity", sldOpacity.Value
        .AddParam "blendmode", cboBlendMode.ListIndex
        .AddParam "seed", rndSeed.Value
        .AddParam "use-gradient", CBool(btsType.ListIndex = 1)
        .AddParam "shadow-color", cpShadow.Color
        .AddParam "highlight-color", cpHighlight.Color
        .AddParam "gradient", grdColors.Gradient
        .AddParam "noise-generator", cboGenerator.ListIndex
    End With
    
    GetLocalParamString = cParams.GetParamString()
    
End Function
