VERSION 5.00
Begin VB.Form FormFxFibers 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   " Fibers"
   ClientHeight    =   7155
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
   ScaleHeight     =   477
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   806
   Begin PhotoDemon.pdButtonStrip btsType 
      Height          =   975
      Left            =   6000
      TabIndex        =   10
      Top             =   1800
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1720
      Caption         =   "colors"
   End
   Begin PhotoDemon.pdGradientSelector grdColors 
      Height          =   855
      Left            =   6000
      TabIndex        =   9
      Top             =   2880
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
      Top             =   6405
      Width           =   12090
      _ExtentX        =   21325
      _ExtentY        =   1323
   End
   Begin PhotoDemon.pdSlider sldSize 
      Height          =   705
      Left            =   6000
      TabIndex        =   2
      Top             =   120
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   1244
      Caption         =   "length"
      Max             =   500
      ScaleStyle      =   1
      Value           =   25
      NotchPosition   =   2
      NotchValueCustom=   25
   End
   Begin PhotoDemon.pdFxPreviewCtl pdFxPreview 
      Height          =   6225
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   10980
      DisableZoomPan  =   -1  'True
   End
   Begin PhotoDemon.pdSlider sldOpacity 
      Height          =   705
      Left            =   6000
      TabIndex        =   3
      Top             =   3840
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
      Top             =   4680
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1296
      Caption         =   "blend mode"
   End
   Begin PhotoDemon.pdRandomizeUI rndSeed 
      Height          =   735
      Left            =   6000
      TabIndex        =   5
      Top             =   5580
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1296
      Caption         =   "random seed:"
   End
   Begin PhotoDemon.pdColorSelector cpHighlight 
      Height          =   855
      Left            =   9000
      TabIndex        =   6
      Top             =   2880
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
      Top             =   2880
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   1508
      Caption         =   "shadow color"
      curColor        =   50
   End
   Begin PhotoDemon.pdSlider sldStrength 
      Height          =   705
      Left            =   9000
      TabIndex        =   8
      Top             =   120
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   1244
      Caption         =   "strength"
      Min             =   1
      Max             =   99
      ScaleExponent   =   0
      Value           =   50
      NotchPosition   =   2
      NotchValueCustom=   50
   End
   Begin PhotoDemon.pdSlider sldNoise 
      Height          =   705
      Left            =   6000
      TabIndex        =   11
      Top             =   960
      Width           =   2895
      _ExtentX        =   10398
      _ExtentY        =   1270
      Caption         =   "grain"
      Max             =   100
      SigDigits       =   1
      ScaleStyle      =   1
      Value           =   10
      NotchPosition   =   2
      NotchValueCustom=   10
   End
   Begin PhotoDemon.pdSlider sldContrast 
      Height          =   705
      Left            =   9000
      TabIndex        =   12
      Top             =   960
      Width           =   2895
      _ExtentX        =   10398
      _ExtentY        =   1270
      Caption         =   "contrast"
      Min             =   -100
      Max             =   100
      NotchPosition   =   2
   End
End
Attribute VB_Name = "FormFxFibers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Render Fibers Effect
'Copyright 2019-2026 by Tanner Helland
'Created: 01/August/19
'Last updated: 01/August/19
'Last update: initial build
'
'Render clouds has been available in Photoshop since... CS1, I think?  This is a rough analog, with some verbiage
' changes to better match similar options in other PD tools.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'To improve performance, we cache a local temporary DIB when previewing the effect
Private m_tmpFiberDIB As pdDIB, m_tmpMergeDIB As pdDIB

Public Sub FxRenderFibers(ByVal effectParams As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)

    If (Not toPreview) Then Message "Rendering fibers..."
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString effectParams
    
    'At present, some parameters are hard-coded.  This is primarily to free up UI space and simplify the
    ' potential set of effect parameters.
    Dim fxSize As Double, fxStrength As Double, fxNoise As Double, fxContrast As Double
    Dim fxOpacity As Double, fxBlendMode As PD_BlendMode, fxSeed As String
    Dim fxColorShadow As Long, fxColorHighlight As Long
    Dim cGradient As pd2DGradient, useGradient As Boolean, fxGenerator As PD_NoiseGenerator
    
    With cParams
        
        fxSize = .GetLong("size", sldSize.Value)
        fxStrength = .GetDouble("strength", sldStrength.Value)
        fxNoise = .GetDouble("noise", sldNoise.Value)
        fxContrast = .GetDouble("contrast", sldContrast.Value)
        fxOpacity = .GetDouble("opacity", sldOpacity.Value)
        fxBlendMode = .GetLong("blendmode", cboBlendMode.ListIndex)
        fxSeed = .GetString("seed")
        fxGenerator = .GetLong("noise-generator", ng_Simplex)
        
        useGradient = .GetBool("use-gradient", False)
        Set cGradient = New pd2DGradient
        
        If useGradient Then
            cGradient.CreateGradientFromString .GetString("gradient", vbNullString, True)
        Else
            fxColorShadow = .GetLong("shadow-color", vbBlack)
            fxColorHighlight = .GetLong("highlight-color", vbWhite)
            cGradient.CreateTwoPointGradient fxColorShadow, fxColorHighlight
        End If
        
    End With
    
    'Modify input variables to be computationally friendly (e.g. shrink variance to the
    ' range [0, 1] and rotate the angle by 90 degrees so that "vertical" is the default)
    If (fxStrength < 1#) Then fxStrength = 1#
    fxStrength = 1# - (fxStrength * 0.01)
    
    'Create a local array and point it at the pixel data of the current image
    Dim tmpSA As SafeArray2D
    EffectPrep.PrepImageData tmpSA, toPreview, dstPic, , , True
    
    'If this is a preview, we need to adjust the radius to match the size of the preview box
    If toPreview Then
        fxSize = fxSize * curDIBValues.previewModifier
        If (fxSize < 1) Then fxSize = 1
    
    'If this is *not* a preview, we need to figure out a progressbar max value.  At present,
    ' we simply break the fiber generation process into discrete steps (five) and update
    ' after each step is complete.
    Else
        ProgressBars.SetProgBarVal 0
        ProgressBars.SetProgBarMax 5
    End If
    
    'Generate a fiber image
    Dim cRandom As pdRandomize
    Set cRandom = New pdRandomize
    cRandom.SetSeed_String fxSeed
    
    If (m_tmpFiberDIB Is Nothing) Then Set m_tmpFiberDIB = New pdDIB
    If (m_tmpFiberDIB.GetDIBWidth <> workingDIB.GetDIBWidth) Or (m_tmpFiberDIB.GetDIBHeight <> workingDIB.GetDIBHeight) Then m_tmpFiberDIB.CreateFromExistingDIB workingDIB
    
    If useGradient Then
        
        'Pull a lookup table from the gradient object
        Dim palColors() As Long, numPalColors As Long
        numPalColors = cGradient.GetNumOfNodes * 256
        If (numPalColors < 256) Then numPalColors = 256
        If (numPalColors > 1024) Then numPalColors = 1024
        cGradient.GetLookupTable palColors, numPalColors
        
        Filters_Render.RenderFibers_LUT m_tmpFiberDIB, palColors, fxStrength, cRandom.GetSeed(), True
        
    Else
        Filters_Render.RenderFibers_TwoColor m_tmpFiberDIB, Colors.GetRGBAFromRGBAndA(fxColorShadow, 255), Colors.GetRGBAFromRGBAndA(fxColorHighlight, 255), fxStrength, cRandom.GetSeed(), True
    End If
    
    ProgressBars.SetProgBarVal 1
    
    m_tmpFiberDIB.SetInitialAlphaPremultiplicationState True
    
    'Apply a fast blur
    If (m_tmpMergeDIB Is Nothing) Then Set m_tmpMergeDIB = New pdDIB
    m_tmpMergeDIB.CreateFromExistingDIB m_tmpFiberDIB
    m_tmpMergeDIB.SetInitialAlphaPremultiplicationState True
    If (fxSize > 0) Then Filters_Layers.CreateVerticalBlurDIB fxSize, fxSize, m_tmpFiberDIB, m_tmpMergeDIB, True
    
    ProgressBars.SetProgBarVal 2
    
    'Adjust contrast, if any
    If (fxContrast <> 0#) Then
        
        If m_tmpMergeDIB.GetAlphaPremultiplication() Then m_tmpMergeDIB.SetAlphaPremultiplication False
        
        'Use the calculated mean to complete the look-up table
        Dim conTable() As Byte
        Dim cLUT As pdFilterLUT
        Set cLUT = New pdFilterLUT
        cLUT.FillLUT_Contrast conTable, fxContrast
        cLUT.ApplyLUTToAllColorChannels m_tmpMergeDIB, conTable
        
    End If
    
    ProgressBars.SetProgBarVal 3
    
    'Apply noise, if any
    If (fxNoise <> 0#) Then
    
        If m_tmpMergeDIB.GetAlphaPremultiplication() Then m_tmpMergeDIB.SetAlphaPremultiplication False
        
        Dim xDepth As Long, x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
        xDepth = workingDIB.GetDIBColorDepth \ 8
        initX = 0
        initY = 0
        finalX = (workingDIB.GetDIBWidth - 1) * xDepth
        finalY = workingDIB.GetDIBHeight - 1
        
        Dim pxData() As Byte, tmpSA1D As SafeArray1D, dibPtr As Long, dibStride As Long
        dibPtr = m_tmpMergeDIB.GetDIBPointer
        dibStride = m_tmpMergeDIB.GetDIBStride
        m_tmpMergeDIB.WrapArrayAroundScanline pxData, tmpSA1D, 0
        
        'Color variables
        Dim r As Long, g As Long, b As Long, nColor As Long
        
        'fxNoise is returned on the range [0.0, 100.0].  At maximum strength, we want to scale this to
        ' [-255.0, 255.0], or a large enough number to turn white pixels black (and vice-versa).
        fxNoise = fxNoise * 2.55 * 0.33333333333
        
        'Loop through each pixel in the image, converting values as we go
        For y = initY To finalY
            tmpSA1D.pvData = dibPtr + y * dibStride
        For x = initX To finalX Step xDepth
            
            'Get source pixel color values
            b = pxData(x)
            g = pxData(x + 1)
            r = pxData(x + 2)
            
            'Monochromatic noise - same amount for each color
            nColor = fxNoise * cRandom.GetGaussianFloat_WH()
            r = r + nColor
            g = g + nColor
            b = b + nColor
            
            'Bounds-checking
            If (r > 255) Then r = 255
            If (r < 0) Then r = 0
            If (g > 255) Then g = 255
            If (g < 0) Then g = 0
            If (b > 255) Then b = 255
            If (b < 0) Then b = 0
            
            'Assign new colors
            pxData(x) = b
            pxData(x + 1) = g
            pxData(x + 2) = r
            
        Next x
            If (Not toPreview) Then
                If Interface.UserPressedESC() Then Exit For
            End If
        Next y
    
        m_tmpMergeDIB.UnwrapArrayFromDIB pxData
        
    End If
    
    ProgressBars.SetProgBarVal 4
    
    If (Not m_tmpMergeDIB.GetAlphaPremultiplication) Then m_tmpMergeDIB.SetAlphaPremultiplication True
    
    'Merge the result down, then exit
    Dim cCompositor As pdCompositor
    Set cCompositor = New pdCompositor
    cCompositor.QuickMergeTwoDibsOfEqualSize workingDIB, m_tmpMergeDIB, fxBlendMode, fxOpacity
    ProgressBars.SetProgBarVal 5
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

Private Sub cmdBar_OKClick()
    Process "Fibers", , GetLocalParamString(), UNDO_Layer
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub cmdBar_ResetClick()
    cboBlendMode.ListIndex = BM_Normal
    cpShadow.Color = RGB(0, 0, 0)
    cpHighlight.Color = RGB(255, 255, 255)
End Sub

Private Sub cpHighlight_ColorChanged()
    UpdatePreview
End Sub

Private Sub cpShadow_ColorChanged()
    UpdatePreview
End Sub

Private Sub Form_Load()

    'Disable previews
    cmdBar.SetPreviewStatus False
    
    'Populate the blend mode drop-down
    Interface.PopulateBlendModeDropDown cboBlendMode, BM_Normal
    
    'Populate button strips
    btsType.AddItem "simple", 0
    btsType.AddItem "gradient", 1
    btsType.ListIndex = 0
    UpdateCloudTypeUI
    
    'Apply visual themes and translations
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

Private Sub sldContrast_Change()
    UpdatePreview
End Sub

Private Sub sldNoise_Change()
    UpdatePreview
End Sub

Private Sub sldOpacity_Change()
    UpdatePreview
End Sub

'Redraw the on-screen preview of the transformed image
Private Sub UpdatePreview()
    If cmdBar.PreviewsAllowed Then FxRenderFibers GetLocalParamString(), True, pdFxPreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub pdFxPreview_ViewportChanged()
    UpdatePreview
End Sub

Private Function GetLocalParamString() As String
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    
    With cParams
        .AddParam "size", sldSize.Value
        .AddParam "strength", sldStrength.Value
        .AddParam "noise", sldNoise.Value
        .AddParam "contrast", sldContrast.Value
        .AddParam "opacity", sldOpacity.Value
        .AddParam "blendmode", cboBlendMode.ListIndex
        .AddParam "seed", rndSeed.Value
        .AddParam "use-gradient", CBool(btsType.ListIndex = 1)
        .AddParam "shadow-color", cpShadow.Color
        .AddParam "highlight-color", cpHighlight.Color
        .AddParam "gradient", grdColors.Gradient
    End With
    
    GetLocalParamString = cParams.GetParamString()
    
End Function

Private Sub sldSize_Change()
    UpdatePreview
End Sub

Private Sub sldStrength_Change()
    UpdatePreview
End Sub
