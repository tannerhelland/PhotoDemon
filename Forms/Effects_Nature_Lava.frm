VERSION 5.00
Begin VB.Form FormLava 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   " Lava"
   ClientHeight    =   6555
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
   ScaleHeight     =   437
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   806
   Begin PhotoDemon.pdCommandBar cmdBar 
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
      Top             =   360
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
   Begin PhotoDemon.pdSlider sldOpacity 
      Height          =   705
      Left            =   6000
      TabIndex        =   3
      Top             =   1320
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
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
      Top             =   2280
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1296
      Caption         =   "blend mode"
   End
   Begin PhotoDemon.pdRandomizeUI rndSeed 
      Height          =   735
      Left            =   6000
      TabIndex        =   5
      Top             =   4680
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1296
      Caption         =   "random seed:"
   End
   Begin PhotoDemon.pdColorSelector cpHighlight 
      Height          =   975
      Left            =   9000
      TabIndex        =   6
      Top             =   3360
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   1720
      Caption         =   "highlight color"
      curColor        =   6262010
   End
   Begin PhotoDemon.pdColorSelector cpShadow 
      Height          =   975
      Left            =   6000
      TabIndex        =   7
      Top             =   3360
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   1720
      Caption         =   "shadow color"
      curColor        =   50
   End
End
Attribute VB_Name = "FormLava"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Lava Effect
'Copyright 2002-2026 by Tanner Helland
'Created: 8/April/02
'Last updated: 09/May/19
'Last update: fix potential overflow on 32-bpp images with fully transparent regions
'
'This (silly) effect uses a combination of a pdNoise instance (for generating a base fog-like effect),
' which is then chrome-ified in red/orange hues, rotated 180 degrees, and merged onto itself to create
' a lava-like map.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'To improve performance, we cache a local temporary DIB when previewing the effect
Private m_tmpDIB As pdDIB

Public Sub fxLava(ByVal effectParams As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)

    If (Not toPreview) Then Message "Exploding imaginary volcano..."
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString effectParams
    
    'At present, some parameters are hard-coded.  This is primarily to free up UI space and simplify the
    ' potential set of effect parameters.
    Dim fxScale As Double, fxOpacity As Double, fxBlendMode As PD_BlendMode, fxSeed As String
    Dim fxColorShadow As Long, fxColorHighlight As Long
    
    With cParams
        fxScale = .GetDouble("scale", sltScale.Value)
        fxOpacity = .GetDouble("opacity", sldOpacity.Value)
        fxBlendMode = .GetLong("blendmode", cboBlendMode.ListIndex)
        fxSeed = .GetString("seed")
        fxColorShadow = .GetLong("shadow-color", RGB(50, 0, 0))
        fxColorHighlight = .GetLong("highlight-color", RGB(250, 140, 95))
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
    
    'Generate a palette for the incoming colors
    Dim palColors() As RGBQuad
    Palettes.GetPalette_Grayscale palColors
    
    'The initial noise render is handled by a dedicated function
    Filters_Render.GetCloudDIB m_tmpDIB, fxScale, VarPtr(palColors(0)), 256, ng_Simplex, 4, cRandom.GetSeed(), toPreview, m_tmpDIB.GetDIBHeight + m_tmpDIB.GetDIBWidth, 0
    
    'Chrome-ify it using hard-coded "lava" colors
    Filters_Natural.GetChromeDIB m_tmpDIB, 8, fxScale * 0.25, fxColorShadow, fxColorHighlight, toPreview, m_tmpDIB.GetDIBHeight + m_tmpDIB.GetDIBWidth, m_tmpDIB.GetDIBHeight
    
    'Duplicate that layer
    Dim rotDIB As pdDIB
    Set rotDIB = New pdDIB
    rotDIB.CreateFromExistingDIB m_tmpDIB
    
    'Rotate the DIB 180 degrees
    GDI_Plus.GDIPlusRotateFlip_InPlace rotDIB, GP_RF_180FlipNone
    
    'Merge the result back onto the original temporary DIB
    Dim cCompositor As pdCompositor
    Set cCompositor = New pdCompositor
    cCompositor.QuickMergeTwoDibsOfEqualSize m_tmpDIB, rotDIB, BM_VividLight
    
    'Free our rotated DIB
    Set rotDIB = Nothing
    
    'Merge the result down, then exit
    cCompositor.QuickMergeTwoDibsOfEqualSize workingDIB, m_tmpDIB, fxBlendMode, fxOpacity
    EffectPrep.FinalizeImageData toPreview, dstPic, True
    
End Sub

Private Sub cboBlendMode_Click()
    UpdatePreview
End Sub

Private Sub cmdBar_OKClick()
    Process "Lava", , GetLocalParamString(), UNDO_Layer
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub cmdBar_ResetClick()
    cboBlendMode.ListIndex = BM_Overlay
    cpShadow.Color = RGB(50, 0, 0)
    cpHighlight.Color = RGB(250, 140, 95)
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
    Interface.PopulateBlendModeDropDown cboBlendMode, BM_Overlay
    
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

Private Sub sldOpacity_Change()
    UpdatePreview
End Sub

Private Sub sltScale_Change()
    UpdatePreview
End Sub

'Redraw the on-screen preview of the transformed image
Private Sub UpdatePreview()
    If cmdBar.PreviewsAllowed Then fxLava GetLocalParamString(), True, pdFxPreview
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
        .AddParam "opacity", sldOpacity.Value
        .AddParam "blendmode", cboBlendMode.ListIndex
        .AddParam "seed", rndSeed.Value
        .AddParam "shadow-color", cpShadow.Color
        .AddParam "highlight-color", cpHighlight.Color
    End With
    
    GetLocalParamString = cParams.GetParamString()
    
End Function
