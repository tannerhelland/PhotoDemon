VERSION 5.00
Begin VB.Form FormFxTruchet 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   " Truchet tiles"
   ClientHeight    =   6570
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
   ScaleHeight     =   438
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   806
   Begin PhotoDemon.pdSlider sldLineWidth 
      Height          =   735
      Left            =   9000
      TabIndex        =   10
      Top             =   2760
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   1296
      Caption         =   "width"
      Min             =   1
      Max             =   100
      Value           =   20
      NotchPosition   =   2
      NotchValueCustom=   20
   End
   Begin PhotoDemon.pdSlider sldForeground 
      Height          =   495
      Left            =   9000
      TabIndex        =   8
      Top             =   1575
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   873
      Max             =   100
      Value           =   100
      GradientColorRight=   1703935
      NotchPosition   =   2
      NotchValueCustom=   100
   End
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   375
      Index           =   0
      Left            =   6000
      Top             =   1080
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   661
      Caption         =   "color and opacity"
      FontSize        =   12
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
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1244
      Caption         =   "scale"
      Min             =   3
      Max             =   500
      ScaleStyle      =   1
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
   Begin PhotoDemon.pdDropDown cboBlendMode 
      Height          =   735
      Left            =   6000
      TabIndex        =   3
      Top             =   4680
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1296
      Caption         =   "blend mode"
   End
   Begin PhotoDemon.pdRandomizeUI rndSeed 
      Height          =   735
      Left            =   9000
      TabIndex        =   4
      Top             =   3720
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   1296
      Caption         =   "random seed:"
      MaxLength       =   15
   End
   Begin PhotoDemon.pdColorSelector cpForeground 
      Height          =   495
      Left            =   6000
      TabIndex        =   5
      Top             =   1560
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   873
      curColor        =   0
   End
   Begin PhotoDemon.pdDropDown cboGenerator 
      Height          =   735
      Left            =   6000
      TabIndex        =   6
      Top             =   2760
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   1296
      Caption         =   "shape"
   End
   Begin PhotoDemon.pdColorSelector cpBackground 
      Height          =   495
      Left            =   6000
      TabIndex        =   7
      Top             =   2160
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   873
   End
   Begin PhotoDemon.pdSlider sldBackground 
      Height          =   495
      Left            =   9000
      TabIndex        =   9
      Top             =   2175
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   873
      Max             =   100
      Value           =   100
      GradientColorRight=   1703935
      NotchPosition   =   2
      NotchValueCustom=   100
   End
   Begin PhotoDemon.pdDropDown cboPattern 
      Height          =   735
      Left            =   6000
      TabIndex        =   11
      Top             =   3720
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   1296
      Caption         =   "pattern"
   End
End
Attribute VB_Name = "FormFxTruchet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Render Truchet Tiles Effect
'Copyright 2021-2026 by Tanner Helland
'Created: 03/August/21
'Last updated: 15/August/21
'Last update: wrap up initial build
'
'Truchet tiles date back to the 18th century:
' https://en.wikipedia.org/wiki/Truchet_tiles
'
'PD attempts to provide just enough truchet-related features to be interesting!
' (Note that PD's implementation shares a lot of code with the Render > Clouds effect, by design.)
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'To improve performance, we cache a local temporary DIB when previewing the effect
Private m_tmpDIB As pdDIB

Public Sub FxRenderTruchet(ByVal effectParams As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)

    If (Not toPreview) Then Message "Rendering Truchet tiles..."
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString effectParams
    
    'At present, some parameters are hard-coded.  This is primarily to free up UI space and simplify the
    ' potential set of effect parameters.
    Dim fxScale As Double, fxBlendMode As PD_BlendMode, fxSeed As String
    Dim fxForegroundColor As Long, fxBackgroundColor As Long
    Dim fxForegroundOpacity As Single, fxBackgroundOpacity As Single
    Dim fxShape As PD_TruchetShape, fxPattern As PD_TruchetPattern, fxLineWidth As Single
    
    With cParams
        
        fxScale = .GetDouble("scale", sldScale.Value)
        fxBackgroundColor = .GetLong("background-color", cpBackground.Color)
        fxBackgroundOpacity = .GetLong("background-opacity", sldBackground.Value)
        fxForegroundColor = .GetLong("foreground-color", cpForeground.Color)
        fxForegroundOpacity = .GetLong("foreground-opacity", sldForeground.Value)
        
        fxBlendMode = .GetLong("blendmode", cboBlendMode.ListIndex)
        fxSeed = .GetString("seed")
        fxShape = Filters_Render.GetTruchetShapeFromName(.GetString("shape", vbNullString, True))
        fxPattern = Filters_Render.GetTruchetPatternFromName(.GetString("pattern", vbNullString, True))
        fxLineWidth = .GetSingle("line-width", 1!)
        
    End With
    
    'Random number generation is handled by pdRandomize
    Dim cRandom As pdRandomize
    Set cRandom = New pdRandomize
    cRandom.SetSeed_String fxSeed
    
    'Create a local array and point it at the pixel data of the current image
    Dim tmpSA As SafeArray2D
    EffectPrep.PrepImageData tmpSA, toPreview, dstPic, , , True
    
    'Some values need to be modified to reflect the current preview scale factor
    If toPreview Then
        fxScale = Int(fxScale * curDIBValues.previewModifier + 0.5)
    End If
    
    If (m_tmpDIB Is Nothing) Then Set m_tmpDIB = New pdDIB
    m_tmpDIB.CreateFromExistingDIB workingDIB
    
    'The actual cloud render is handled by a dedicated function
    Filters_Render.GetTruchetDIB m_tmpDIB, fxScale, fxLineWidth, fxForegroundColor, fxForegroundOpacity, fxBackgroundColor, fxBackgroundOpacity, fxShape, fxPattern, cRandom.GetSeed(), toPreview
    
    'Merge the result down, then exit
    Dim cCompositor As pdCompositor
    Set cCompositor = New pdCompositor
    cCompositor.QuickMergeTwoDibsOfEqualSize workingDIB, m_tmpDIB, fxBlendMode, 100!
    EffectPrep.FinalizeImageData toPreview, dstPic, True
    
End Sub

Private Sub cboBlendMode_Click()
    UpdatePreview
End Sub

Private Sub cboGenerator_Click()
    sldLineWidth.Visible = (cboGenerator.ListIndex <= ts_Maze)
    UpdatePreview
End Sub

Private Sub cboPattern_Click()
    rndSeed.Visible = (cboPattern.ListIndex = tp_Random)
    UpdatePreview
End Sub

Private Sub cmdBar_OKClick()
    Process "Truchet", , GetLocalParamString(), UNDO_Layer
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub cmdBar_ResetClick()
    cboBlendMode.ListIndex = BM_Normal
    cpBackground.Color = RGB(255, 255, 255)
    cpForeground.Color = RGB(0, 0, 0)
    cboGenerator.ListIndex = ts_Arc
End Sub

Private Sub cpBackground_ColorChanged()
    UpdatePreview
End Sub

Private Sub cpForeground_ColorChanged()
    UpdatePreview
End Sub

Private Sub Form_Load()

    'Disable previews
    cmdBar.SetPreviewStatus False
    
    'Populate the blend mode drop-down
    Interface.PopulateBlendModeDropDown cboBlendMode, BM_Normal
    
    'Populate available tile shapes and patterns
    cboGenerator.SetAutomaticRedraws False
    Dim i As PD_TruchetShape
    For i = 0 To ts_Max - 1
        cboGenerator.AddItem Filters_Render.GetNameOfTruchetShape(i)
    Next i
    cboGenerator.ListIndex = ts_Arc
    cboGenerator.SetAutomaticRedraws True, True
    
    cboPattern.SetAutomaticRedraws False
    Dim j As PD_TruchetPattern
    For j = 0 To tp_Max - 1
        cboPattern.AddItem Filters_Render.GetNameOfTruchetPattern(j)
    Next j
    cboPattern.ListIndex = tp_Random
    cboPattern.SetAutomaticRedraws True, True
    
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

Private Sub sldBackground_Change()
    UpdatePreview
End Sub

Private Sub sldForeground_Change()
    UpdatePreview
End Sub

Private Sub sldLineWidth_Change()
    UpdatePreview
End Sub

Private Sub sldScale_Change()
    UpdatePreview
End Sub

'Redraw the on-screen preview of the transformed image
Private Sub UpdatePreview()
    If cmdBar.PreviewsAllowed Then FxRenderTruchet GetLocalParamString(), True, pdFxPreview
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
        .AddParam "background-color", cpBackground.Color
        .AddParam "background-opacity", sldBackground.Value
        .AddParam "foreground-color", cpForeground.Color
        .AddParam "foreground-opacity", sldForeground.Value
        .AddParam "blendmode", cboBlendMode.ListIndex
        .AddParam "seed", rndSeed.Value
        .AddParam "shape", Filters_Render.GetNameOfTruchetShape(cboGenerator.ListIndex, False)
        .AddParam "pattern", Filters_Render.GetNameOfTruchetPattern(cboPattern.ListIndex, False)
        .AddParam "line-width", sldLineWidth.Value
    End With
    
    GetLocalParamString = cParams.GetParamString()
    
End Function
