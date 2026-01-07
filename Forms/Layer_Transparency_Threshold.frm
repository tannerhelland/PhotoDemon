VERSION 5.00
Begin VB.Form FormThresholdAlpha 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   " Threshold alpha"
   ClientHeight    =   6540
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
   ScaleHeight     =   436
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   786
   Begin PhotoDemon.pdColorSelector csMatte 
      Height          =   975
      Left            =   6000
      TabIndex        =   5
      Top             =   3600
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   1720
      Caption         =   "matte"
   End
   Begin PhotoDemon.pdSlider sldDitheringAmount 
      Height          =   855
      Left            =   6000
      TabIndex        =   4
      Top             =   2640
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   1508
      Caption         =   "dithering amount"
      Max             =   100
      Value           =   50
      NotchPosition   =   2
      NotchValueCustom=   50
   End
   Begin PhotoDemon.pdDropDown cboDither 
      Height          =   855
      Left            =   6000
      TabIndex        =   3
      Top             =   1680
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   1508
      Caption         =   "dithering"
   End
   Begin PhotoDemon.pdSlider sldThreshold 
      Height          =   705
      Left            =   6000
      TabIndex        =   2
      Top             =   840
      Width           =   5565
      _ExtentX        =   9816
      _ExtentY        =   1244
      Caption         =   "threshold"
      Min             =   1
      Max             =   254
      Value           =   127
      NotchPosition   =   2
      NotchValueCustom=   127
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
      Width           =   11790
      _ExtentX        =   20796
      _ExtentY        =   1323
   End
End
Attribute VB_Name = "FormThresholdAlpha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Threshold alpha dialog
'Copyright 2020-2026 by Tanner Helland
'Created: 15/May/20
'Last updated: 15/May/20
'Last update: initial build
'
'This tool allows the user to reduce an image's alpha channel to just two values: 0 and 255.
' This is identical to converting an image to monochrome (black and white), except it is
' only applied to an image's alpha channel.
'
'This is useful when producing "retro" imagery, or when saving to a legacy file format like
' GIF or ICO.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

Private Sub cboDither_Click()
    sldDitheringAmount.Visible = (cboDither.ListIndex <> 0)
    UpdatePreview
End Sub

'OK button
Private Sub cmdBar_OKClick()
    Process "Threshold alpha", , GetFunctionParamString(), UNDO_Layer
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

'When resetting, set the color boxes to black and white, and the dithering combo box to 6 (Stucki)
Private Sub cmdBar_ResetClick()
    
    'Stucki dithering w/out bleed reduction
    cboDither.ListIndex = 6
    
    'Standard threshold value
    sldThreshold.Reset
    
End Sub

Private Function GetFunctionParamString() As String
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    With cParams
        .AddParam "threshold", sldThreshold.Value
        .AddParam "dither", cboDither.ListIndex
        .AddParam "dither-amount", sldDitheringAmount.Value
        .AddParam "matte-color", csMatte.Color
    End With
    GetFunctionParamString = cParams.GetParamString
End Function

Private Sub csMatte_ColorChanged()
    UpdatePreview
End Sub

Private Sub Form_Load()
    
    cmdBar.SetPreviewStatus False
    
    'Populate the dither dropdown
    Palettes.PopulateDitheringDropdown cboDither
    cboDither.ListIndex = 6
    
    'Apply translations and visual themes
    ApplyThemeAndTranslations Me, True, True
    cmdBar.SetPreviewStatus True
    UpdatePreview
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'Convert an image to black and white (1-bit image)
Public Sub FxThresholdAlpha(ByVal fxParams As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)
    
    If (Not toPreview) Then Message "Converting alpha channel..."
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString fxParams
    
    Dim cThreshold As Long, ditherMethod As PD_DITHER_METHOD, ditherAmount As Single, matteColor As Long
    With cParams
        cThreshold = .GetLong("threshold", 127)
        ditherMethod = .GetLong("dither", 6)
        ditherAmount = .GetSingle("dither-amount", 100!)
        matteColor = .GetLong("matte-color", vbWhite)
    End With
    
    'Create a local array and point it at the pixel data we want to operate on
    Dim tmpSA As SafeArray2D
    EffectPrep.PrepImageData tmpSA, toPreview, dstPic, doNotUnPremultiplyAlpha:=True
    
    'Pass handling off to the dedicated alpha-threshold function
    DIBs.ThresholdAlphaChannel workingDIB, cThreshold, ditherMethod, ditherAmount, matteColor, toPreview
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    EffectPrep.FinalizeImageData toPreview, dstPic, True

End Sub

Private Sub sldDitheringAmount_Change()
    UpdatePreview
End Sub

Private Sub sldThreshold_Change()
    UpdatePreview
End Sub

Private Sub UpdatePreview()
    If cmdBar.PreviewsAllowed Then FxThresholdAlpha GetFunctionParamString(), True, pdFxPreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub pdFxPreview_ViewportChanged()
    UpdatePreview
End Sub
