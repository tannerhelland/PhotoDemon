VERSION 5.00
Begin VB.Form FormAntique 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   " Antique"
   ClientHeight    =   6510
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
   ScaleHeight     =   434
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   808
   Begin PhotoDemon.pdCommandBar cmdBar 
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5760
      Width           =   12120
      _ExtentX        =   21378
      _ExtentY        =   1323
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
   Begin PhotoDemon.pdSlider sldColor 
      Height          =   705
      Left            =   6000
      TabIndex        =   2
      Top             =   720
      Width           =   5925
      _ExtentX        =   10451
      _ExtentY        =   1270
      Caption         =   "color fade"
      Max             =   100
      Value           =   100
      NotchPosition   =   2
      NotchValueCustom=   100
   End
   Begin PhotoDemon.pdSlider sldSoftness 
      Height          =   705
      Left            =   6000
      TabIndex        =   3
      Top             =   1800
      Width           =   5925
      _ExtentX        =   10451
      _ExtentY        =   1270
      Caption         =   "diffuse light"
      Max             =   100
      Value           =   10
      NotchPosition   =   2
      NotchValueCustom=   10
   End
   Begin PhotoDemon.pdSlider sldGrain 
      Height          =   705
      Left            =   6000
      TabIndex        =   4
      Top             =   2880
      Width           =   5925
      _ExtentX        =   10451
      _ExtentY        =   1270
      Caption         =   "grain"
      Max             =   100
      Value           =   25
      NotchPosition   =   2
      NotchValueCustom=   25
   End
   Begin PhotoDemon.pdSlider sldVignette 
      Height          =   705
      Left            =   6000
      TabIndex        =   5
      Top             =   3960
      Width           =   5925
      _ExtentX        =   10451
      _ExtentY        =   1270
      Caption         =   "fade edges"
      Max             =   100
      Value           =   75
      NotchPosition   =   2
      NotchValueCustom=   75
   End
End
Attribute VB_Name = "FormAntique"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Antique Effect
'Copyright 2012-2026 by Tanner Helland
'Created: 03/July/12
'Last updated: 19/October/17
'Last update: rework algorithm with a full UI and user-controllable params
'
'PD's "Antique" effect is all-new for 7.0.  It wraps a number of different effects into a single UI,
' which should make it easier for users to achieve an old-timey look without resorting to long tutorials.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

Public Sub AntiqueEffect(ByVal effectParams As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)
    
    If (Not toPreview) Then Message "Sending image back in time..."
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString effectParams
    
    Dim antiqueStrength As Double, antiqueSoftness As Double, antiqueGrain As Double, antiqueVignette As Double
    
    With cParams
        antiqueStrength = cParams.GetDouble("color", 50#)
        antiqueSoftness = cParams.GetDouble("softness", 5#)
        antiqueGrain = cParams.GetDouble("grain", 5#)
        antiqueVignette = cParams.GetDouble("vignette", 75#)
    End With
    
    'To ensure that inputs of '0' return the base image, we gently fade-in the softness overlay
    ' as the radius increases.
    Dim fadeOpacity As Double
    If (antiqueSoftness >= 10#) Then fadeOpacity = 100# Else fadeOpacity = antiqueSoftness * 10#
    
    'Generate a workingDIB object and modify any relevant parameters by the preview size (if this is a preview)
    Dim tmpSA As SafeArray2D
    EffectPrep.PrepImageData tmpSA, toPreview, dstPic
    
    If toPreview Then antiqueSoftness = antiqueSoftness * curDIBValues.previewModifier
    
    'The actual antique-ification is handled elsewhere
    Filters_Stylize.ApplyAntiqueEffect workingDIB, antiqueStrength, antiqueSoftness, fadeOpacity, antiqueGrain, antiqueVignette, toPreview
    
    EffectPrep.FinalizeImageData toPreview, dstPic, True
    
End Sub

Private Sub cmdBar_OKClick()
    Process "Antique", , GetLocalParamString(), UNDO_Layer
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub Form_Load()
    cmdBar.SetPreviewStatus False
    ApplyThemeAndTranslations Me, True, True
    cmdBar.SetPreviewStatus True
    UpdatePreview
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub sldColor_Change()
    UpdatePreview
End Sub

Private Sub UpdatePreview()
    If cmdBar.PreviewsAllowed Then Me.AntiqueEffect GetLocalParamString(), True, pdFxPreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub pdFxPreview_ViewportChanged()
    UpdatePreview
End Sub

Private Sub sldGrain_Change()
    UpdatePreview
End Sub

Private Sub sldSoftness_Change()
    UpdatePreview
End Sub

Private Function GetLocalParamString() As String
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    
    With cParams
        .AddParam "color", sldColor.Value
        .AddParam "softness", sldSoftness.Value
        .AddParam "grain", sldGrain.Value
        .AddParam "vignette", sldVignette.Value
    End With
    
    GetLocalParamString = cParams.GetParamString()
    
End Function

Private Sub sldVignette_Change()
    UpdatePreview
End Sub
