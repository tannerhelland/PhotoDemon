VERSION 5.00
Begin VB.Form FormColorHalftone 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   " Color halftone"
   ClientHeight    =   6510
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11670
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
   ScaleWidth      =   778
   Visible         =   0   'False
   Begin PhotoDemon.pdCommandBar cmdBar 
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5760
      Width           =   11670
      _ExtentX        =   20585
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
   Begin PhotoDemon.pdSlider sltAngle 
      Height          =   705
      Index           =   0
      Left            =   6000
      TabIndex        =   2
      Top             =   2520
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   1244
      Caption         =   "cyan angle"
      Max             =   360
      SigDigits       =   1
   End
   Begin PhotoDemon.pdSlider sltRadius 
      Height          =   705
      Left            =   6000
      TabIndex        =   3
      Top             =   600
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   1244
      Caption         =   "radius"
      Min             =   2
      Max             =   50
      SigDigits       =   1
      Value           =   5
      DefaultValue    =   5
   End
   Begin PhotoDemon.pdSlider sltAngle 
      Height          =   705
      Index           =   1
      Left            =   6000
      TabIndex        =   4
      Top             =   3480
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   1244
      Caption         =   "magenta angle"
      Max             =   360
      SigDigits       =   1
      DefaultValue    =   33.3
   End
   Begin PhotoDemon.pdSlider sltAngle 
      Height          =   705
      Index           =   2
      Left            =   6000
      TabIndex        =   5
      Top             =   4440
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   1244
      Caption         =   "yellow angle"
      Max             =   360
      SigDigits       =   1
      DefaultValue    =   66.7
   End
   Begin PhotoDemon.pdSlider sltDensity 
      Height          =   705
      Left            =   6000
      TabIndex        =   6
      Top             =   1560
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   1244
      Caption         =   "density"
      Max             =   100
      SigDigits       =   1
      Value           =   100
      NotchPosition   =   2
      NotchValueCustom=   100
   End
End
Attribute VB_Name = "FormColorHalftone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Color Halftone Effect Interface
'Copyright 2014-2026 by Tanner Helland
'Created: 01/April/15
'Last updated: 01/April/15
'Last update: initial build
'
'Color halftoning creates a magazine-like effect, using circles of varying size, varying angle, and density to
' recreate an image using a traditional CMYK print function.
'
'Thank you to Plinio Garcia for suggesting this effect to me.
'
'This tool's algorithm is a modified version of a function originally written by Jerry Huxtable of JH Labs.
' Jerry's original code is licensed under an Apache 2.0 license (http://www.apache.org/licenses/LICENSE-2.0).
' You may download his original version from the following link (good as of March '15):
' http://www.jhlabs.com/ip/filters/index.html
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'To improve preview performance, a copy of the working image is cached locally
Private m_EffectDIB As pdDIB

'Apply a CMYK halftone filter to the current image.
Public Sub ColorHalftoneFilter(ByVal effectParams As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)
    
    If (Not toPreview) Then Message "Printing image to digital halftone surface..."
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString effectParams
    
    Dim pxRadius As Double, cyanAngle As Double, magentaAngle As Double, yellowAngle As Double, dotDensity As Double
    
    With cParams
        pxRadius = .GetDouble("radius", sltRadius.Value)
        dotDensity = .GetDouble("density", sltDensity.Value)
        cyanAngle = .GetDouble("cyanangle", sltAngle(0).Value)
        magentaAngle = .GetDouble("magentaangle", sltAngle(1).Value)
        yellowAngle = .GetDouble("yellowangle", sltAngle(2).Value)
    End With
    
    'Create a local array and point it at the pixel data of the current image
    Dim dstSA As SafeArray2D
    EffectPrep.PrepImageData dstSA, toPreview, dstPic
    
    'Create a second local array.  This will contain the a copy of the current image, and we will use it as our source reference
    ' (This is necessary to prevent converted pixel values from spreading across the image as we go.)
    If (m_EffectDIB Is Nothing) Then Set m_EffectDIB = New pdDIB
    m_EffectDIB.CreateFromExistingDIB workingDIB
    
    'Modify the radius value for previews
    If toPreview Then pxRadius = pxRadius * curDIBValues.previewModifier
    If (pxRadius < 2#) Then pxRadius = 2#
    
    'Use the external function to apply the actual effect
    Filters_Stylize.CreateColorHalftoneDIB pxRadius, cyanAngle, magentaAngle, yellowAngle, dotDensity, m_EffectDIB, workingDIB, toPreview
    
    'If this is *not* a preview, we won't need our local image copy any more
    If (Not toPreview) Then Set m_EffectDIB = Nothing
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    EffectPrep.FinalizeImageData toPreview, dstPic
    
End Sub

Private Sub cmdBar_OKClick()
    Process "Color halftone", , GetLocalParamString(), UNDO_Layer
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub Form_Load()

    'Suspend previews while we initialize controls
    cmdBar.SetPreviewStatus False
        
    'Apply translations and themes
    ApplyThemeAndTranslations Me, True, True
    
    'Request a preview
    cmdBar.SetPreviewStatus True
    UpdatePreview
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'Redraw the effect preview
Private Sub UpdatePreview()
    If cmdBar.PreviewsAllowed Then ColorHalftoneFilter GetLocalParamString(), True, pdFxPreview
End Sub

Private Sub pdFxPreview_ViewportChanged()
    UpdatePreview
End Sub

Private Sub sltAngle_Change(Index As Integer)
    UpdatePreview
End Sub

Private Sub sltDensity_Change()
    UpdatePreview
End Sub

Private Sub sltRadius_Change()
    UpdatePreview
End Sub

Private Function GetLocalParamString() As String
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    
    With cParams
        .AddParam "radius", sltRadius.Value
        .AddParam "density", sltDensity.Value
        .AddParam "cyanangle", sltAngle(0).Value
        .AddParam "magentaangle", sltAngle(1).Value
        .AddParam "yellowangle", sltAngle(2).Value
    End With
    
    GetLocalParamString = cParams.GetParamString()
    
End Function
