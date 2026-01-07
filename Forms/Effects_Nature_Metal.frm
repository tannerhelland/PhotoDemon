VERSION 5.00
Begin VB.Form FormMetal 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   " Metal"
   ClientHeight    =   6540
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11550
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
   ScaleWidth      =   770
   Begin PhotoDemon.pdCommandBar cmdBar 
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5790
      Width           =   11550
      _ExtentX        =   20373
      _ExtentY        =   1323
   End
   Begin PhotoDemon.pdSlider sltRadius 
      Height          =   705
      Left            =   6000
      TabIndex        =   2
      Top             =   1680
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   1244
      Caption         =   "smoothness"
      Max             =   200
      SigDigits       =   1
      Value           =   20
      DefaultValue    =   20
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
   Begin PhotoDemon.pdSlider sltDetail 
      Height          =   705
      Left            =   6000
      TabIndex        =   3
      Top             =   600
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   1244
      Caption         =   "detail"
      Max             =   16
      Value           =   4
      NotchPosition   =   2
      NotchValueCustom=   4
   End
   Begin PhotoDemon.pdColorSelector csHighlight 
      Height          =   975
      Left            =   6000
      TabIndex        =   4
      Top             =   2760
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   1720
      Caption         =   "highlight color"
      curColor        =   14737632
   End
   Begin PhotoDemon.pdColorSelector csShadow 
      Height          =   975
      Left            =   6000
      TabIndex        =   5
      Top             =   3960
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   1720
      Caption         =   "shadow color"
      curColor        =   4210752
   End
End
Attribute VB_Name = "FormMetal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'"Metal" or "Chrome" Image effect
'Copyright 2002-2026 by Tanner Helland
'Created: sometime 2002
'Last updated: 16/October/17
'Last update: migrate the actual function code elsewhere; it's helpful in other filter scenarios
'
'PhotoDemon's "Metal" filter is the rough equivalent of "Chrome" in Photoshop.  Our implementation is relatively
' straightforward; a normalized graymap is created for the image, then remapped according to a sinusoidal-like
' lookup table (created using the pdFilterLUT class).
'
'The user currently has control over two parameters: "smoothness", which determines a pre-effect blur radius,
' and "detail" which controls the number of octaves in the lookup table.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Apply a metallic "shimmer" to an image
Public Sub ApplyMetalFilter(ByVal effectParams As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)
    
    If (Not toPreview) Then Message "Pouring smoldering metal onto image..."
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString effectParams
    
    Dim steelDetail As Long, steelSmoothness As Double
    Dim shadowColor As Long, highlightColor As Long
    
    With cParams
        steelDetail = .GetLong("detail", sltDetail.Value)
        steelSmoothness = .GetDouble("radius", sltRadius.Value)
        shadowColor = .GetLong("shadowcolor", csShadow.Color)
        highlightColor = .GetLong("highlightcolor", csHighlight.Color)
    End With
    
    'Create a local array and point it at the pixel data of the current image
    Dim dstSA As SafeArray2D
    EffectPrep.PrepImageData dstSA, toPreview, dstPic
    
    'If this is a preview, we need to adjust the smoothness (kernel radius) to match the size of the preview box
    If toPreview Then steelSmoothness = steelSmoothness * curDIBValues.previewModifier
    
    'The actual chrome filter lives elsewhere
    Filters_Natural.GetChromeDIB workingDIB, steelDetail, steelSmoothness, shadowColor, highlightColor, toPreview
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering using the data inside workingDIB
    EffectPrep.FinalizeImageData toPreview, dstPic
            
End Sub

'OK button
Private Sub cmdBar_OKClick()
    Process "Metal", , GetLocalParamString(), UNDO_Layer
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub cmdBar_ResetClick()
    sltRadius.Value = 20
    sltDetail.Value = 4
    csShadow.Color = RGB(30, 30, 30)
    csHighlight.Color = RGB(230, 230, 230)
End Sub

Private Sub csHighlight_ColorChanged()
    UpdatePreview
End Sub

Private Sub csShadow_ColorChanged()
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

Private Sub UpdatePreview()
    If cmdBar.PreviewsAllowed Then Me.ApplyMetalFilter GetLocalParamString(), True, pdFxPreview
End Sub

Private Sub sltDetail_Change()
    UpdatePreview
End Sub

Private Sub sltRadius_Change()
    UpdatePreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub pdFxPreview_ViewportChanged()
    UpdatePreview
End Sub

Private Function GetLocalParamString() As String
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    
    With cParams
        .AddParam "detail", sltDetail.Value
        .AddParam "radius", sltRadius.Value
        .AddParam "shadowcolor", csShadow.Color
        .AddParam "highlightcolor", csHighlight.Color
    End With
    
    GetLocalParamString = cParams.GetParamString()
    
End Function
