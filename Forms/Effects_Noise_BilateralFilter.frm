VERSION 5.00
Begin VB.Form FormBilateral 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Bilateral smoothing"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   285
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
   ScaleHeight     =   436
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   806
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.pdCommandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5790
      Width           =   12090
      _ExtentX        =   21325
      _ExtentY        =   1323
   End
   Begin PhotoDemon.pdSlider sltRadius 
      Height          =   705
      Left            =   6000
      TabIndex        =   2
      Top             =   1320
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
      Caption         =   "radius"
      Min             =   1
      Max             =   50
      Value           =   3
      DefaultValue    =   3
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
   Begin PhotoDemon.pdSlider sltSpatialFactor 
      Height          =   705
      Left            =   6000
      TabIndex        =   3
      Top             =   2280
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
      Caption         =   "spatial strength"
      Min             =   1
      Max             =   100
      SigDigits       =   1
      Value           =   10
      DefaultValue    =   10
   End
   Begin PhotoDemon.pdSlider sltColorFactor 
      Height          =   705
      Left            =   6000
      TabIndex        =   4
      Top             =   3240
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
      Caption         =   "color strength"
      Min             =   1
      Max             =   100
      SigDigits       =   1
      Value           =   50
      DefaultValue    =   50
   End
End
Attribute VB_Name = "FormBilateral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Bilateral Smoothing Form
'Copyright 2014-2019 by Tanner Helland
'Created: 19/June/14
'Last updated: 15/November/19
'Last update: rewrite from scratch for much-improved performance
'
'Per Wikipedia (https://en.wikipedia.org/wiki/Bilateral_filter):
' "A bilateral filter is a non-linear, edge-preserving, and noise-reducing smoothing filter for images.
' It replaces the intensity of each pixel with a weighted average of intensity values from nearby pixels.
' This weight can be based on a Gaussian distribution. Crucially, the weights depend not only on
' Euclidean distance of pixels, but also on the radiometric differences (e.g., range differences, such as
' color intensity, depth distance, etc.). This preserves sharp edges."
'
'More details on bilateral filtering can be found at:
' http://www.cs.duke.edu/~tomasi/papers/tomasi/tomasiIccv98.pdf
'
'Because traditional 2D kernel convolution is extremely slow on images of any size, PhotoDemon uses a
' separable bilateral filter implementation.  This provides a good approximation of a true bilateral,
' and it transforms the filter from an O(w*h*r^2) process to O(w*h*2r).
'
'For details on a separable bilateral approach, see:
' http://homepage.tudelft.nl/e3q6n/publications/2005/ICME2005_TPLV.pdf
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

Public Sub BilateralSmoothingSeparable(ByVal effectParams As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)
    
    If (Not toPreview) Then Message "Applying bilateral smoothing..."
    
    Dim cParams As pdParamXML
    Set cParams = New pdParamXML
    cParams.SetParamString effectParams
    
    Dim kernelRadius As Long, spatialFactor As Double, colorFactor As Double
    
    With cParams
        kernelRadius = .GetLong("radius", 1)
        spatialFactor = .GetDouble("spatialfactor", 10#)
        colorFactor = .GetDouble("colorfactor", 10#)
    End With
    
    'PrepImageData generates a working copy of the current filter target
    Dim dstSA As SafeArray2D
    EffectPrep.PrepImageData dstSA, toPreview, dstPic
    
    'If this is a preview, we need to adjust kernel size to match
    If toPreview Then kernelRadius = kernelRadius * curDIBValues.previewModifier
    
    'PD now provides a dedicated function for separable bilateral processing
    CreateBilateralDIB workingDIB, kernelRadius, spatialFactor, colorFactor, toPreview
    
    'Finalize result
    EffectPrep.FinalizeImageData toPreview, dstPic
    
End Sub

Private Sub cmdBar_OKClick()
    Process "Bilateral smoothing", , GetLocalParamString(), UNDO_Layer
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub Form_Load()
    cmdBar.MarkPreviewStatus False
    ApplyThemeAndTranslations Me
    cmdBar.MarkPreviewStatus True
    UpdatePreview
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub sltColorFactor_Change()
    UpdatePreview
End Sub

Private Sub sltRadius_Change()
    UpdatePreview
End Sub

Private Sub sltSpatialFactor_Change()
    UpdatePreview
End Sub

Private Sub UpdatePreview()
    If cmdBar.PreviewsAllowed Then BilateralSmoothingSeparable GetLocalParamString(), True, pdFxPreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub pdFxPreview_ViewportChanged()
    UpdatePreview
End Sub

Private Function GetLocalParamString() As String
    
    Dim cParams As pdParamXML
    Set cParams = New pdParamXML
    
    With cParams
        .AddParam "radius", sltRadius.Value
        .AddParam "spatialfactor", sltSpatialFactor.Value
        .AddParam "colorfactor", sltColorFactor.Value
    End With
    
    GetLocalParamString = cParams.GetParamString()
    
End Function
