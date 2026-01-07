VERSION 5.00
Begin VB.Form FormRotateDistort 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   " Rotate"
   ClientHeight    =   6510
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
   ScaleHeight     =   434
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   806
   Visible         =   0   'False
   Begin PhotoDemon.pdButtonStrip btsRender 
      Height          =   975
      Left            =   6000
      TabIndex        =   6
      Top             =   3600
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1720
      Caption         =   "mode"
   End
   Begin PhotoDemon.pdCommandBar cmdBar 
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5760
      Width           =   12090
      _ExtentX        =   21325
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
      DisableZoomPan  =   -1  'True
      PointSelection  =   -1  'True
   End
   Begin PhotoDemon.pdSlider sltAngle 
      Height          =   705
      Left            =   6000
      TabIndex        =   4
      Top             =   1680
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
      Caption         =   "angle"
      Min             =   -360
      Max             =   360
      SigDigits       =   1
   End
   Begin PhotoDemon.pdSlider sltXCenter 
      Height          =   405
      Left            =   6000
      TabIndex        =   2
      Top             =   600
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   873
      Max             =   1
      SigDigits       =   2
      Value           =   0.5
      NotchPosition   =   2
      NotchValueCustom=   0.5
   End
   Begin PhotoDemon.pdSlider sltYCenter 
      Height          =   405
      Left            =   9000
      TabIndex        =   3
      Top             =   600
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   873
      Max             =   1
      SigDigits       =   2
      Value           =   0.5
      NotchPosition   =   2
      NotchValueCustom=   0.5
   End
   Begin PhotoDemon.pdDropDown cboEdges 
      Height          =   735
      Left            =   6000
      TabIndex        =   5
      Top             =   2640
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1296
      Caption         =   "if pixels lie outside the image..."
   End
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   285
      Index           =   0
      Left            =   6000
      Top             =   240
      Width           =   5925
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "center position (x, y)"
      FontSize        =   12
      ForeColor       =   4210752
   End
   Begin PhotoDemon.pdLabel lblExplanation 
      Height          =   435
      Index           =   0
      Left            =   6120
      Top             =   1170
      Width           =   5655
      _ExtentX        =   0
      _ExtentY        =   0
      Alignment       =   2
      Caption         =   "you can also set a center position by clicking the preview window"
      FontSize        =   9
      ForeColor       =   4210752
      Layout          =   1
   End
   Begin PhotoDemon.pdLabel lblExplanation 
      Height          =   885
      Index           =   1
      Left            =   6000
      Top             =   4800
      Width           =   5925
      _ExtentX        =   0
      _ExtentY        =   0
      Alignment       =   2
      Caption         =   ""
      ForeColor       =   4210752
      Layout          =   1
   End
End
Attribute VB_Name = "FormRotateDistort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Rotate Distort Effect Interface (separate from image rotation for a reason - see below)
'Copyright 2013-2026 by Tanner Helland
'Created: 22/August/13
'Last updated: 28/July/17
'Last update: performance improvements, migrate to XML params
'
'Dialog for handling rotation via PhotoDemon's distort filter engine.  This is kept separate from full-image rotation,
' because I needed a rotate that could be applied to selections.  Also, full-image rotation allows you to resize the
' canvas.  This does not.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

Private Sub btsRender_Click(ByVal buttonIndex As Long)
    UpdatePreview
End Sub

Private Sub cboEdges_Click()
    UpdatePreview
End Sub

'Apply a basic rotation to the image or selected area
Public Sub RotateFilter(ByVal effectParams As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)
    
    If (Not toPreview) Then Message "Rotating area..."
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString effectParams
    
    Dim rotateAngle As Double, centerX As Double, centerY As Double
    Dim edgeHandling As Long, useBilinear As Boolean
    
    With cParams
        rotateAngle = .GetDouble("angle", sltAngle.Value)
        edgeHandling = .GetLong("edges", cboEdges.ListIndex)
        useBilinear = .GetBool("bilinear", False)
        centerX = .GetDouble("centerx", 0.5)
        centerY = .GetDouble("centery", 0.5)
    End With
    
    'Create a local array and point it at the pixel data of the current image
    Dim dstSA As SafeArray2D
    EffectPrep.PrepImageData dstSA, toPreview, dstPic
    
    'Create a second local array.  This will contain the a copy of the current image, and we will use it as our source reference
    ' (This is necessary to prevent converted pixel values from spreading across the image as we go.)
    Dim srcDIB As pdDIB
    Set srcDIB = New pdDIB
    srcDIB.CreateFromExistingDIB workingDIB
    
    'Use the external function to create a rotated DIB
    CreateRotatedDIB rotateAngle, edgeHandling, useBilinear, srcDIB, workingDIB, centerX, centerY, toPreview
    
    Set srcDIB = Nothing
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    EffectPrep.FinalizeImageData toPreview, dstPic
    
End Sub

Private Sub cmdBar_OKClick()
    Process "Rotate", , GetLocalParamString(), UNDO_Layer
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub cmdBar_ResetClick()
    sltXCenter.Value = 0.5
    sltYCenter.Value = 0.5
    cboEdges.ListIndex = pdeo_Wrap
End Sub

Private Sub Form_Load()

    'Suspend previews while we initialize all the controls
    cmdBar.SetPreviewStatus False
    
    btsRender.AddItem "fast", 0
    btsRender.AddItem "precise", 1
    btsRender.ListIndex = 1
    
    'I use a central function to populate the edge handling combo box; this way, I can add new methods and have
    ' them immediately available to all distort functions.
    PopDistortEdgeBox cboEdges, pdeo_Wrap
    
    'Provide an explanation on why this tool doesn't enlarge the canvas to match
    lblExplanation(1).Caption = g_Language.TranslateMessage("If you want to enlarge the canvas to fit the rotated image, please use the Image -> Rotate menu instead.")
    
    'Apply translations and visual themes
    ApplyThemeAndTranslations Me, True, True
    cmdBar.SetPreviewStatus True
    UpdatePreview
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'Redraw the effect preview
Private Sub UpdatePreview()
    If cmdBar.PreviewsAllowed Then RotateFilter GetLocalParamString(), True, pdFxPreview
End Sub

'The user can right-click the preview area to select a new center point
Private Sub pdFxPreview_PointSelected(xRatio As Double, yRatio As Double)
    
    cmdBar.SetPreviewStatus False
    sltXCenter.Value = xRatio
    sltYCenter.Value = yRatio
    cmdBar.SetPreviewStatus True
    UpdatePreview

End Sub

Private Sub sltAngle_Change()
    UpdatePreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub pdFxPreview_ViewportChanged()
    UpdatePreview
End Sub

Private Sub sltXCenter_Change()
    UpdatePreview
End Sub

Private Sub sltYCenter_Change()
    UpdatePreview
End Sub

Private Function GetLocalParamString() As String
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    
    With cParams
        .AddParam "angle", sltAngle.Value
        .AddParam "edges", cboEdges.ListIndex
        .AddParam "bilinear", (btsRender.ListIndex = 1)
        .AddParam "centerx", sltXCenter.Value
        .AddParam "centery", sltYCenter.Value
    End With
    
    GetLocalParamString = cParams.GetParamString()
    
End Function
