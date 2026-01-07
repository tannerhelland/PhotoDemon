VERSION 5.00
Begin VB.Form FormBoxBlur 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   " Box blur"
   ClientHeight    =   6540
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12030
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
   ScaleWidth      =   802
   Begin PhotoDemon.pdSlider sltWidth 
      Height          =   705
      Left            =   6000
      TabIndex        =   3
      Top             =   2040
      Width           =   5880
      _ExtentX        =   10372
      _ExtentY        =   1270
      Caption         =   "box width"
      Min             =   1
      Max             =   500
      ScaleStyle      =   1
      Value           =   2
      DefaultValue    =   2
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
   Begin PhotoDemon.pdCheckBox chkUnison 
      Height          =   330
      Left            =   6120
      TabIndex        =   2
      Top             =   3960
      Width           =   5700
      _ExtentX        =   10054
      _ExtentY        =   582
      Caption         =   "keep both dimensions in sync"
   End
   Begin PhotoDemon.pdSlider sltHeight 
      Height          =   705
      Left            =   6000
      TabIndex        =   4
      Top             =   3000
      Width           =   5880
      _ExtentX        =   10372
      _ExtentY        =   1270
      Caption         =   "box height"
      Min             =   1
      Max             =   500
      ScaleStyle      =   1
      Value           =   2
      DefaultValue    =   2
   End
   Begin PhotoDemon.pdCommandBar cmdBar 
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5790
      Width           =   12030
      _ExtentX        =   21220
      _ExtentY        =   1323
   End
End
Attribute VB_Name = "FormBoxBlur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Box Blur Tool
'Copyright 2000-2026 by Tanner Helland
'Created: some time 2000
'Last updated: 27/July/17
'Last update: performance improvements, migrate to XML params
'
'This is a heavily optimized box blur.  Separable horizontal and vertical blurs are used, instead of the standard sliding
' window mechanism.  (See http://web.archive.org/web/20060718054020/http://www.acm.uiuc.edu/siggraph/workshops/wjarosz_convolution_2001.pdf)
' This allows the algorithm to perform extremely well, despite being written in pure VB.
'
'That said, it is still unfortunately slow in the IDE.  I STRONGLY recommend compiling the project before
' applying any box blur of a large radius.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

Private Sub chkUnison_Click()
    If chkUnison.Value Then syncScrollBars True
End Sub

'Convolve an image using a box blur.  An accumulation approach is used to maximize speed.
'Input: horizontal and vertical size of the box (I call it radius, because the final box size is 2r + 1)
Public Sub BoxBlurFilter(ByVal effectParams As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)
    
    If (Not toPreview) Then Message "Applying box blur to image..."
    
    'Parse out specific parameters
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString effectParams
    
    Dim hRadius As Long, vRadius As Long
    hRadius = cParams.GetLong("radius-x", sltWidth.Value)
    vRadius = cParams.GetLong("radius-y", sltHeight.Value)
    
    'Create a local array and point it at the pixel data of the current image.  (Note that we deliberately
    ' leave alpha byte premultiplied!)
    Dim dstSA As SafeArray2D
    EffectPrep.PrepImageData dstSA, toPreview, dstPic, , , True
    
    'Create a second local array.  This will contain the a copy of the current image, and we will use it as our source reference
    ' (This is necessary to prevent blurred pixel values from spreading across the image as we go.)
    Dim srcDIB As pdDIB
    Set srcDIB = New pdDIB
    srcDIB.CreateFromExistingDIB workingDIB
        
    'If this is a preview, we need to adjust the kernel radius to match the size of the preview box
    If toPreview Then
        hRadius = hRadius * curDIBValues.previewModifier
        vRadius = vRadius * curDIBValues.previewModifier
        If (hRadius = 0) Then hRadius = 1
        If (vRadius = 0) Then vRadius = 1
    End If
    
    'Apply the box blur in two steps: a fast horizontal blur, then a fast vertical blur
    CreateHorizontalBlurDIB hRadius, hRadius, workingDIB, srcDIB, toPreview, workingDIB.GetDIBHeight * 2
    CreateVerticalBlurDIB vRadius, vRadius, srcDIB, workingDIB, toPreview, workingDIB.GetDIBHeight * 2, workingDIB.GetDIBHeight
    
    srcDIB.EraseDIB
    Set srcDIB = Nothing
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering using the data inside workingDIB
    EffectPrep.FinalizeImageData toPreview, dstPic, True

End Sub

Private Sub cmdBar_OKClick()
    Process "Box blur", , GetLocalParamString(), UNDO_Layer
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

'Keep the two scroll bars in sync.  Some extra work has to be done to makes sure scrollbar max values aren't exceeded.
Private Sub syncScrollBars(ByVal srcHorizontal As Boolean)
    
    If (sltWidth.Value = sltHeight.Value) Then Exit Sub
    
    Dim tmpVal As Long
    
    If srcHorizontal Then
        tmpVal = sltWidth.Value
        If (tmpVal < sltHeight.Max) Then sltHeight.Value = sltWidth.Value Else sltHeight.Value = sltHeight.Max
    Else
        tmpVal = sltHeight.Value
        If (tmpVal < sltWidth.Max) Then sltWidth.Value = sltHeight.Value Else sltWidth.Value = sltWidth.Max
    End If
    
End Sub

Private Sub UpdatePreview()
    If cmdBar.PreviewsAllowed Then BoxBlurFilter GetLocalParamString(), True, pdFxPreview
End Sub

Private Sub sltHeight_Change()
    If chkUnison.Value Then syncScrollBars False
    UpdatePreview
End Sub

Private Sub sltWidth_Change()
    If chkUnison.Value Then syncScrollBars True
    UpdatePreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub pdFxPreview_ViewportChanged()
    UpdatePreview
End Sub

Private Function GetLocalParamString() As String
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.AddParam "radius-x", sltWidth.Value
    cParams.AddParam "radius-y", sltHeight.Value
    GetLocalParamString = cParams.GetParamString()
End Function
