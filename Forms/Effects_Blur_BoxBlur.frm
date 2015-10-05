VERSION 5.00
Begin VB.Form FormBoxBlur 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Box blur"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   12030
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
   ScaleWidth      =   802
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.sliderTextCombo sltWidth 
      Height          =   720
      Left            =   6000
      TabIndex        =   3
      Top             =   2040
      Width           =   5880
      _ExtentX        =   10372
      _ExtentY        =   1270
      Caption         =   "box width"
      Min             =   1
      Max             =   500
      Value           =   2
   End
   Begin PhotoDemon.fxPreviewCtl fxPreview 
      Height          =   5625
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9922
   End
   Begin PhotoDemon.smartCheckBox chkUnison 
      Height          =   330
      Left            =   6120
      TabIndex        =   2
      Top             =   3840
      Width           =   5700
      _ExtentX        =   10054
      _ExtentY        =   582
      Caption         =   "keep both dimensions in sync"
   End
   Begin PhotoDemon.sliderTextCombo sltHeight 
      Height          =   720
      Left            =   6000
      TabIndex        =   4
      Top             =   3000
      Width           =   5880
      _ExtentX        =   10372
      _ExtentY        =   1270
      Caption         =   "box height"
      Min             =   1
      Max             =   500
      Value           =   2
   End
   Begin PhotoDemon.commandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5790
      Width           =   12030
      _ExtentX        =   21220
      _ExtentY        =   1323
      BackColor       =   14802140
   End
End
Attribute VB_Name = "FormBoxBlur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Box Blur Tool
'Copyright 2000-2015 by Tanner Helland
'Created: some time 2000
'Last updated: 17/September/13
'Last update: replace on the old accumulation technique with separable horizontal and vertical blurs.  Hot damn,
'              this thing is fast!
'
'This is a heavily optimized box blur.  Separable horizontal and vertical blurs are used, instead of the standard sliding
' window mechanism.  (See http://web.archive.org/web/20060718054020/http://www.acm.uiuc.edu/siggraph/workshops/wjarosz_convolution_2001.pdf)
' This allows the algorithm to perform extremely well, despite being written in pure VB.
'
'That said, it is still unfortunately slow in the IDE.  I STRONGLY recommend compiling the project before
' applying any box blur of a large radius.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

Private Sub chkUnison_Click()
    If CBool(chkUnison) Then syncScrollBars True
End Sub

'Convolve an image using a box blur.  An accumulation approach is used to maximize speed.
'Input: horizontal and vertical size of the box (I call it radius, because the final box size is 2r + 1)
Public Sub BoxBlurFilter(ByVal hRadius As Long, ByVal vRadius As Long, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)
    
    If Not toPreview Then Message "Applying box blur to image..."
        
    'Create a local array and point it at the pixel data of the current image
    Dim dstSA As SAFEARRAY2D
    prepImageData dstSA, toPreview, dstPic, , , True
    
    'Create a second local array.  This will contain the a copy of the current image, and we will use it as our source reference
    ' (This is necessary to prevent blurred pixel values from spreading across the image as we go.)
    Dim srcDIB As pdDIB
    Set srcDIB = New pdDIB
    srcDIB.createFromExistingDIB workingDIB
        
    'If this is a preview, we need to adjust the kernel radius to match the size of the preview box
    If toPreview Then
        hRadius = hRadius * curDIBValues.previewModifier
        vRadius = vRadius * curDIBValues.previewModifier
        If hRadius = 0 Then hRadius = 1
        If vRadius = 0 Then vRadius = 1
    End If
    
    'Apply the box blur in two steps: a fast horizontal blur, then a fast vertical blur
    CreateHorizontalBlurDIB hRadius, hRadius, workingDIB, srcDIB, toPreview, workingDIB.getDIBWidth + workingDIB.getDIBHeight
    CreateVerticalBlurDIB vRadius, vRadius, srcDIB, workingDIB, toPreview, workingDIB.getDIBWidth + workingDIB.getDIBHeight, workingDIB.getDIBWidth
    
    srcDIB.eraseDIB
    Set srcDIB = Nothing
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering using the data inside workingDIB
    finalizeImageData toPreview, dstPic, True

End Sub

Private Sub cmdBar_OKClick()
    Process "Box blur", , buildParams(sltWidth, sltHeight), UNDO_LAYER
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    updatePreview
End Sub

Private Sub Form_Activate()
    
    'Apply translations and visual themes
    MakeFormPretty Me
    
    'Draw a preview of the effect
    updatePreview
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'Keep the two scroll bars in sync.  Some extra work has to be done to makes sure scrollbar max values aren't exceeded.
Private Sub syncScrollBars(ByVal srcHorizontal As Boolean)
    
    If sltWidth.Value = sltHeight.Value Then Exit Sub
    
    Dim tmpVal As Long
    
    If srcHorizontal Then
        tmpVal = sltWidth.Value
        If tmpVal < sltHeight.Max Then sltHeight.Value = sltWidth.Value Else sltHeight.Value = sltHeight.Max
    Else
        tmpVal = sltHeight.Value
        If tmpVal < sltWidth.Max Then sltWidth.Value = sltHeight.Value Else sltWidth.Value = sltWidth.Max
    End If
    
End Sub
Private Sub updatePreview()
    If cmdBar.previewsAllowed Then BoxBlurFilter sltWidth, sltHeight, True, fxPreview
End Sub

Private Sub sltHeight_Change()
    If CBool(chkUnison) Then syncScrollBars False
    updatePreview
End Sub

Private Sub sltWidth_Change()
    If CBool(chkUnison) Then syncScrollBars True
    updatePreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub fxPreview_ViewportChanged()
    updatePreview
End Sub
