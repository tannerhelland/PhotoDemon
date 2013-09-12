VERSION 5.00
Begin VB.Form FormBoxBlur 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Box Blur"
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
   StartUpPosition =   1  'CenterOwner
   Begin PhotoDemon.sliderTextCombo sltWidth 
      Height          =   495
      Left            =   6000
      TabIndex        =   6
      Top             =   2280
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   873
      Min             =   1
      Max             =   500
      Value           =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      Height          =   480
      Left            =   6120
      TabIndex        =   5
      Top             =   3840
      Width           =   2880
      _ExtentX        =   5080
      _ExtentY        =   847
      Caption         =   "keep both dimensions in sync"
      Value           =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PhotoDemon.sliderTextCombo sltHeight 
      Height          =   495
      Left            =   6000
      TabIndex        =   7
      Top             =   3240
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   873
      Min             =   1
      Max             =   500
      Value           =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblHeight 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "box height:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   6000
      TabIndex        =   4
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label lblWidth 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "box width:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   6000
      TabIndex        =   3
      Top             =   1920
      Width           =   1140
   End
   Begin VB.Label lblIDEWarning 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   975
      Left            =   6000
      TabIndex        =   2
      Top             =   4440
      Visible         =   0   'False
      Width           =   5775
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "FormBoxBlur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Box Blur Tool
'Copyright ©2000-2013 by Tanner Helland
'Created: some time 2000
'Last updated: 24/April/13
'Last update: greatly simplified code by using the new slider/text custom control
'
'This is a heavily optimized box blur.  An "accumulation" technique is used instead of the standard sliding
' window mechanism.  (See http://web.archive.org/web/20060718054020/http://www.acm.uiuc.edu/siggraph/workshops/wjarosz_convolution_2001.pdf)
' This allows the algorithm to perform extremely well, despite being written in pure VB.
'
'That said, it is still unfortunately slow in the IDE.  I STRONGLY recommend compiling the project before
' applying any box blur of a large radius.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://www.tannerhelland.com/photodemon/#license
'
'***************************************************************************

Option Explicit

'When previewing, we need to modify the strength to be representative of the final filter.  This means dividing by the
' original image dimensions in order to establish the right ratio.
Dim iWidth As Long, iHeight As Long

'Custom tooltip class allows for things like multiline, theming, and multiple monitor support
Dim m_ToolTip As clsToolTip

Private Sub chkUnison_Click()
    If CBool(chkUnison) Then syncScrollBars True
End Sub

'Convolve an image using a box blur.  An accumulation approach is used to maximize speed.
'Input: horizontal and vertical size of the box (I call it radius, because the final box size is 2r + 1)
Public Sub BoxBlurFilter(ByVal hRadius As Long, ByVal vRadius As Long, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)
    
    If toPreview = False Then Message "Applying box blur to image..."
        
    'Create a local array and point it at the pixel data of the current image
    Dim dstSA As SAFEARRAY2D
    prepImageData dstSA, toPreview, dstPic
    
    'Create a second local array.  This will contain the a copy of the current image, and we will use it as our source reference
    ' (This is necessary to prevent blurred pixel values from spreading across the image as we go.)
    Dim srcLayer As pdLayer
    Set srcLayer = New pdLayer
    srcLayer.createFromExistingLayer workingLayer
        
    'If this is a preview, we need to adjust the kernel radius to match the size of the preview box
    If toPreview Then
        hRadius = (hRadius / iWidth) * curLayerValues.Width
        vRadius = (vRadius / iHeight) * curLayerValues.Height
        If hRadius = 0 Then hRadius = 1
        If vRadius = 0 Then vRadius = 1
    End If
    
    'Apply the box blur
    CreateBoxBlurLayer hRadius, vRadius, srcLayer, workingLayer, toPreview
        
    srcLayer.eraseLayer
    Set srcLayer = Nothing
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering using the data inside workingLayer
    finalizeImageData toPreview, dstPic

End Sub

Private Sub cmdBar_OKClick()
    Process "Box blur", , buildParams(sltWidth, sltHeight)
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    updatePreview
End Sub

Private Sub Form_Activate()
    
    'Assign the system hand cursor to all relevant objects
    Set m_ToolTip = New clsToolTip
    makeFormPretty Me, m_ToolTip
    
    'If the program is not compiled, display a special warning for this tool
    If Not g_IsProgramCompiled Then
        lblIDEWarning.Caption = g_Language.TranslateMessage("WARNING! This tool is very slow when used inside the IDE. Please compile for best results.")
        lblIDEWarning.Visible = True
    End If
    
    'Draw a preview of the effect
    updatePreview
    
End Sub

Private Sub Form_Load()

    'Note the current image's width and height, which will be needed to adjust the preview effect
    If pdImages(CurrentImage).selectionActive Then
        iWidth = pdImages(CurrentImage).mainSelection.boundWidth
        iHeight = pdImages(CurrentImage).mainSelection.boundHeight
    Else
        iWidth = pdImages(CurrentImage).Width
        iHeight = pdImages(CurrentImage).Height
    End If
    
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
