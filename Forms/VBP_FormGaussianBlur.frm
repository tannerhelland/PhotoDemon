VERSION 5.00
Begin VB.Form FormGaussianBlur 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Gaussian Blur"
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
   Begin PhotoDemon.sliderTextCombo sltRadius 
      Height          =   495
      Left            =   6000
      TabIndex        =   6
      Top             =   2760
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   873
      Min             =   0.1
      Max             =   200
      SigDigits       =   1
      Value           =   5
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
   Begin VB.CommandButton CmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   9030
      TabIndex        =   0
      Top             =   5910
      Width           =   1365
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   10500
      TabIndex        =   1
      Top             =   5910
      Width           =   1365
   End
   Begin PhotoDemon.fxPreviewCtl fxPreview 
      Height          =   5625
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9922
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
      Height          =   1215
      Left            =   6000
      TabIndex        =   5
      Top             =   4440
      Visible         =   0   'False
      Width           =   5775
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblBackground 
      Height          =   855
      Left            =   0
      TabIndex        =   3
      Top             =   5760
      Width           =   12135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "blur radius:"
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
      TabIndex        =   2
      Top             =   2400
      Width           =   1230
   End
End
Attribute VB_Name = "FormGaussianBlur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Gaussian Blur Tool
'Copyright ©2010-2013 by Tanner Helland
'Created: 01/July/10
'Last updated: 24/April/13
'Last update: simplified code by converting interface to new scroll/text custom control
'
'To my knowledge, this tool is the first of its kind in VB6 - a variable radius gaussian blur filter
' that utilizes a separable convolution kernel AND allows for sub-pixel radii.

'The use of separable kernels makes this much, much faster than a standard Gaussian blur.  The approimate
' speed gain for a P x Q kernel is PQ/(P + Q) - so for a radius of 4 (which is an actual kernel of 9x9)
' the processing time is 4.5x faster.  For a radius of 100, my technique is 100x faster than a traditional
' method.
'
'Despite this, it's still quite slow in the IDE due to the number of array accesses required.  I STRONGLY
' recommend compiling the project before applying any Gaussian blur of a large radius.
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

'CANCEL button
Private Sub CmdCancel_Click()
    Unload Me
End Sub

'OK button
Private Sub cmdOK_Click()

    If sltRadius.IsValid Then
        Me.Visible = False
        Process GaussianBlur, sltRadius
        Unload Me
    End If
    
End Sub

'Convolve an image using a gaussian kernel (separable implementation!)
'Input: radius of the blur (min 1, no real max - but the scroll bar is maxed at 200 presently)
Public Sub GaussianBlurFilter(ByVal gRadius As Double, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)
        
    If Not toPreview Then Message "Applying gaussian blur..."
        
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
        If iWidth > iHeight Then
            gRadius = (gRadius / iWidth) * curLayerValues.Width
        Else
            gRadius = (gRadius / iHeight) * curLayerValues.Height
        End If
        If gRadius = 0 Then gRadius = 1
    End If
    
    CreateGaussianBlurLayer gRadius, srcLayer, workingLayer, toPreview
    
    srcLayer.eraseLayer
    Set srcLayer = Nothing
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering using the data inside workingLayer
    finalizeImageData toPreview, dstPic
            
End Sub

Private Sub Form_Activate()

    'Note the current image's width and height, which will be needed to adjust the preview effect
    If pdImages(CurrentImage).selectionActive Then
        iWidth = pdImages(CurrentImage).mainSelection.boundWidth
        iHeight = pdImages(CurrentImage).mainSelection.boundHeight
    Else
        iWidth = pdImages(CurrentImage).Width
        iHeight = pdImages(CurrentImage).Height
    End If

    'Draw a preview of the effect
    updatePreview
    
    'Assign the system hand cursor to all relevant objects
    makeFormPretty Me, m_ToolTip
    
    'If the program is not compiled, display a special warning for this tool
    If Not g_IsProgramCompiled Then
        sltRadius.Max = 50
        lblIDEWarning.Caption = g_Language.TranslateMessage("WARNING!  This tool has been heavily optimized, but at high radius values it will still be quite slow inside the IDE.  Please compile before applying or previewing any radius larger than 20.")
        lblIDEWarning.Visible = True
    Else
        '32bpp images take quite a bit longer to process.  Limit the radius to 100 in this case.
        If pdImages(CurrentImage).mainLayer.getLayerColorDepth = 32 Then sltRadius.Max = 100 Else sltRadius.Max = 200
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub updatePreview()
    GaussianBlurFilter sltRadius.Value, True, fxPreview
End Sub

Private Sub sltRadius_Change()
    updatePreview
End Sub
