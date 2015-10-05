VERSION 5.00
Begin VB.Form FormGaussianBlur 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Gaussian blur"
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
   Begin PhotoDemon.sliderTextCombo sltRadius 
      Height          =   720
      Left            =   6000
      TabIndex        =   2
      Top             =   1800
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
      Caption         =   "radius"
      Min             =   0.1
      Max             =   500
      SigDigits       =   1
      Value           =   5
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
   Begin PhotoDemon.buttonStrip btsQuality 
      Height          =   600
      Left            =   6000
      TabIndex        =   4
      Top             =   3180
      Width           =   5835
      _ExtentX        =   11774
      _ExtentY        =   1058
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "quality"
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
      Index           =   1
      Left            =   6000
      TabIndex        =   3
      Top             =   2730
      Width           =   705
   End
End
Attribute VB_Name = "FormGaussianBlur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Gaussian Blur Tool
'Copyright 2010-2015 by Tanner Helland
'Created: 01/July/10
'Last updated: 25/September/14
'Last update: switch quality option buttons to button strip
'
'To my knowledge, this tool is the first of its kind in VB6 - a variable radius gaussian blur filter
' that utilizes a separable convolution kernel AND allows for sub-pixel radii (at "best" quality, anyway).

'The use of separable kernels makes this much, much faster than a standard Gaussian blur.  The approximate
' speed gain for a P x Q kernel is PQ/(P + Q) - so for a radius of 4 (which is an actual kernel of 9x9)
' the processing time is 4.5x faster.  For a radius of 100, my technique is 100x faster than a traditional
' method.
'
'For an even faster blur, iterative box blurs can be applied.  The "good" quality uses a 3x box blur estimate, while
' the "better" quality uses a 5x.  The way the box blur radius is calculated also varies; 3x uses a quadratic
' estimation to try and improve output, while the 5x uses a quick-and-dirty estimation as the number of iterations
' makes a more elegant one pointless.
'
'Note that "good quality" is ~20x faster than "best quality", and "best quality" is 100x faster than a naive implementation.
' Pretty fast stuff!
'
'Despite this, it's still quite slow in the IDE due to the number of array accesses required.  I STRONGLY
' recommend compiling the project before applying any Gaussian blur of a large radius.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Convolve an image using a gaussian kernel (separable implementation!)
'Input: radius of the blur (min 1, no real max - but the scroll bar is maxed at 200 presently)
Public Sub GaussianBlurFilter(ByVal gRadius As Double, Optional ByVal gaussQuality As Long = 2, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)
        
    If Not toPreview Then Message "Applying gaussian blur..."
        
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
        gRadius = gRadius * curDIBValues.previewModifier
        If gRadius = 0 Then gRadius = 0.01
    End If
        
    'I almost always recommend quality over speed for PD tools, but in this case, the fast option is SO much faster,
    ' and the results so indistinguishable (3% different according to the Central Limit Theorem:
    ' https://www.khanacademy.org/math/probability/statistics-inferential/sampling_distribution/v/central-limit-theorem?playlist=Statistics
    ' ), that I recommend the faster methods instead.
    Select Case gaussQuality
    
        '3 iteration box blur
        Case 0
            CreateApproximateGaussianBlurDIB gRadius, srcDIB, workingDIB, 3, toPreview
        
        'IIR Gaussian estimation
        Case 1
            Filters_Area.GaussianBlur_IIRImplementation workingDIB, gRadius, 3, toPreview
            
        'True Gaussian
        Case 2
            CreateGaussianBlurDIB gRadius, srcDIB, workingDIB, toPreview
        
    End Select
    
    srcDIB.eraseDIB
    Set srcDIB = Nothing
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering using the data inside workingDIB
    finalizeImageData toPreview, dstPic, True
            
End Sub

Private Sub btsQuality_Click(ByVal buttonIndex As Long)
    updatePreview
End Sub

'OK button
Private Sub cmdBar_OKClick()
    Process "Gaussian blur", , buildParams(sltRadius, btsQuality.ListIndex), UNDO_LAYER
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    updatePreview
End Sub

Private Sub cmdBar_ResetClick()
    sltRadius.Value = 1
End Sub

Private Sub Form_Activate()
    
    'Apply translations and visual themes
    MakeFormPretty Me
        
    'Draw a preview of the effect
    updatePreview
    
End Sub

Private Sub Form_Load()
    
    'Populate the quality selector
    btsQuality.AddItem "good", 0
    btsQuality.AddItem "better", 1
    btsQuality.AddItem "best", 2
    btsQuality.ListIndex = 0
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub updatePreview()
    If cmdBar.previewsAllowed Then GaussianBlurFilter sltRadius.Value, btsQuality.ListIndex, True, fxPreview
End Sub

Private Sub sltRadius_Change()
    updatePreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub fxPreview_ViewportChanged()
    updatePreview
End Sub


