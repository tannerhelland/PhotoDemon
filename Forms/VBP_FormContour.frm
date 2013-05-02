VERSION 5.00
Begin VB.Form FormContour 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Trace Contour"
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
   Begin PhotoDemon.smartCheckBox chkBlackBackground 
      Height          =   570
      Left            =   6120
      TabIndex        =   6
      Top             =   3120
      Width           =   2670
      _extentx        =   4710
      _extenty        =   1005
      caption         =   "use black background"
      font            =   "VBP_FormContour.frx":0000
      value           =   1
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
      TabIndex        =   3
      Top             =   120
      Width           =   5625
      _extentx        =   9922
      _extenty        =   9922
   End
   Begin PhotoDemon.smartCheckBox chkSmoothing 
      Height          =   570
      Left            =   6120
      TabIndex        =   7
      Top             =   3720
      Width           =   3030
      _extentx        =   5345
      _extenty        =   1005
      caption         =   "apply contour smoothing"
      font            =   "VBP_FormContour.frx":0028
      value           =   1
   End
   Begin PhotoDemon.sliderTextCombo sltThickness 
      Height          =   495
      Left            =   6000
      TabIndex        =   8
      Top             =   2400
      Width           =   5895
      _extentx        =   10398
      _extenty        =   873
      font            =   "VBP_FormContour.frx":0050
      min             =   1
      max             =   30
      value           =   1
   End
   Begin VB.Label lblThickness 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "thickness:"
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
      TabIndex        =   5
      Top             =   2040
      Width           =   1050
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
      Height          =   1095
      Left            =   6000
      TabIndex        =   4
      Top             =   4440
      Visible         =   0   'False
      Width           =   5775
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblBackground 
      Height          =   855
      Left            =   0
      TabIndex        =   2
      Top             =   5760
      Width           =   12135
   End
End
Attribute VB_Name = "FormContour"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Trace Contour (Outline) Tool
'Copyright ©2012-2013 by Tanner Helland
'Created: 15/Feb/13
'Last updated: 25/April/13
'Last update: simplified code by relying on new slider/text custom control
'
'Contour tracing is performed by "stacking" a series of filters together:
' 1) Gaussian blur to smooth out fine details
' 2) Median to unify colors and round out edges
' 3) Edge detection
' 4) Auto white balance (as the original edge detection function is quite dark)
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://www.tannerhelland.com/photodemon/#license
'
'***************************************************************************

Option Explicit

'When previewing, we need to modify the strength to be representative of the final filter.  This means dividing by the
' original image dimensions in order to establish the right ratio.
Dim iWidth As Long, iHeight As Long

Dim allowPreview As Boolean

Private Sub chkSmoothing_Click()
    updatePreview
End Sub

'CANCEL button
Private Sub CmdCancel_Click()
    Unload Me
End Sub

'OK button
Private Sub cmdOK_Click()

    'Validate text box entries
    If sltThickness.IsValid Then
        Me.Visible = False
        Process Contour, sltThickness, CBool(chkBlackBackground.Value), CBool(chkSmoothing.Value)
        Unload Me
    End If
        
End Sub

'Convolve an image using a gaussian kernel (separable implementation!)
'Input: radius of the contour (min 1, no real max - but the scroll bar is maxed at 200 presently)
Public Sub TraceContour(ByVal cRadius As Long, ByVal useBlackBackground As Boolean, ByVal useSmoothing As Boolean, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)
    
    If Not toPreview Then Message "Tracing image contour..."
            
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
            cRadius = (cRadius / iWidth) * curLayerValues.Width
        Else
            cRadius = (cRadius / iHeight) * curLayerValues.Height
        End If
        If cRadius = 0 Then cRadius = 1
    End If
    
    Dim finalX As Long, finalY As Long
    finalX = workingLayer.getLayerWidth
    finalY = workingLayer.getLayerHeight
        
    If useSmoothing Then
        'Blur the current layer
        CreateGaussianBlurLayer cRadius, srcLayer, workingLayer, toPreview, finalY * 2 + finalX * 3
        
        'Use the median filter to round out edges
        CreateMedianLayer cRadius, 50, workingLayer, srcLayer, toPreview, finalY * 2 + finalX * 3, finalY * 2
        
        'Next, create a contour of the layer
        CreateContourLayer useBlackBackground, srcLayer, workingLayer, toPreview, finalY * 2 + finalX * 3, finalY * 2 + finalX
            
        srcLayer.eraseLayer
        Set srcLayer = Nothing
        
        'Finally, white balance the resulting layer
        WhiteBalanceLayer 0.01, workingLayer, toPreview, finalY * 2 + finalX * 2, finalY * 2 + finalX * 2
    Else
        'Blur the current layer
        CreateGaussianBlurLayer cRadius, workingLayer, srcLayer, toPreview, finalY * 2 + finalX * 2
        
        'Next, create a contour of the layer
        CreateContourLayer useBlackBackground, srcLayer, workingLayer, toPreview, finalY * 2 + finalX * 2, finalY * 2
            
        srcLayer.eraseLayer
        Set srcLayer = Nothing
        
        'Finally, white balance the resulting layer
        WhiteBalanceLayer 0.01, workingLayer, toPreview, finalY * 2 + finalX * 2, finalY * 2 + finalX
    End If
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering using the data inside workingLayer
    finalizeImageData toPreview, dstPic

End Sub

Private Sub Form_Activate()

    'Note the current image's width and height, which will be needed to adjust the preview effect
    If pdImages(CurrentImage).selectionActive Then
        iWidth = pdImages(CurrentImage).mainSelection.selWidth
        iHeight = pdImages(CurrentImage).mainSelection.selHeight
    Else
        iWidth = pdImages(CurrentImage).Width
        iHeight = pdImages(CurrentImage).Height
    End If

    allowPreview = True

    'Draw a preview of the effect
    updatePreview
    
    'Assign the system hand cursor to all relevant objects
    makeFormPretty Me
    
    'If the program is not compiled, display a special warning for this tool
    If Not g_IsProgramCompiled Then
        lblIDEWarning.Caption = g_Language.TranslateMessage("WARNING!  This tool has been heavily optimized, but at high radius values it will still be quite slow inside the IDE.  Please compile before applying or previewing any radius larger than 20.")
        lblIDEWarning.Visible = True
    End If
    
End Sub

Private Sub Form_Load()
    allowPreview = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub chkBlackBackground_Click()
    updatePreview
End Sub

Private Sub sltThickness_Change()
    updatePreview
End Sub

Private Sub updatePreview()
    If allowPreview Then TraceContour sltThickness, CBool(chkBlackBackground.Value), CBool(chkSmoothing.Value), True, fxPreview
End Sub

