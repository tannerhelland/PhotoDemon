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
   Begin PhotoDemon.smartCheckBox chkBlackBackground 
      Height          =   570
      Left            =   6120
      TabIndex        =   4
      Top             =   3120
      Width           =   2670
      _ExtentX        =   4710
      _ExtentY        =   1005
      Caption         =   "use black background"
      Value           =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
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
   Begin PhotoDemon.smartCheckBox chkSmoothing 
      Height          =   570
      Left            =   6120
      TabIndex        =   5
      Top             =   3720
      Width           =   3030
      _ExtentX        =   5345
      _ExtentY        =   1005
      Caption         =   "apply contour smoothing"
      Value           =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PhotoDemon.sliderTextCombo sltThickness 
      Height          =   495
      Left            =   6000
      TabIndex        =   6
      Top             =   2400
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   873
      Min             =   1
      Max             =   30
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
      TabIndex        =   3
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
      TabIndex        =   2
      Top             =   4440
      Visible         =   0   'False
      Width           =   5775
      WordWrap        =   -1  'True
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
'Last updated: 10/May/13
'Last update: allow the user to cancel at any time by pressing ESC
'
'Contour tracing is performed by "stacking" a series of filters together:
' 1) Gaussian blur to smooth out fine details
' 2) Median to unify colors and round out edges
' 3) Edge detection
' 4) Auto white balance (as the original edge detection function is quite dark)
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Custom tooltip class allows for things like multiline, theming, and multiple monitor support
Dim m_ToolTip As clsToolTip

Private Sub chkSmoothing_Click()
    updatePreview
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
        cRadius = cRadius * curLayerValues.previewModifier
        If cRadius = 0 Then cRadius = 1
    End If
    
    Dim finalX As Long, finalY As Long
    finalX = workingLayer.getLayerWidth
    finalY = workingLayer.getLayerHeight
        
    If useSmoothing Then
    
        'Blur the current layer
        If CreateGaussianBlurLayer(cRadius, srcLayer, workingLayer, toPreview, finalY * 2 + finalX * 3) Then
        
            'Use the median filter to round out edges
            If CreateMedianLayer(cRadius, 50, workingLayer, srcLayer, toPreview, finalY * 2 + finalX * 3, finalY * 2) Then
        
                'Next, create a contour of the layer
                If CreateContourLayer(useBlackBackground, srcLayer, workingLayer, toPreview, finalY * 2 + finalX * 3, finalY * 2 + finalX) Then
            
                    'Finally, white balance the resulting layer
                    WhiteBalanceLayer 0.01, workingLayer, toPreview, finalY * 2 + finalX * 2, finalY * 2 + finalX * 2
                    
                End If
            End If
        End If
    Else
        
        'Blur the current layer
        If CreateGaussianBlurLayer(cRadius, workingLayer, srcLayer, toPreview, finalY * 2 + finalX * 2) Then
        
            'Next, create a contour of the layer
            If CreateContourLayer(useBlackBackground, srcLayer, workingLayer, toPreview, finalY * 2 + finalX * 2, finalY * 2) Then
            
                'Finally, white balance the resulting layer
                WhiteBalanceLayer 0.01, workingLayer, toPreview, finalY * 2 + finalX * 2, finalY * 2 + finalX
                
            End If
        End If
    End If
    
    srcLayer.eraseLayer
    Set srcLayer = Nothing
        
    'Pass control to finalizeImageData, which will handle the rest of the rendering using the data inside workingLayer
    finalizeImageData toPreview, dstPic

End Sub

Private Sub cmdBar_OKClick()
    Process "Trace contour", , buildParams(sltThickness, CBool(chkBlackBackground.Value), CBool(chkSmoothing.Value))
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
    If cmdBar.previewsAllowed Then TraceContour sltThickness, CBool(chkBlackBackground.Value), CBool(chkSmoothing.Value), True, fxPreview
End Sub

