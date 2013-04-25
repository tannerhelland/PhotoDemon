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
      TabIndex        =   8
      Top             =   3000
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
   Begin VB.TextBox txtRadius 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Left            =   11160
      TabIndex        =   6
      Text            =   "1"
      Top             =   2340
      Width           =   615
   End
   Begin VB.HScrollBar hsRadius 
      Height          =   255
      Left            =   6120
      Max             =   30
      Min             =   1
      TabIndex        =   5
      Top             =   2400
      Value           =   1
      Width           =   4935
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
      _ExtentX        =   9922
      _ExtentY        =   9922
   End
   Begin PhotoDemon.smartCheckBox chkSmoothing 
      Height          =   570
      Left            =   6120
      TabIndex        =   9
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
      TabIndex        =   7
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
'Last updated: 15/Feb/13
'Last update: initial build, though previously a simplified version of this was available from
'              Effects -> Edges -> Find Edges -> Artistic Contour
'
'This is a heavily optimized "extreme rank" function.  An accumulation technique is used instead of the standard sliding
' window mechanism.  (See http://web.archive.org/web/20060718054020/http://www.acm.uiuc.edu/siggraph/workshops/wjarosz_convolution_2001.pdf)
' This allows the algorithm to perform extremely well, despite being written in pure VB.
'
'That said, it is still unfortunately slow in the IDE.  I STRONGLY recommend compiling the project before applying any
' filter of a large radius (> 20).
'
'Extreme rank is a function of my own creation.  Basically, it performs both a minimum and a maxmimum rank calculation,
' and then it sets the pixel to whichever value is further from the current one.  This leads to an odd cut-out or stencil
' look unlike any other filter I've seen.  I'm not sure how much utility such a function provides, but it's fun so I
' include it.  :)
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
    If Not EntryValid(txtRadius, hsRadius.Min, hsRadius.Max, True, True) Then
        AutoSelectText txtRadius
        Exit Sub
    End If
    
    Me.Visible = False
    Process Contour, hsRadius.Value, CBool(chkBlackBackground.Value), CBool(chkSmoothing.Value)
    Unload Me
    
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

'These routines keep the scroll bar and text box values in sync
Private Sub hsRadius_Change()
    copyToTextBoxI txtRadius, hsRadius.Value
    updatePreview
End Sub

Private Sub hsRadius_Scroll()
    copyToTextBoxI txtRadius, hsRadius.Value
    updatePreview
End Sub

Private Sub chkBlackBackground_Click()
    updatePreview
End Sub

Private Sub txtRadius_KeyUp(KeyCode As Integer, Shift As Integer)
    textValidate txtRadius
    If EntryValid(txtRadius, hsRadius.Min, hsRadius.Max, False, False) Then hsRadius.Value = Val(txtRadius)
End Sub

Private Sub txtRadius_GotFocus()
    AutoSelectText txtRadius
End Sub

Private Sub updatePreview()
    If allowPreview Then TraceContour hsRadius.Value, CBool(chkBlackBackground.Value), CBool(chkSmoothing.Value), True, fxPreview
End Sub
