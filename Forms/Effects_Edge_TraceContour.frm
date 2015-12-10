VERSION 5.00
Begin VB.Form FormContour 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Trace contour"
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
   Begin PhotoDemon.smartCheckBox chkBlackBackground 
      Height          =   375
      Left            =   6120
      TabIndex        =   2
      Top             =   3120
      Width           =   5700
      _ExtentX        =   10054
      _ExtentY        =   582
      Caption         =   "use black background"
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
      Height          =   375
      Left            =   6120
      TabIndex        =   3
      Top             =   3720
      Width           =   5700
      _ExtentX        =   10054
      _ExtentY        =   582
      Caption         =   "apply contour smoothing"
   End
   Begin PhotoDemon.sliderTextCombo sltThickness 
      Height          =   705
      Left            =   6000
      TabIndex        =   4
      Top             =   2160
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
      Caption         =   "thickness"
      Min             =   1
      Max             =   30
      Value           =   1
   End
End
Attribute VB_Name = "FormContour"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Trace Contour (Outline) Tool
'Copyright 2013-2015 by Tanner Helland
'Created: 15/Feb/13
'Last updated: 10/January/14
'Last update: greatly improve performance by switching to approximate gaussian blur function
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
    Dim srcDIB As pdDIB
    Set srcDIB = New pdDIB
    srcDIB.createFromExistingDIB workingDIB
    
    'If this is a preview, we need to adjust the kernel radius to match the size of the preview box
    If toPreview Then
        cRadius = cRadius * curDIBValues.previewModifier
        If cRadius = 0 Then cRadius = 1
    End If
    
    Dim finalX As Long, finalY As Long
    finalX = workingDIB.getDIBWidth
    finalY = workingDIB.getDIBHeight
        
    If useSmoothing Then
    
        'Blur the current DIB
        If CreateApproximateGaussianBlurDIB(cRadius, srcDIB, workingDIB, 3, toPreview, finalX * 6 + finalY * 3) Then
        
            'Use the median filter to round out edges
            If CreateMedianDIB(cRadius, 50, PDPRS_Circle, workingDIB, srcDIB, toPreview, finalX * 6 + finalY * 3, finalX * 3 + finalY * 3) Then
        
                'Next, create a contour of the DIB
                If CreateContourDIB(useBlackBackground, srcDIB, workingDIB, toPreview, finalX * 6 + finalY * 3, finalX * 4 + finalY * 3) Then
            
                    'Finally, white balance the resulting DIB
                    WhiteBalanceDIB 0.01, workingDIB, toPreview, finalX * 6 + finalY * 3, finalX * 5 + finalY * 3
                    
                End If
            End If
        End If
    Else
        
        'Blur the current DIB
        If CreateApproximateGaussianBlurDIB(cRadius, workingDIB, srcDIB, 3, toPreview, finalX * 5 + finalY * 3) Then
        
            'Next, create a contour of the DIB
            If CreateContourDIB(useBlackBackground, srcDIB, workingDIB, toPreview, finalX * 5 + finalY * 3, finalX * 3 + finalY * 3) Then
            
                'Finally, white balance the resulting DIB
                WhiteBalanceDIB 0.01, workingDIB, toPreview, finalX * 5 + finalY * 3, finalX * 4 + finalY * 3
                
            End If
        End If
    End If
    
    srcDIB.eraseDIB
    Set srcDIB = Nothing
        
    'Pass control to finalizeImageData, which will handle the rest of the rendering using the data inside workingDIB
    finalizeImageData toPreview, dstPic

End Sub

Private Sub cmdBar_OKClick()
    Process "Trace contour", , buildParams(sltThickness, CBool(chkBlackBackground.Value), CBool(chkSmoothing.Value)), UNDO_LAYER
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

Private Sub chkBlackBackground_Click()
    updatePreview
End Sub

Private Sub sltThickness_Change()
    updatePreview
End Sub

Private Sub updatePreview()
    If cmdBar.previewsAllowed Then TraceContour sltThickness, CBool(chkBlackBackground.Value), CBool(chkSmoothing.Value), True, fxPreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub fxPreview_ViewportChanged()
    updatePreview
End Sub


