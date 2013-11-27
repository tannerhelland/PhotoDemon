VERSION 5.00
Begin VB.Form FormMotionBlur 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Motion Blur"
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
      TabIndex        =   2
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9922
   End
   Begin PhotoDemon.sliderTextCombo sltAngle 
      Height          =   495
      Left            =   6000
      TabIndex        =   4
      Top             =   1440
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   873
      Max             =   359.9
      SigDigits       =   1
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
   Begin PhotoDemon.smartOptionButton OptInterpolate 
      Height          =   330
      Index           =   0
      Left            =   6120
      TabIndex        =   5
      Top             =   3990
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   635
      Caption         =   "quality"
      Value           =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PhotoDemon.smartOptionButton OptInterpolate 
      Height          =   330
      Index           =   1
      Left            =   7920
      TabIndex        =   6
      Top             =   3990
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   635
      Caption         =   "speed"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PhotoDemon.smartCheckBox chkSymmetry 
      Height          =   480
      Left            =   6120
      TabIndex        =   8
      Top             =   3000
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   847
      Caption         =   "blur symmetrically"
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
   Begin PhotoDemon.sliderTextCombo sltDistance 
      Height          =   495
      Left            =   6000
      TabIndex        =   10
      Top             =   2400
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   873
      Min             =   1
      Max             =   500
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
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "distance:"
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
      Index           =   2
      Left            =   6000
      TabIndex        =   9
      Top             =   2040
      Width           =   945
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "render emphasis:"
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
      TabIndex        =   7
      Top             =   3600
      Width           =   1845
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
      TabIndex        =   3
      Top             =   4680
      Visible         =   0   'False
      Width           =   5775
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "angle:"
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
      Index           =   0
      Left            =   6000
      TabIndex        =   1
      Top             =   1080
      Width           =   660
   End
End
Attribute VB_Name = "FormMotionBlur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Motion Blur Tool
'Copyright ©2012-2013 by Tanner Helland
'Created: 26/August/13
'Last updated: 26/August/13
'Last update: initial build
'
'To my knowledge, this tool is the first of its kind in VB6 - a motion blur tool that supports variable angle
' and strength, while still capable of operating in real-time.  This function is mostly just a wrapper to PD's
' horizontal blur and rotate functions; they do all the heavy lifting, as you can see from the code below.
'
'Performance is pretty good, all things considered, but be careful in the IDE.  I STRONGLY recommend compiling
' the project before applying any actions at a large radius.
'
'If FreeImage is available, it is used to estimate a new size for the rotated image.  This is not the best way
' to estimate that value, but it's easier than doing the trig by hand, and FreeImage's rotate is *very* fast.  :)
'
'All source code in this file is licensed under a modified BSD license. This means you may use the code in your own
' projects IF you provide attribution. For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Custom tooltip class allows for things like multiline, theming, and multiple monitor support
Dim m_ToolTip As clsToolTip

'Apply motion blur to an image
'Inputs: angle of the blur, distance of the blur
Public Sub MotionBlurFilter(ByVal bAngle As Double, ByVal bDistance As Long, ByVal blurSymmetrically As Boolean, ByVal useBilinear As Boolean, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)
    
    If Not toPreview Then Message "Applying motion blur..."
    
    'Call prepImageData, which will initialize a workingLayer object for us (with all selection tool masks applied)
    Dim dstSA As SAFEARRAY2D
    prepImageData dstSA, toPreview, dstPic
    
    'If this is a preview, we need to adjust the kernel radius to match the size of the preview box
    If toPreview Then
        bDistance = bDistance * curLayerValues.previewModifier
        If bDistance = 0 Then bDistance = 1
    End If
    
    Dim finalX As Long, finalY As Long
    finalX = workingLayer.getLayerWidth
    finalY = workingLayer.getLayerHeight
    
    'Before doing any rotating or blurring, we need to increase the size of the image we're working with.  If we
    ' don't do this, the rotation will chop off the image's corners, and the resulting motion blur will look terrible.
        
    'If FreeImage is enabled, use it to calculate an optimal extension size.  If it is not enabled, do a
    ' quick-and-dirty estimation using basic geometry.
    Dim hScaleAmount As Long, vScaleAmount As Long
    If g_ImageFormats.FreeImageEnabled Then
                
        'Convert our current layer to a FreeImage-type DIB
        Dim fi_DIB As Long
        fi_DIB = FreeImage_CreateFromDC(workingLayer.getLayerDC)
        
        'Use that handle to request an image rotation, then store the new image's width and height
        Dim nWidth As Long, nHeight As Long
        If fi_DIB <> 0 Then
        
            Dim returnDIB As Long
            returnDIB = FreeImage_Rotate(fi_DIB, -bAngle, 0)
                    
            nWidth = FreeImage_GetWidth(returnDIB)
            nHeight = FreeImage_GetHeight(returnDIB)
            
            If returnDIB <> 0 Then FreeImage_Unload returnDIB
            FreeImage_Unload fi_DIB
    
        Else
            nWidth = workingLayer.getLayerWidth * 2
            nHeight = workingLayer.getLayerHeight * 2
        End If
        
        'Use the returned size to calculate optimal offsets
        hScaleAmount = (nWidth - workingLayer.getLayerWidth) \ 2
        vScaleAmount = (nHeight - workingLayer.getLayerHeight) \ 2
        
        If hScaleAmount < 0 Then hScaleAmount = 0
        If vScaleAmount < 0 Then vScaleAmount = 0
        
    Else
        
        'This is basically a worst-case estimation of the final image size, and because of that, the function will
        ' be quite slow.  This is a very fringe case, however, as most users should have FreeImage available.
        hScaleAmount = Sqr(workingLayer.getLayerWidth * workingLayer.getLayerWidth + workingLayer.getLayerHeight * workingLayer.getLayerHeight)
        If toPreview Then hScaleAmount = hScaleAmount \ 4 Else hScaleAmount = hScaleAmount \ 2
        vScaleAmount = hScaleAmount
        
    End If
    
    'I built a separate function to enlarge the image and fill the blank borders with clamped pixels from the source image:
    Dim tmpClampLayer As pdLayer
    Set tmpClampLayer = New pdLayer
    CreateExtendedLayer hScaleAmount, vScaleAmount, workingLayer, tmpClampLayer
    
    'Create a second layer, which will receive the results of this one
    Dim rotateLayer As pdLayer
    Set rotateLayer = New pdLayer
    rotateLayer.createBlank tmpClampLayer.getLayerWidth, tmpClampLayer.getLayerHeight, tmpClampLayer.getLayerColorDepth
    
    'Start by rotating the image by the requested amount.  Clamped edges are used to improve the blur output below
    If CreateRotatedLayer(bAngle, EDGE_CLAMP, useBilinear, tmpClampLayer, rotateLayer, toPreview, tmpClampLayer.getLayerWidth * 3) Then
    
        'Next, apply a horizontal blur, using the blur radius supplied by the user
        Dim rightRadius As Long
        If blurSymmetrically Then rightRadius = bDistance Else rightRadius = 0
        
        If CreateHorizontalBlurLayer(bDistance, rightRadius, rotateLayer, tmpClampLayer, toPreview, tmpClampLayer.getLayerWidth * 3, tmpClampLayer.getLayerWidth) Then
            
            'Finally, rotate the image back to its original orientation, using the opposite parameters of the first conversion
            CreateRotatedLayer -bAngle, EDGE_CLAMP, useBilinear, tmpClampLayer, rotateLayer, toPreview, tmpClampLayer.getLayerWidth * 3, tmpClampLayer.getLayerWidth * 2
            
            'Erase the temporary clamp layer
            tmpClampLayer.eraseLayer
            Set tmpClampLayer = Nothing
            
            'rotateLayer now contains the image we want, but it also has all the (now-useless) padding from
            ' the rotate operation.  Chop out the valid section and copy it into workingLayer.
            BitBlt workingLayer.getLayerDC, 0, 0, workingLayer.getLayerWidth, workingLayer.getLayerHeight, rotateLayer.getLayerDC, hScaleAmount, vScaleAmount, vbSrcCopy
            
            'Erase the temporary rotation layer
            rotateLayer.eraseLayer
            Set rotateLayer = Nothing
            
        End If
        
    End If
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering using the data inside workingLayer
    finalizeImageData toPreview, dstPic
    
End Sub

Private Sub chkSymmetry_Click()
    updatePreview
End Sub

Private Sub cmdBar_OKClick()
    Process "Motion blur", , buildParams(sltAngle, sltDistance, CBool(chkSymmetry), OptInterpolate(0))
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
    cmdBar.markPreviewStatus True
    updatePreview
    
End Sub

Private Sub Form_Load()
    
    'Disable previews until the form is fully initialized
    cmdBar.markPreviewStatus False
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub OptInterpolate_Click(Index As Integer)
    updatePreview
End Sub

'Render a new effect preview
Private Sub updatePreview()
    If cmdBar.previewsAllowed Then MotionBlurFilter sltAngle, sltDistance, CBool(chkSymmetry), OptInterpolate(0), True, fxPreview
End Sub

Private Sub sltAngle_Change()
    updatePreview
End Sub

Private Sub sltDistance_Change()
    updatePreview
End Sub
