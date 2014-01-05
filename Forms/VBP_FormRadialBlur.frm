VERSION 5.00
Begin VB.Form FormRadialBlur 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Radial Blur"
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
   Begin PhotoDemon.sliderTextCombo sltRadius 
      Height          =   495
      Left            =   6000
      TabIndex        =   4
      Top             =   2160
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   873
      Min             =   1
      Max             =   360
      SigDigits       =   1
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
   Begin PhotoDemon.smartOptionButton OptInterpolate 
      Height          =   330
      Index           =   0
      Left            =   6120
      TabIndex        =   5
      Top             =   3750
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
      Top             =   3750
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
      Top             =   2760
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
      Top             =   3360
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
      Top             =   4440
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
      Top             =   1800
      Width           =   660
   End
End
Attribute VB_Name = "FormRadialBlur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Radial Blur Tool
'Copyright ©2013-2014 by Tanner Helland
'Created: 26/August/13
'Last updated: 15/September/13
'Last update: adjust radius calculation method to produce correct ANGLE values - because of the polar-conversion
'              shortcut we use, angle is actually a ratio of the horizontal width of the image, where 360 degrees
'              is equivalent to the full width.  Now the output is identical to GIMP, Paint.NET, etc. (actually,
'              our output quality is better :) with no noticeable speed drop.
'
'To my knowledge, this tool is the first of its kind in VB6 - a radial blur tool that supports variable angles,
' and capable of operating in real-time.  This function is mostly just a wrapper to PD's horizontal blur and
' polar coordinate conversion functions; they do all the heavy lifting, as you can see from the code below.
'
'Performance is pretty good, all things considered, but be careful in the IDE. I STRONGLY recommend compiling
' the project before applying any actions at a large radius.
'
'All source code in this file is licensed under a modified BSD license. This means you may use the code in your own
' projects IF you provide attribution. For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Custom tooltip class allows for things like multiline, theming, and multiple monitor support
Dim m_ToolTip As clsToolTip

'Apply radial blur to an image
'Inputs: angle of the blur, and whether it should be symmetrical (e.g. equal in +/- angle amounts)
Public Sub RadialBlurFilter(ByVal bRadius As Double, ByVal blurSymmetrically As Boolean, ByVal useBilinear As Boolean, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)
    
    If Not toPreview Then Message "Applying radial blur..."
    
    'Create a local array and point it at the pixel data of the current image
    Dim dstSA As SAFEARRAY2D
    prepImageData dstSA, toPreview, dstPic
    
    'Create a second local array. This will contain the a copy of the current image, and we will use it as our source reference
    ' (This is necessary to prevent blurred pixel values from spreading across the image as we go.)
    Dim srcLayer As pdLayer
    Set srcLayer = New pdLayer
    srcLayer.createFromExistingLayer workingLayer
    
    'By dividing blur radius by 360 (its maximum value), we can use it as a fractional amount to determine the strength of our horizontal blur.
    Dim actualBlurSize As Long
    actualBlurSize = (bRadius / 360) * curLayerValues.Width
    If actualBlurSize < 1 Then actualBlurSize = 1
    
    Dim finalX As Long, finalY As Long
    finalX = workingLayer.getLayerWidth
    finalY = workingLayer.getLayerHeight
    
    'Because this function actually wraps three functions, calculating the progress bar maximum is a bit convoluted
    Dim newProgBarMax As Long
    newProgBarMax = finalX * 2 + (workingLayer.getLayerWidth + actualBlurSize * 2)
    
    'Start by converting the image to polar coordinates, using a specific set of actions to maximize quality
    If CreatePolarCoordLayer(1, 100, EDGE_CLAMP, useBilinear, srcLayer, workingLayer, toPreview, newProgBarMax) Then
    
        'We now need to do something a little unconventional.  When converting to polar coordinates, the line running from
        ' the top-center of the image to the center point ends up being separated onto the full left and right sides of the
        ' polar coordinate image.  Because PD's box blur does not wrap around image edges (for performance reasons), this line
        ' doesn't get blurred properly, and when we convert back to rectangular coordinates, it forms a visible abberation
        ' running from the top-center of the image to the center point.  To prevent this, we must create a temporary copy of
        ' the image that is larger (by the width of the blur radius) on both sides.  We then place the polar coord image in the
        ' center of this larger image, then copy the relevant edge pixels onto either side of it.  When the blur is complete,
        ' we copy back just this center portion before converting from polar to rect coords for the final time.  This results
        ' in a proper blur.  (Hope you caught all that!  :p)
        
        'Start by calculating the temporary image's size and offset
        Dim srcWidth As Long, srcHeight As Long
        srcWidth = workingLayer.getLayerWidth
        srcHeight = workingLayer.getLayerHeight
        
        Dim dstWidth As Long
        dstWidth = srcWidth + actualBlurSize * 2
        
        Dim dstX As Long
        dstX = (dstWidth - srcWidth) \ 2
        
        'Create a temporary layer to hold the blurred image
        Dim tmpLayer As pdLayer
        Set tmpLayer = New pdLayer
        tmpLayer.createBlank dstWidth, srcHeight, workingLayer.getLayerColorDepth
        
        'Bitblt the original image onto the center of the temporary canvas
        BitBlt tmpLayer.getLayerDC, dstX, 0, srcWidth, srcHeight, workingLayer.getLayerDC, 0, 0, vbSrcCopy
        
        'Apply two more blts - each of these will mirror an edge section of the source image
        BitBlt tmpLayer.getLayerDC, 0, 0, dstX, srcHeight, workingLayer.getLayerDC, srcWidth - dstX, 0, vbSrcCopy
        BitBlt tmpLayer.getLayerDC, dstX + srcWidth, 0, dstX, srcHeight, workingLayer.getLayerDC, 0, 0, vbSrcCopy
        
        'Change the srcLayer to be the same size as this working layer, so it can receive the fully blurred image
        srcLayer.createBlank tmpLayer.getLayerWidth, tmpLayer.getLayerHeight, workingLayer.getLayerColorDepth
    
        'Now we can apply the box blur to the temporary layer, using the blur radius supplied by the user
        Dim leftRadius As Long
        If blurSymmetrically Then leftRadius = actualBlurSize Else leftRadius = 0
        
        If CreateHorizontalBlurLayer(leftRadius, actualBlurSize, tmpLayer, srcLayer, toPreview, newProgBarMax, finalX) Then
        
            'Copy the blurred results of the source layer back into the temporary layer
            tmpLayer.createFromExistingLayer srcLayer
            
            'Resize the source layer to match the original image
            srcLayer.createBlank workingLayer.getLayerWidth, workingLayer.getLayerHeight, workingLayer.getLayerColorDepth
            
            'Copy the correct chunk of the temporary layer into the source layer
            BitBlt srcLayer.getLayerDC, 0, 0, srcWidth, srcHeight, tmpLayer.getLayerDC, dstX, 0, vbSrcCopy
            tmpLayer.eraseLayer
            
            'Finally, convert back to rectangular coordinates, using the opposite parameters of the first conversion
            CreatePolarCoordLayer 0, 100, EDGE_CLAMP, useBilinear, srcLayer, workingLayer, toPreview, newProgBarMax, finalX + dstWidth
            
        End If
        
        Set tmpLayer = Nothing
        
    End If
    
    srcLayer.eraseLayer
    Set srcLayer = Nothing
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering using the data inside workingLayer
    finalizeImageData toPreview, dstPic
    
End Sub

Private Sub chkSymmetry_Click()
    updatePreview
End Sub

Private Sub cmdBar_OKClick()
    Process "Radial blur", , buildParams(sltRadius, CBool(chkSymmetry), OptInterpolate(0))
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

Private Sub sltRadius_Change()
    updatePreview
End Sub

'Render a new effect preview
Private Sub updatePreview()
    If cmdBar.previewsAllowed Then RadialBlurFilter sltRadius, CBool(chkSymmetry), OptInterpolate(0), True, fxPreview
End Sub
