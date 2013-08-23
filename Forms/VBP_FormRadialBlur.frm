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
   StartUpPosition =   1  'CenterOwner
   Begin PhotoDemon.smartCheckBox chkSpin 
      Height          =   540
      Left            =   6120
      TabIndex        =   9
      Top             =   1800
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   953
      Caption         =   "spin"
      Value           =   1
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
   Begin PhotoDemon.commandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   5
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
      TabIndex        =   1
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9922
   End
   Begin PhotoDemon.sliderTextCombo sltRadius 
      Height          =   495
      Left            =   6000
      TabIndex        =   4
      Top             =   2880
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   873
      Min             =   1
      Max             =   200
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
   Begin PhotoDemon.smartOptionButton OptInterpolate 
      Height          =   330
      Index           =   0
      Left            =   6120
      TabIndex        =   6
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
      TabIndex        =   7
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
   Begin PhotoDemon.smartCheckBox chkZoom 
      Height          =   540
      Left            =   7920
      TabIndex        =   10
      Top             =   1800
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   953
      Caption         =   "zoom"
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
      TabIndex        =   8
      Top             =   3600
      Width           =   1845
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "style:"
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
      TabIndex        =   3
      Top             =   1440
      Width           =   570
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
      TabIndex        =   2
      Top             =   4440
      Visible         =   0   'False
      Width           =   5775
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "radius:"
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
      TabIndex        =   0
      Top             =   2520
      Width           =   735
   End
End
Attribute VB_Name = "FormRadialBlur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Radial Blur Tool
'Copyright ©2012-2013 by Tanner Helland
'Created: 23/August/13
'Last updated: 23/August/13
'Last update: initial build
'
'To my knowledge, this tool is the first of its kind in VB6 - a radial blur tool that supports variable radii,
' as well as both zoom and spin modes (or both, if you'd like).  This function is mostly just a wrapper to PD's
' box blur and polar coordinate conversion functions; they do all the heavy lifting.
'
'Performance is pretty good, all things considered, but be careful in the IDE.  I STRONGLY recommend compiling
' the project before applying any actions at a large radius.
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

'Apply a radial blur an image
'Input: radius of the blur (min 1, no real max - but processing speed obviously drops as the radius increases)
Public Sub RadialBlurFilter(ByVal bRadius As Long, ByVal useSpinStyle As Boolean, ByVal useZoomStyle As Boolean, ByVal useBilinear As Boolean, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)
    
    If Not toPreview Then Message "Applying radial blur..."
    
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
            bRadius = (bRadius / iWidth) * curLayerValues.Width
        Else
            bRadius = (bRadius / iHeight) * curLayerValues.Height
        End If
        If bRadius = 0 Then bRadius = 1
    End If
    
    Dim finalX As Long, finalY As Long
    finalX = workingLayer.getLayerWidth
    finalY = workingLayer.getLayerHeight
    
    'Start by converting the image to polar coordinates, using a specific set of actions to maximize quality
    If CreatePolarCoordLayer(1, 100, EDGE_CLAMP, useBilinear, srcLayer, workingLayer, toPreview, finalX * 3) Then
    
        'Next, apply the box blur to that layer, using the radius supplied by the user
        Dim hRadius As Long, vRadius As Long
        hRadius = 1
        vRadius = 1
        
        If useZoomStyle Then vRadius = bRadius
        If useSpinStyle Then hRadius = bRadius
        
        If CreateBoxBlurLayer(hRadius, vRadius, workingLayer, srcLayer, toPreview, finalX * 3, finalX) Then
        
            'Finally, convert back to rectangular coordinates, using the opposite parameters of the first conversion
            CreatePolarCoordLayer 0, 100, EDGE_CLAMP, useBilinear, srcLayer, workingLayer, toPreview, finalX * 3, finalX * 2
            
        End If
        
    End If
    
    srcLayer.eraseLayer
    Set srcLayer = Nothing
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering using the data inside workingLayer
    finalizeImageData toPreview, dstPic
    
End Sub

Private Sub chkSpin_Click()
    
    'Always force one checkbox to be active
    If (Not CBool(chkSpin)) And (Not CBool(chkZoom)) Then
        cmdBar.markPreviewStatus False
        chkZoom.Value = vbChecked
        cmdBar.markPreviewStatus True
    End If
    
    updatePreview
    
End Sub

Private Sub chkZoom_Click()

    'Always force one checkbox to be active
    If (Not CBool(chkZoom)) And (Not CBool(chkSpin)) Then
        cmdBar.markPreviewStatus False
        chkSpin.Value = vbChecked
        cmdBar.markPreviewStatus True
    End If
    
    updatePreview
    
End Sub

Private Sub cmdBar_OKClick()
    Process "Radial blur", , buildParams(sltRadius, CBool(chkSpin), CBool(chkZoom), OptInterpolate(0))
End Sub

Private Sub cmdBar_RandomizeClick()

    'When randomizing, require at least one check box to be active
    If (Not CBool(chkZoom)) And (Not CBool(chkSpin)) Then
        If Int(Rnd * 2) = 0 Then chkSpin.Value = vbChecked Else chkZoom.Value = vbChecked
    End If

End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    updatePreview
End Sub

Private Sub cmdBar_ResetClick()
    chkSpin.Value = vbChecked
    chkZoom.Value = vbUnchecked
End Sub

Private Sub Form_Activate()

    'Assign the system hand cursor to all relevant objects
    Set m_ToolTip = New clsToolTip
    makeFormPretty Me, m_ToolTip
    
    'If the program is not compiled, display a special warning for this tool
    If Not g_IsProgramCompiled Then
        lblIDEWarning.Caption = g_Language.TranslateMessage("WARNING!  This tool has been heavily optimized, but at high radius values it will still be quite slow inside the IDE.  Please compile before applying or previewing any radius larger than 20.")
        lblIDEWarning.Visible = True
    End If
    
    'Draw a preview of the effect
    cmdBar.markPreviewStatus True
    updatePreview
    
End Sub

Private Sub Form_Load()
    
    'Disable previews until the form is fully initialized
    cmdBar.markPreviewStatus False
    
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

Private Sub OptInterpolate_Click(Index As Integer)
    updatePreview
End Sub

Private Sub OptStyle_Click(Index As Integer)
    updatePreview
End Sub

Private Sub sltRadius_Change()
    updatePreview
End Sub

'Render a new effect preview
Private Sub updatePreview()
    If cmdBar.previewsAllowed Then RadialBlurFilter sltRadius, CBool(chkSpin), CBool(chkZoom), OptInterpolate(0), True, fxPreview
End Sub
