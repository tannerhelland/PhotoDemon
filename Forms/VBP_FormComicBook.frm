VERSION 5.00
Begin VB.Form FormComicBook 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Comic book"
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
   Begin PhotoDemon.sliderTextCombo sltInk 
      Height          =   495
      Left            =   6000
      TabIndex        =   4
      Top             =   2040
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Min             =   1
      Max             =   100
      Value           =   50
   End
   Begin PhotoDemon.sliderTextCombo sltColor 
      Height          =   495
      Left            =   6000
      TabIndex        =   5
      Top             =   3240
      Width           =   5925
      _ExtentX        =   10451
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Max             =   50
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "color smoothing:"
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
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ink:"
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
      Top             =   1680
      Width           =   405
   End
End
Attribute VB_Name = "FormComicBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Comic Book Image Effect
'Copyright ©2013-2014 by Tanner Helland
'Created: sometime 2013, I think??
'Last updated: 24/July/14
'Last update: overhauled algorithm, gave tool its own dialog
'
'PhotoDemon has provided a "comic book" effect for a long time, but despite going through many incarnations, it always
' used low-quality, "quick and dirty" approximations.
'
'In July '14, this changed, and the entire tool was rethought from the ground up.  A dialog is now provided, with
' various user-settable options.  This yields much more flexible results, and the use of PD's central compositor for
' overlaying intermediate image copies keeps things nice and fast.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Custom tooltip class allows for things like multiline, theming, and multiple monitor support
Dim m_ToolTip As clsToolTip

'Apply a "comic book" effect to an image
'Inputs:
' 1) strength of the inking
' 2) color smudging, which controls the radius of the median effect applied to the base image
Public Sub fxComicBook(ByVal inkOpacity As Long, ByVal colorSmudge As Long, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)
    
    If Not toPreview Then Message "Animating image (stage %1 of %2)...", 1, 3
    
    'Initiate PhotoDemon's central image handler
    Dim dstSA As SAFEARRAY2D
    prepImageData dstSA, toPreview, dstPic
    
    'During a preview, the smudge radius must be reduced to match the preview size
    If toPreview Then colorSmudge = colorSmudge * curDIBValues.previewModifier
    
    'Create two copies of the working image.  This filter overlays an inked image over a color-smudged version, and we need to
    ' handle these separately until the final step.
    Dim inkDIBh As pdDIB
    Set inkDIBh = New pdDIB
    inkDIBh.createFromExistingDIB workingDIB
    
    'To generate the inked image, we'll use bidirectional Sobel edge detection.
    Dim tmpParamHeader As String, finalParamString As String
    
    '1a) Generate a name for the requested filter (none, in this case, but it's required by the central convolver)
    tmpParamHeader = " |"
    
    '1b) Add in the invert (black background) parameter
    tmpParamHeader = tmpParamHeader & "True" & "|"
    
    'Now, build a specific convolution matrix for the horizontal direction.
    Dim convoString As String
    
        'Set divisor and offset
        convoString = "1|0|"
        
        'Build a horizontal convolution matrix
        convoString = convoString & "0|0|0|0|0|"
        convoString = convoString & "0|-1|0|1|0|"
        convoString = convoString & "0|-2|0|2|0|"
        convoString = convoString & "0|-1|0|1|0|"
        convoString = convoString & "0|0|0|0|0"
        
    'Merge the convolution header and matrix into a single param string
    finalParamString = tmpParamHeader & convoString
    
    'Use PD's central convolver to generate the horizontal edge image
    ConvolveDIB finalParamString, workingDIB, inkDIBh, toPreview, workingDIB.getDIBWidth * 3, 0
    
    'Repeat the steps above, but in the vertical direction
    If Not toPreview Then Message "Animating image (stage %1 of %2)...", 2, 3
    
    Dim inkDIBv As pdDIB
    Set inkDIBv = New pdDIB
    inkDIBv.createFromExistingDIB workingDIB
    
    'As before, create a Sobel edge detection matrix - but this time, in the vertical direction
    
        'Set divisor and offset
        convoString = "1|0|"
        
        'Build a vertical convolution matrix
        convoString = convoString & "0|0|0|0|0|"
        convoString = convoString & "0|1|2|1|0|"
        convoString = convoString & "0|0|0|0|0|"
        convoString = convoString & "0|-1|-2|-1|0|"
        convoString = convoString & "0|0|0|0|0"
    
    'Merge the convolution header and matrix into a single param string
    finalParamString = tmpParamHeader & convoString
    
    'Use PD's central convolver to generate the vertical edge image
    ConvolveDIB finalParamString, workingDIB, inkDIBv, toPreview, workingDIB.getDIBWidth * 3, workingDIB.getDIBWidth
    
    'With both ink images now available, we can composite the two into a single bidirectional ink image, using
    ' PD's central compositor class.
    Dim cComposite As pdCompositor
    Set cComposite = New pdCompositor
    
    cComposite.compositeDIBs inkDIBv, inkDIBh, BL_MULTIPLY, 0, 0
    
    'The bottom DIB (inkDIBh) now contains the composite image.  Release the vertical copy.
    Set inkDIBv = Nothing
    
    'Convert the ink DIB to grayscale
    GrayscaleDIB inkDIBh, True
    
    'We now need to obtain the underlying color-smudged version of the source image
    If Not toPreview Then Message "Animating image (stage %1 of %2)...", 3, 3
    
    If colorSmudge > 0 Then
        
        'Use PD's excellent bilateral smoothing function to handle color smudging.
        createBilateralDIB workingDIB, colorSmudge, 100, 2, 10, 10, toPreview, workingDIB.getDIBWidth * 4, workingDIB.getDIBWidth * 2
        
    End If
    
    'Finally, composite the ink over the color smudge, using the opacity supplied by the user.  To make the composite
    ' operation easier, we're going to place our DIBs inside temporary layers.  This allows us to use existing layer
    ' code to handle the merge.
    Dim tmpLayerTop As pdLayer, tmpLayerBottom As pdLayer
    Set tmpLayerTop = New pdLayer
    Set tmpLayerBottom = New pdLayer
    
    tmpLayerTop.CreateNewImageLayer inkDIBh
    Set inkDIBh = Nothing
    
    tmpLayerBottom.CreateNewImageLayer workingDIB
    workingDIB.eraseDIB
    
    tmpLayerTop.setLayerBlendMode BL_MULTIPLY
    tmpLayerTop.setLayerOpacity inkOpacity
    
    cComposite.mergeLayers tmpLayerTop, tmpLayerBottom, True
    Set tmpLayerTop = Nothing
    
    'Refresh the workingDIB instance, then exit!
    workingDIB.createFromExistingDIB tmpLayerBottom.layerDIB
    Set tmpLayerBottom = Nothing
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering using the data inside workingDIB
    finalizeImageData toPreview, dstPic
    
End Sub

Private Sub cmdBar_OKClick()
    Process "Comic book", , buildParams(sltInk, sltColor), UNDO_LAYER
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    updatePreview
End Sub

Private Sub cmdBar_ResetClick()
    sltInk.Value = 50
    sltColor.Value = 5
End Sub

Private Sub Form_Activate()
    
    'Assign the system hand cursor to all relevant objects
    Set m_ToolTip = New clsToolTip
    makeFormPretty Me, m_ToolTip
        
    'Draw a preview of the effect
    cmdBar.markPreviewStatus True
    updatePreview
    
End Sub

Private Sub Form_Load()
    
    'Disable previews until the dialog is fully loaded
    cmdBar.markPreviewStatus False
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'Render a new effect preview
Private Sub updatePreview()
    If cmdBar.previewsAllowed Then fxComicBook sltInk, sltColor, True, fxPreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub fxPreview_ViewportChanged()
    updatePreview
End Sub

Private Sub sltColor_Change()
    updatePreview
End Sub

Private Sub sltInk_Change()
    updatePreview
End Sub
