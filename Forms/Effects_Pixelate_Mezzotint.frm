VERSION 5.00
Begin VB.Form FormMezzotint 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Mezzotint"
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
   Begin PhotoDemon.buttonStrip btsType 
      Height          =   615
      Left            =   6000
      TabIndex        =   4
      Top             =   600
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1085
   End
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
      BackColor       =   14802140
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
   Begin PhotoDemon.sliderTextCombo sltSmoothness 
      Height          =   720
      Left            =   6000
      TabIndex        =   3
      Top             =   2520
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
      Caption         =   "smoothness"
      Max             =   100
      Value           =   10
      NotchPosition   =   2
      NotchValueCustom=   10
   End
   Begin PhotoDemon.sliderTextCombo sltRandom 
      Height          =   720
      Left            =   6000
      TabIndex        =   5
      Top             =   1560
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
      Caption         =   "randomness"
      Max             =   100
      Value           =   50
      NotchPosition   =   2
      NotchValueCustom=   50
   End
   Begin PhotoDemon.buttonStrip btsStippling 
      Height          =   615
      Left            =   6000
      TabIndex        =   6
      Top             =   4020
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1085
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "stippling"
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
      Index           =   3
      Left            =   6000
      TabIndex        =   7
      Top             =   3600
      Width           =   900
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "type"
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
      TabIndex        =   2
      Top             =   180
      Width           =   450
   End
End
Attribute VB_Name = "FormMezzotint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Mezzotint Effect Tool
'Copyright 2014-2015 by Tanner Helland
'Created: 03/April/15
'Last updated: 04/April/15
'Last update: wrap up initial build
'
'This tool is roughly inspired by Photoshop's Mezzotint, but with many more options, to give the user some control
' over the final result.  (I also have no idea how Photoshop's Mezzotint works, so this is merely a rough approximation
' of whatever they do.)
'
'Traditional mezzotinting was developed as a less labor-intensive alternative to traditional cross-hatching or
' stippling (http://en.wikipedia.org/wiki/Mezzotint).  It makes very little sense in a digital world, but Photoshop's
' incredibly weird implementation shows up frequently in effect tutorials, so even though the digital filter bears little
' resemblance to the traditional technique, there seems to be merit in including it.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Apply a Photoshop-like "mezzotint" effect to an image
Public Sub ApplyMezzotintEffect(ByVal mType As Long, ByVal mRandom As Long, ByVal mSmoothness As Long, ByVal mStipplingLevel As Long, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)
    
    Debug.Print mType, mRandom, mSmoothness, mStipplingLevel
    
    If Not toPreview Then Message "Mezzotinting image..."
    
    'The way we calculate mezzotinting varies depending on whether points or strokes are being used.
    
    'Start by prepping a workingDIB instance
    Dim dstImageData() As Byte
    Dim dstSA As SAFEARRAY2D
    prepImageData dstSA, toPreview, dstPic, , , True
    
    'Previews require us to adjust the coarseness parameter to match the preview size
    If toPreview Then
        mSmoothness = mSmoothness * curDIBValues.previewModifier
    
    'If this isn't a preview, prep the on-screen progress bar
    Else
        SetProgBarMax 8
        SetProgBarVal 0
    End If
    
    'From that, grab a grayscale map
    Dim grayMap() As Byte
    DIB_Handler.getDIBGrayscaleMap workingDIB, grayMap, True
    
    If Not toPreview Then SetProgBarVal 1
    
    'Randomness roughly corresponds to the strength of the "divots" used in the mezzotinting plate.  PD provides a graymap
    ' version of this, to which we simply supply the mRandom parameter (normalized from [0, 100] to [0, 255]).
    Filters_ByteArray.addNoiseByteArray grayMap, workingDIB.getDIBWidth, workingDIB.getDIBHeight, mRandom * 2.55
    
    If Not toPreview Then SetProgBarVal 2
    
    'Coarseness controls the amount of blurring applied
    If mSmoothness > 0 Then
        
        'Point and horizontal stroke mezzotinting blurs horizontally
        If (mType = 0) Or (mType = 1) Then
            Filters_ByteArray.HorizontalBlur_ByteArray grayMap, workingDIB.getDIBWidth, workingDIB.getDIBHeight, mSmoothness, mSmoothness
        End If
        
        'Point and vertical stroke mezzotinting blurs vertically
        If (mType = 0) Or (mType = 2) Then
            Filters_ByteArray.VerticalBlur_ByteArray grayMap, workingDIB.getDIBWidth, workingDIB.getDIBHeight, mSmoothness, mSmoothness
        End If
        
        If Not toPreview Then SetProgBarVal 3
        
        'After blurring, we want to white-balance the graymap, so that everything isn't just a muddy gray.
        Filters_ByteArray.ContrastCorrect_ByteArray grayMap, workingDIB.getDIBWidth, workingDIB.getDIBHeight, 10
        
        If Not toPreview Then SetProgBarVal 4
    
    End If
    
    'Further corrections are dependent on the level of stippling requested by the user
    Select Case mStipplingLevel
    
        'None
        Case 0
        
        'Coarse (monochrome, no dithering)
        Case 1
            Filters_ByteArray.Dither_ByteArray grayMap, workingDIB.getDIBWidth, workingDIB.getDIBHeight, 3, True
         
        'Fine (monochrome, with dithering)
        Case 2
            Filters_ByteArray.thresholdPlusDither_ByteArray grayMap, workingDIB.getDIBWidth, workingDIB.getDIBHeight, 127, True
            
    End Select
        
    If Not toPreview Then SetProgBarVal 5
    
    'Our overlay is now complete.  We now need to convert it back into a DIB.
    Dim overlayDIB As pdDIB
    DIB_Handler.createDIBFromGrayscaleMap overlayDIB, grayMap, workingDIB.getDIBWidth, workingDIB.getDIBHeight
    
    If Not toPreview Then SetProgBarVal 6
    
    'We can save a lot of time by avoiding alpha handling.  Query the base image to see if we need to deal with alpha.
    Dim alphaIsRelevant As Boolean
    alphaIsRelevant = Not DIB_Handler.isDIBAlphaBinary(workingDIB, False)
    
    If alphaIsRelevant Then
        overlayDIB.copyAlphaFromExistingDIB workingDIB
        overlayDIB.setAlphaPremultiplication True
    End If
    
    If Not toPreview Then SetProgBarVal 7
    
    'Finally, composite the new overlay DIB over working DIB.
    Dim cCompositor As pdCompositor
    Set cCompositor = New pdCompositor
    
    'Fine stippling uses a totally different approach, but the results are (IMO) much more interesting than Photoshop's
    If mStipplingLevel = 2 Then
        cCompositor.quickMergeTwoDibsOfEqualSize workingDIB, overlayDIB, BL_OVERLAY
    Else
        cCompositor.quickMergeTwoDibsOfEqualSize overlayDIB, workingDIB, BL_HARDMIX
    End If
    
    If Not toPreview Then SetProgBarVal 8
    
    'Erase our temporary image copy
    Set overlayDIB = Nothing
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering using the data inside workingDIB
    finalizeImageData toPreview, dstPic, True

End Sub

Private Sub btsStippling_Click(ByVal buttonIndex As Long)
    updatePreview
End Sub

Private Sub btsType_Click(ByVal buttonIndex As Long)
    updatePreview
End Sub

'OK button
Private Sub cmdBar_OKClick()
    Process "Mezzotint", , buildParams(btsType.ListIndex, sltRandom, sltSmoothness, btsStippling.ListIndex), UNDO_LAYER
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    updatePreview
End Sub

Private Sub cmdBar_ResetClick()
    sltRandom = 50
    sltSmoothness = 10
    btsType.ListIndex = 0
    btsStippling.ListIndex = 2
End Sub

Private Sub Form_Activate()
    
    'Apply translations and visual themes
    makeFormPretty Me
    
    'Draw a preview of the effect
    cmdBar.markPreviewStatus True
    updatePreview
    
End Sub

Private Sub Form_Load()
    
    'Disable previews while we initialize the dialog
    cmdBar.markPreviewStatus False
    
    'Populate the "type" button strip
    btsType.AddItem "dot", 0
    btsType.AddItem "horizontal stroke", 1
    btsType.AddItem "vertical stroke", 2
    btsType.ListIndex = 0
    
    'populate the "stippling" button strip
    btsStippling.AddItem "none", 0
    btsStippling.AddItem "coarse", 1
    btsStippling.AddItem "fine", 2
    btsStippling.ListIndex = 2
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub updatePreview()
    If cmdBar.previewsAllowed Then ApplyMezzotintEffect btsType.ListIndex, sltRandom, sltSmoothness, btsStippling.ListIndex, True, fxPreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub fxPreview_ViewportChanged()
    updatePreview
End Sub

Private Sub sltCoarseness_Change()
    updatePreview
End Sub

Private Sub sltRandom_Change()
    updatePreview
End Sub

Private Sub sltSmoothness_Change()
    updatePreview
End Sub
