VERSION 5.00
Begin VB.Form FormMotionBlur 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Motion blur"
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
   Begin PhotoDemon.buttonStrip btsStyle 
      Height          =   615
      Left            =   6120
      TabIndex        =   5
      Top             =   3600
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   1085
   End
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   375
      Left            =   6000
      Top             =   3120
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   661
      Caption         =   "style"
      FontSize        =   12
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
   Begin PhotoDemon.sliderTextCombo sltAngle 
      Height          =   720
      Left            =   6000
      TabIndex        =   2
      Top             =   1200
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
      Caption         =   "angle"
      Max             =   359.9
      SigDigits       =   1
   End
   Begin PhotoDemon.smartCheckBox chkSymmetry 
      Height          =   330
      Left            =   6120
      TabIndex        =   3
      Top             =   4440
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   582
      Caption         =   "blur symmetrically"
   End
   Begin PhotoDemon.sliderTextCombo sltDistance 
      Height          =   720
      Left            =   6000
      TabIndex        =   4
      Top             =   2160
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
      Caption         =   "distance"
      Min             =   1
      Max             =   500
      Value           =   5
   End
End
Attribute VB_Name = "FormMotionBlur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Motion Blur Tool
'Copyright 2013-2015 by Tanner Helland
'Created: 26/August/13
'Last updated: 02/October/15
'Last update: rewrite against new all-in-one rotate/edge-extend function
'
'To my knowledge, this tool is the first of its kind in VB6 - a motion blur tool that supports variable angle
' and strength, while still capable of operating in real-time.  This function is mostly just a wrapper to PD's
' horizontal blur and rotate functions; they do all the heavy lifting, as you can see from the code below.
'
'Performance is an order of magnitude faster than GIMP or Paint.NET, and even when uncompiled, we're *still*
' faster than either program.  Not bad, eh?
'
'All source code in this file is licensed under a modified BSD license. This means you may use the code in your own
' projects IF you provide attribution. For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Apply motion blur to an image
'Inputs: angle of the blur, distance of the blur
Public Sub MotionBlurFilter(ByVal bAngle As Double, ByVal bDistance As Long, ByVal blurSymmetrically As Boolean, ByVal blurAlgorithm As Long, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)
    
    If Not toPreview Then Message "Applying motion blur..."
    
    'Call prepImageData, which will initialize a workingDIB object for us (with all selection tool masks applied)
    Dim dstSA As SAFEARRAY2D
    prepImageData dstSA, toPreview, dstPic, , , True
    
    'If this is a preview, we need to adjust the kernel radius to match the size of the preview box
    If toPreview Then
        bDistance = bDistance * curDIBValues.previewModifier
        If bDistance = 0 Then bDistance = 1
    End If
    
    Dim finalX As Long, finalY As Long
    finalX = workingDIB.getDIBWidth
    finalY = workingDIB.getDIBHeight
    
    'Create a second DIB, which will receive the results of this one
    Dim rotateDIB As pdDIB
    Set rotateDIB = New pdDIB
    
    'As of October 2015, I've finally cracked the math to have GDI+ generate a rotated+padded+clamped DIB for us.
    ' This greatly simplifies this function, while also providing higher-quality results!
    GDI_Plus.GDIPlus_GetRotatedClampedDIB workingDIB, rotateDIB, bAngle
    
    'Next, apply a horizontal blur to the rotated image, using the blur radius supplied by the user
    Dim rightRadius As Long
    If blurSymmetrically Then rightRadius = bDistance Else rightRadius = 0
    
    Dim blurSuccess As Boolean
    
    'Motion blur currently supports two different blur algorithms
    Select Case blurAlgorithm
    
        'Box blur (requires intermediary DIB, as the transform can't be performed in-place)
        Case 0
            Dim tmpDIB As pdDIB
            Set tmpDIB = New pdDIB
            tmpDIB.createFromExistingDIB rotateDIB
            blurSuccess = CreateHorizontalBlurDIB(bDistance, rightRadius, tmpDIB, rotateDIB, toPreview)
            Set tmpDIB = Nothing
        
        'Gaussian blur (IIR estimation, no intermediary DIB required)
        Case 1
            blurSuccess = Filters_Area.HorizontalBlur_IIR(rotateDIB, bDistance, 1, blurSymmetrically, toPreview)
    
    End Select
    
    If blurSuccess Then
            
        'Finally, we need to rotate the image back to its original orientation, using the opposite parameters of the
        ' first conversion.
        
        'Use GDI+ to apply the inverse rotation.  Note that it will automatically center the rotated image within
        ' the destination boundaries, sparing us the trouble of manually trimming the clamped edges
        GDI_Plus.GDIPlus_RotateDIBPlgStyle rotateDIB, workingDIB, -bAngle, True
        
    End If
    
    'Release our temporary rotation DIB
    rotateDIB.eraseDIB
    Set rotateDIB = Nothing
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering using the data inside workingDIB
    finalizeImageData toPreview, dstPic, True
    
End Sub

Private Sub btsStyle_Click(ByVal buttonIndex As Long)
    updatePreview
End Sub

Private Sub chkSymmetry_Click()
    updatePreview
End Sub

Private Sub cmdBar_OKClick()
    Process "Motion blur", , buildParams(sltAngle, sltDistance, CBool(chkSymmetry), btsStyle.ListIndex), UNDO_LAYER
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    updatePreview
End Sub

Private Sub Form_Activate()

    'Apply visual themes and translations
    MakeFormPretty Me
    
    'Draw a preview of the effect
    cmdBar.markPreviewStatus True
    updatePreview
    
End Sub

Private Sub Form_Load()
    
    'Disable previews until the form is fully initialized
    cmdBar.markPreviewStatus False
    
    btsStyle.AddItem "constant", 0
    btsStyle.AddItem "gaussian", 1
    btsStyle.ListIndex = 0
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'Render a new effect preview
Private Sub updatePreview()
    If cmdBar.previewsAllowed Then MotionBlurFilter sltAngle, sltDistance, CBool(chkSymmetry), btsStyle.ListIndex, True, fxPreview
End Sub

Private Sub sltAngle_Change()
    updatePreview
End Sub

Private Sub sltDistance_Change()
    updatePreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub fxPreview_ViewportChanged()
    updatePreview
End Sub

