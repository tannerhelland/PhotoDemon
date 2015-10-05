VERSION 5.00
Begin VB.Form FormRotateDistort 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Rotate"
   ClientHeight    =   6510
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   12090
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
   ScaleHeight     =   434
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   806
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin PhotoDemon.commandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5760
      Width           =   12090
      _ExtentX        =   21325
      _ExtentY        =   1323
      BackColor       =   14802140
   End
   Begin VB.ComboBox cmbEdges 
      BackColor       =   &H00FFFFFF&
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
      Left            =   6120
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   3015
      Width           =   5700
   End
   Begin PhotoDemon.fxPreviewCtl fxPreview 
      Height          =   5625
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9922
      DisableZoomPan  =   -1  'True
      PointSelection  =   -1  'True
   End
   Begin PhotoDemon.smartOptionButton OptInterpolate 
      Height          =   360
      Index           =   0
      Left            =   6120
      TabIndex        =   3
      Top             =   3960
      Width           =   5685
      _ExtentX        =   10028
      _ExtentY        =   582
      Caption         =   "quality"
      Value           =   -1  'True
   End
   Begin PhotoDemon.smartOptionButton OptInterpolate 
      Height          =   360
      Index           =   1
      Left            =   6120
      TabIndex        =   4
      Top             =   4380
      Width           =   5685
      _ExtentX        =   10028
      _ExtentY        =   582
      Caption         =   "speed"
   End
   Begin PhotoDemon.sliderTextCombo sltAngle 
      Height          =   720
      Left            =   6000
      TabIndex        =   7
      Top             =   1680
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
      Caption         =   "angle"
      Min             =   -360
      Max             =   360
      SigDigits       =   1
   End
   Begin PhotoDemon.sliderTextCombo sltXCenter 
      Height          =   405
      Left            =   6000
      TabIndex        =   9
      Top             =   600
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   873
      Max             =   1
      SigDigits       =   2
      Value           =   0.5
      NotchPosition   =   2
      NotchValueCustom=   0.5
   End
   Begin PhotoDemon.sliderTextCombo sltYCenter 
      Height          =   405
      Left            =   9000
      TabIndex        =   10
      Top             =   600
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   873
      Max             =   1
      SigDigits       =   2
      Value           =   0.5
      NotchPosition   =   2
      NotchValueCustom=   0.5
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "center position (x, y)"
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
      TabIndex        =   12
      Top             =   240
      Width           =   2205
   End
   Begin VB.Label lblExplanation 
      BackStyle       =   0  'Transparent
      Caption         =   "Note: you can also set a center position by clicking the preview window."
      ForeColor       =   &H00404040&
      Height          =   435
      Index           =   0
      Left            =   6120
      TabIndex        =   11
      Top             =   1170
      Width           =   5655
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblExplanation 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   885
      Index           =   1
      Left            =   6000
      TabIndex        =   8
      Top             =   4800
      Width           =   5925
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "render emphasis"
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
      TabIndex        =   6
      Top             =   3570
      Width           =   1755
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "if pixels lie outside the image..."
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
      Index           =   5
      Left            =   6000
      TabIndex        =   5
      Top             =   2640
      Width           =   3315
   End
End
Attribute VB_Name = "FormRotateDistort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Rotate Distort Effect Interface (separate from image rotation for a reason - see below)
'Copyright 2013-2015 by Tanner Helland
'Created: 22/August/13
'Last updated: 10/January/14
'Last update: new feature allows the user to select a custom center point for the rotation
'
'Dialog for handling rotation via PhotoDemon's distort filter engine.  This is kept separate from full-image rotation,
' because I needed a rotate that could be applied to selections.  Also, full-image rotation allows you to resize the
' canvas.  This does not.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

Private Sub cmbEdges_Click()
    updatePreview
End Sub

'Apply a basic rotation to the image or selected area
Public Sub RotateFilter(ByVal rotateAngle As Double, ByVal edgeHandling As Long, ByVal useBilinear As Boolean, Optional ByVal centerX As Double = 0.5, Optional ByVal centerY As Double = 0.5, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)
    
    If Not toPreview Then Message "Rotating area..."
    
    'Create a local array and point it at the pixel data of the current image
    Dim dstSA As SAFEARRAY2D
    prepImageData dstSA, toPreview, dstPic
    
    'Create a second local array.  This will contain the a copy of the current image, and we will use it as our source reference
    ' (This is necessary to prevent converted pixel values from spreading across the image as we go.)
    Dim srcDIB As pdDIB
    Set srcDIB = New pdDIB
    srcDIB.createFromExistingDIB workingDIB
    
    'Use the external function to create a rotated DIB
    CreateRotatedDIB rotateAngle, edgeHandling, useBilinear, srcDIB, workingDIB, centerX, centerY, toPreview
    
    srcDIB.eraseDIB
    Set srcDIB = Nothing
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    finalizeImageData toPreview, dstPic
        
    
End Sub

Private Sub cmdBar_OKClick()
    Process "Rotate", , buildParams(sltAngle.Value, CLng(cmbEdges.ListIndex), OptInterpolate(0).Value, sltXCenter.Value, sltYCenter.Value), UNDO_LAYER
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    updatePreview
End Sub

Private Sub cmdBar_ResetClick()
    sltXCenter.Value = 0.5
    sltYCenter.Value = 0.5
    cmbEdges.ListIndex = EDGE_WRAP
End Sub

Private Sub Form_Activate()

    'Apply translations and visual themes
    MakeFormPretty Me
    
    'Provide an explanation on why this tool doesn't enlarge the canvas to match
    lblExplanation(1).Caption = g_Language.TranslateMessage("If you want to enlarge the canvas to fit the rotated image, please use the Image -> Rotate menu instead.")
    
    'Request a preview
    cmdBar.markPreviewStatus True
    updatePreview
        
End Sub

Private Sub Form_Load()

    'Suspend previews while we initialize all the controls
    cmdBar.markPreviewStatus False
    
    'I use a central function to populate the edge handling combo box; this way, I can add new methods and have
    ' them immediately available to all distort functions.
    PopDistortEdgeBox cmbEdges, EDGE_WRAP
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'Redraw the effect preview
Private Sub updatePreview()
    If cmdBar.previewsAllowed Then RotateFilter sltAngle.Value, CLng(cmbEdges.ListIndex), OptInterpolate(0).Value, sltXCenter.Value, sltYCenter.Value, True, fxPreview
End Sub

'The user can right-click the preview area to select a new center point
Private Sub fxPreview_PointSelected(xRatio As Double, yRatio As Double)
    
    cmdBar.markPreviewStatus False
    sltXCenter.Value = xRatio
    sltYCenter.Value = yRatio
    cmdBar.markPreviewStatus True
    updatePreview

End Sub

Private Sub OptInterpolate_Click(Index As Integer)
    updatePreview
End Sub

Private Sub sltAngle_Change()
    updatePreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub fxPreview_ViewportChanged()
    updatePreview
End Sub

Private Sub sltXCenter_Change()
    updatePreview
End Sub

Private Sub sltYCenter_Change()
    updatePreview
End Sub
