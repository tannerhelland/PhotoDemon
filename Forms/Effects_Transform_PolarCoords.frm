VERSION 5.00
Begin VB.Form FormPolar 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Polar Coordinate Conversion"
   ClientHeight    =   6540
   ClientLeft      =   -15
   ClientTop       =   225
   ClientWidth     =   12105
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
   ScaleWidth      =   807
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.smartCheckBox chkSwapXY 
      Height          =   330
      Left            =   6120
      TabIndex        =   10
      Top             =   1590
      Width           =   5670
      _ExtentX        =   10001
      _ExtentY        =   582
      Caption         =   "swap x and y coordinates"
   End
   Begin PhotoDemon.commandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5790
      Width           =   12105
      _ExtentX        =   21352
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
      TabIndex        =   5
      Top             =   3585
      Width           =   5700
   End
   Begin VB.ComboBox cboConvert 
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
      TabIndex        =   4
      Top             =   1200
      Width           =   5700
   End
   Begin PhotoDemon.fxPreviewCtl fxPreview 
      Height          =   5625
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9922
      DisableZoomPan  =   -1  'True
   End
   Begin PhotoDemon.smartOptionButton OptInterpolate 
      Height          =   360
      Index           =   0
      Left            =   6120
      TabIndex        =   7
      Top             =   4560
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
      TabIndex        =   8
      Top             =   4920
      Width           =   5685
      _ExtentX        =   10028
      _ExtentY        =   582
      Caption         =   "speed"
   End
   Begin PhotoDemon.sliderTextCombo sltRadius 
      Height          =   720
      Left            =   6000
      TabIndex        =   9
      Top             =   2280
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
      Caption         =   "radius (percentage)"
      Min             =   1
      Max             =   100
      Value           =   100
      NotchPosition   =   2
      NotchValueCustom=   100
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
      TabIndex        =   6
      Top             =   3210
      Width           =   3315
   End
   Begin VB.Label lblInterpolation 
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
      Left            =   6000
      TabIndex        =   2
      Top             =   4170
      Width           =   1755
   End
   Begin VB.Label lblConvert 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "conversion"
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
      TabIndex        =   1
      Top             =   840
      Width           =   1140
   End
End
Attribute VB_Name = "FormPolar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Image Polar Coordinate Conversion Tool
'Copyright 2013-2015 by Tanner Helland
'Created: 14/January/13
'Last updated: 23/August/13
'Last update: added command bar, converted the polar coordinate routine itself to operate on any two DIBs
'             (thus making this dialog just a thin wrapper to that function)
'
'This tool allows the user to convert an image between rectangular and polar coordinates.  An optional polar
' inversion technique is also supplied (as this is used by Paint.NET).
'
'The transformation used by this tool is a modified version of a transformation originally written by
' Jerry Huxtable of JH Labs.  Jerry's original code is licensed under an Apache 2.0 license.  You may download his
' original version at the following link (good as of 07 January '13): http://www.jhlabs.com/ip/filters/index.html
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

Private Sub cboConvert_Click()
    updatePreview
End Sub

Private Sub chkSwapXY_Click()
    updatePreview
End Sub

Private Sub cmbEdges_Click()
    updatePreview
End Sub

'Convert an image to/from polar coordinates.
' INPUT PARAMETERS FOR CONVERSION:
' 0) Convert rectangular to polar
' 1) Convert polar to rectangular
' 2) Polar inversion
Public Sub ConvertToPolar(ByVal conversionMethod As Long, ByVal swapXAndY As Boolean, ByVal polarRadius As Double, ByVal edgeHandling As Long, ByVal useBilinear As Boolean, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)

    If Not toPreview Then Message "Performing polar coordinate conversion..."
        
    'Create a local array and point it at the pixel data of the current image
    Dim dstSA As SAFEARRAY2D
    prepImageData dstSA, toPreview, dstPic
    
    'Create a second local array.  This will contain the a copy of the current image, and we will use it as our source reference
    ' (This is necessary to prevent converted pixel values from spreading across the image as we go.)
    Dim srcDIB As pdDIB
    Set srcDIB = New pdDIB
    srcDIB.createFromExistingDIB workingDIB
    
    'Use the external function to create a polar coordinate DIB
    If swapXAndY Then
        CreatePolarCoordDIB conversionMethod, polarRadius, edgeHandling, useBilinear, srcDIB, workingDIB, toPreview
    Else
        CreateXSwappedPolarCoordDIB conversionMethod, polarRadius, edgeHandling, useBilinear, srcDIB, workingDIB, toPreview
    End If
    
    srcDIB.eraseDIB
    Set srcDIB = Nothing
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering using the data inside workingDIB
    finalizeImageData toPreview, dstPic
        
End Sub

'OK button
Private Sub cmdBar_OKClick()
    Process "Polar conversion", , buildParams(cboConvert.ListIndex, CBool(chkSwapXY), sltRadius.Value, CLng(cmbEdges.ListIndex), OptInterpolate(0).Value), UNDO_LAYER
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    updatePreview
End Sub

Private Sub cmdBar_ResetClick()
    sltRadius.Value = 100
    chkSwapXY.Value = vbUnchecked
    cmbEdges.ListIndex = EDGE_ERASE
End Sub

Private Sub Form_Activate()

    'Apply translations and visual themes
    MakeFormPretty Me
        
    'Create the preview
    cmdBar.markPreviewStatus True
    updatePreview
    
End Sub

Private Sub Form_Load()
    
    'Disable previews until the dialog is fully initialized
    cmdBar.markPreviewStatus False
    
    'I use a central function to populate the edge handling combo box; this way, I can add new methods and have
    ' them immediately available to all distort functions.
    PopDistortEdgeBox cmbEdges, EDGE_ERASE
    
    'Populate the polar conversion technique drop-down
    cboConvert.AddItem "Rectangular to polar", 0
    cboConvert.AddItem "Polar to rectangular", 1
    cboConvert.AddItem "Polar inversion", 2
    cboConvert.ListIndex = 0
    
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

'Redraw the on-screen preview of the transformed image
Private Sub updatePreview()
    If cmdBar.previewsAllowed Then ConvertToPolar cboConvert.ListIndex, CBool(chkSwapXY), sltRadius.Value, CLng(cmbEdges.ListIndex), OptInterpolate(0).Value, True, fxPreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub fxPreview_ViewportChanged()
    updatePreview
End Sub

