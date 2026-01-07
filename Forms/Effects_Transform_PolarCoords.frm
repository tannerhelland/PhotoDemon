VERSION 5.00
Begin VB.Form FormPolar 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   " Polar conversion"
   ClientHeight    =   6540
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12105
   DrawStyle       =   5  'Transparent
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HasDC           =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   436
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   807
   Begin PhotoDemon.pdButtonStrip btsRender 
      Height          =   1095
      Left            =   6000
      TabIndex        =   6
      Top             =   4200
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1931
      Caption         =   "mode"
   End
   Begin PhotoDemon.pdCheckBox chkSwapXY 
      Height          =   330
      Left            =   6120
      TabIndex        =   1
      Top             =   1590
      Width           =   5670
      _ExtentX        =   10001
      _ExtentY        =   582
      Caption         =   "swap x and y coordinates"
   End
   Begin PhotoDemon.pdCommandBar cmdBar 
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5790
      Width           =   12105
      _ExtentX        =   21352
      _ExtentY        =   1323
   End
   Begin PhotoDemon.pdFxPreviewCtl pdFxPreview 
      Height          =   5625
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9922
      DisableZoomPan  =   -1  'True
   End
   Begin PhotoDemon.pdSlider sltRadius 
      Height          =   705
      Left            =   6000
      TabIndex        =   5
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
   Begin PhotoDemon.pdDropDown cboEdges 
      Height          =   735
      Left            =   6000
      TabIndex        =   2
      Top             =   3240
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1296
      Caption         =   "if pixels lie outside the image..."
   End
   Begin PhotoDemon.pdDropDown cboConvert 
      Height          =   735
      Left            =   6000
      TabIndex        =   4
      Top             =   720
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1296
      Caption         =   "conversion"
   End
End
Attribute VB_Name = "FormPolar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Image Polar Coordinate Conversion Tool
'Copyright 2013-2026 by Tanner Helland
'Created: 14/January/13
'Last updated: 28/July/17
'Last update: performance improvements, migrate to XML params
'
'This tool allows the user to convert an image between rectangular and polar coordinates.  An optional polar
' inversion technique is also supplied (as this is used by Paint.NET).
'
'The transformation used by this tool is a modified version of a transformation originally written by
' Jerry Huxtable of JH Labs.  Jerry's original code is licensed under an Apache 2.0 license.  You may download his
' original version at the following link (good as of 07 January '13): http://www.jhlabs.com/ip/filters/index.html
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

Private Sub btsRender_Click(ByVal buttonIndex As Long)
    UpdatePreview
End Sub

Private Sub cboConvert_Click()
    UpdatePreview
End Sub

Private Sub chkSwapXY_Click()
    UpdatePreview
End Sub

Private Sub cboEdges_Click()
    UpdatePreview
End Sub

'Convert an image to/from polar coordinates.
' INPUT PARAMETERS FOR CONVERSION:
' 0) Convert rectangular to polar
' 1) Convert polar to rectangular
' 2) Polar inversion
Public Sub ConvertToPolar(ByVal effectParams As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)

    If (Not toPreview) Then Message "Performing polar coordinate conversion..."
        
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString effectParams
    
    Dim conversionMethod As Long, swapXAndY As Boolean, polarRadius As Double
    Dim edgeHandling As Long, useBilinear As Boolean
    
    With cParams
        conversionMethod = .GetLong("method", cboConvert.ListIndex)
        swapXAndY = .GetBool("swapxy", False)
        polarRadius = .GetDouble("radius", sltRadius.Value)
        edgeHandling = .GetLong("edges", cboEdges.ListIndex)
        useBilinear = .GetBool("bilinear", True)
    End With
    
    'Create a local array and point it at the pixel data of the current image
    Dim dstSA As SafeArray2D
    EffectPrep.PrepImageData dstSA, toPreview, dstPic
    
    'Create a second local array.  This will contain the a copy of the current image, and we will use it as our source reference
    ' (This is necessary to prevent converted pixel values from spreading across the image as we go.)
    Dim srcDIB As pdDIB
    Set srcDIB = New pdDIB
    srcDIB.CreateFromExistingDIB workingDIB
    
    'Use the external function to create a polar coordinate DIB
    If swapXAndY Then
        CreatePolarCoordDIB conversionMethod, polarRadius, edgeHandling, useBilinear, srcDIB, workingDIB, toPreview
    Else
        CreateXSwappedPolarCoordDIB conversionMethod, polarRadius, edgeHandling, useBilinear, srcDIB, workingDIB, toPreview
    End If
    
    srcDIB.EraseDIB
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering using the data inside workingDIB
    EffectPrep.FinalizeImageData toPreview, dstPic
        
End Sub

'OK button
Private Sub cmdBar_OKClick()
    Process "Polar conversion", , GetLocalParamString(), UNDO_Layer
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub cmdBar_ResetClick()
    sltRadius.Value = 100
    chkSwapXY.Value = False
    cboEdges.ListIndex = pdeo_Erase
End Sub

Private Sub Form_Load()
    
    'Disable previews until the dialog is fully initialized
    cmdBar.SetPreviewStatus False
    
    btsRender.AddItem "fast", 0
    btsRender.AddItem "precise", 1
    btsRender.ListIndex = 1
    
    'I use a central function to populate the edge handling combo box; this way, I can add new methods and have
    ' them immediately available to all distort functions.
    PopDistortEdgeBox cboEdges, pdeo_Erase
    
    'Populate the polar conversion technique drop-down
    cboConvert.AddItem "Rectangular to polar", 0
    cboConvert.AddItem "Polar to rectangular", 1
    cboConvert.AddItem "Polar inversion", 2
    cboConvert.ListIndex = 0
    
    'Apply translations and visual themes
    ApplyThemeAndTranslations Me, True, True
    cmdBar.SetPreviewStatus True
    UpdatePreview
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub sltRadius_Change()
    UpdatePreview
End Sub

'Redraw the on-screen preview of the transformed image
Private Sub UpdatePreview()
    If cmdBar.PreviewsAllowed Then ConvertToPolar GetLocalParamString(), True, pdFxPreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub pdFxPreview_ViewportChanged()
    UpdatePreview
End Sub

Private Function GetLocalParamString() As String
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    
    With cParams
        .AddParam "method", cboConvert.ListIndex
        .AddParam "swapxy", chkSwapXY.Value
        .AddParam "radius", sltRadius.Value
        .AddParam "edges", cboEdges.ListIndex
        .AddParam "bilinear", (btsRender.ListIndex = 1)
    End With
    
    GetLocalParamString = cParams.GetParamString()
    
End Function
