VERSION 5.00
Begin VB.Form FormPosterize 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   " Posterize"
   ClientHeight    =   6540
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11970
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
   ScaleWidth      =   798
   Begin PhotoDemon.pdButtonStrip btsAdaptiveColoring 
      Height          =   1020
      Left            =   6000
      TabIndex        =   7
      Top             =   2640
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1799
      Caption         =   "adaptive coloring"
   End
   Begin PhotoDemon.pdCommandBar cmdBar 
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5790
      Width           =   11970
      _ExtentX        =   21114
      _ExtentY        =   1323
   End
   Begin PhotoDemon.pdFxPreviewCtl pdFxPreview 
      Height          =   5625
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9922
   End
   Begin PhotoDemon.pdSlider sltRed 
      Height          =   735
      Left            =   6000
      TabIndex        =   2
      Top             =   120
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1296
      Caption         =   "possible red values"
      Min             =   2
      Max             =   64
      Value           =   6
      DefaultValue    =   6
   End
   Begin PhotoDemon.pdSlider sltGreen 
      Height          =   735
      Left            =   6000
      TabIndex        =   3
      Top             =   960
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1296
      Caption         =   "possible green values"
      Min             =   2
      Max             =   64
      Value           =   7
      DefaultValue    =   6
   End
   Begin PhotoDemon.pdSlider sltBlue 
      Height          =   735
      Left            =   6000
      TabIndex        =   4
      Top             =   1800
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1296
      Caption         =   "possible blue values"
      Min             =   2
      Max             =   64
      Value           =   6
      DefaultValue    =   6
   End
   Begin PhotoDemon.pdSlider sldDitherAmount 
      Height          =   735
      Left            =   6000
      TabIndex        =   5
      Top             =   4800
      Width           =   5895
      _ExtentX        =   10610
      _ExtentY        =   1296
      Caption         =   "dithering amount"
      Max             =   100
      Value           =   50
      GradientColorRight=   1703935
      DefaultValue    =   50
   End
   Begin PhotoDemon.pdDropDown cboDither 
      Height          =   780
      Left            =   6000
      TabIndex        =   6
      Top             =   3810
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1376
      Caption         =   "dithering"
   End
End
Attribute VB_Name = "FormPosterize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Posterizing Effect Handler
'Copyright 2001-2026 by Tanner Helland
'Created: 4/15/01
'Last updated: 24/May/19
'Last update: overhaul to implement full dithering feature set; also add a bunch of quality and perf improvements
'
'"Posterizing" effect interface.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

Private Sub btsAdaptiveColoring_Click(ByVal buttonIndex As Long)
    UpdatePreview
End Sub

Private Sub cboDither_Click()
    UpdateDitherVisibility
    UpdatePreview
End Sub

Private Sub cmdBar_OKClick()
    Process "Posterize", , GetLocalParamString(), UNDO_Layer
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub Form_Load()
    
    cmdBar.SetPreviewStatus False
    
    'Use the palette module to populate our available dithering options
    Palettes.PopulateDitheringDropdown cboDither
    cboDither.ListIndex = 0
    UpdateDitherVisibility
    
    btsAdaptiveColoring.AddItem "off", 0
    btsAdaptiveColoring.AddItem "on", 1
    btsAdaptiveColoring.ListIndex = 0
    
    ApplyThemeAndTranslations Me, True, True
    cmdBar.SetPreviewStatus True
    UpdatePreview
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub UpdateDitherVisibility()
    sldDitherAmount.Visible = (cboDither.ListIndex <> 0)
End Sub

Private Sub UpdatePreview()
    If cmdBar.PreviewsAllowed Then Me.fxPosterize GetLocalParamString(), True, pdFxPreview
End Sub

Public Sub fxPosterize(ByVal effectParams As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)

    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString effectParams
    
    Dim ditherMethod As PD_DITHER_METHOD
    ditherMethod = cParams.GetLong("dithering", 0)
    
    Dim ditherAmount As Single
    ditherAmount = cParams.GetDouble("ditheramount", 100#) * 0.01
    
    'Retrieve the target pixel array
    If (Not toPreview) Then Message "Posterizing image..."
    Dim tmpSA2D As SafeArray2D
    EffectPrep.PrepImageData tmpSA2D, toPreview, dstPic
    
    'Dithering currently uses an old-school codebase, and thus requires a separate function
    With cParams
        If (ditherMethod = PDDM_None) Then
            Palettes.Palettize_BitRGB workingDIB, .GetLong("red"), .GetLong("green"), .GetLong("blue"), .GetBool("matchcolors", True), toPreview
        Else
            Palettes.Palettize_BitRGB_Dither workingDIB, .GetLong("red"), .GetLong("green"), .GetLong("blue"), ditherMethod, ditherAmount, .GetBool("matchcolors", True), toPreview
        End If
    End With
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    EffectPrep.FinalizeImageData toPreview, dstPic
    
End Sub

Private Sub sldDitherAmount_Change()
    UpdatePreview
End Sub

Private Sub sltBlue_Change()
    UpdatePreview
End Sub

Private Sub sltGreen_Change()
    UpdatePreview
End Sub

Private Sub sltRed_Change()
    UpdatePreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub pdFxPreview_ViewportChanged()
    UpdatePreview
End Sub

Private Function GetLocalParamString() As String
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    
    With cParams
        
        .AddParam "red", sltRed.Value
        .AddParam "green", sltGreen.Value
        .AddParam "blue", sltBlue.Value
        
        .AddParam "matchcolors", (btsAdaptiveColoring.ListIndex = 1)
        
        .AddParam "dithering", cboDither.ListIndex
        .AddParam "ditheramount", sldDitherAmount.Value
        
    End With
    
    GetLocalParamString = cParams.GetParamString()
    
End Function
