VERSION 5.00
Begin VB.Form FormShadowHighlight 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   " Shadows and highlights"
   ClientHeight    =   6570
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12120
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
   ScaleHeight     =   438
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   808
   Begin PhotoDemon.pdCommandBar cmdBar 
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5820
      Width           =   12120
      _ExtentX        =   21378
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
   Begin PhotoDemon.pdButtonStrip btsOptions 
      Height          =   960
      Left            =   6000
      TabIndex        =   2
      Top             =   4800
      Width           =   5910
      _ExtentX        =   10425
      _ExtentY        =   1693
      Caption         =   "options"
   End
   Begin PhotoDemon.pdContainer picContainer 
      Height          =   4575
      Index           =   0
      Left            =   5880
      Top             =   120
      Width           =   6135
      _ExtentX        =   0
      _ExtentY        =   0
      Begin PhotoDemon.pdSlider sltShadowAmount 
         Height          =   705
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   1270
         Caption         =   "shadows"
         Min             =   -100
         Max             =   100
      End
      Begin PhotoDemon.pdSlider sltHighlightAmount 
         Height          =   705
         Left            =   120
         TabIndex        =   7
         Top             =   2820
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   1270
         Caption         =   "highlights"
         Min             =   -100
         Max             =   100
      End
      Begin PhotoDemon.pdSlider sltMidtoneContrast 
         Height          =   705
         Left            =   120
         TabIndex        =   8
         Top             =   1770
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   1270
         Caption         =   "midtone contrast"
         Min             =   -100
         Max             =   100
      End
   End
   Begin PhotoDemon.pdContainer picContainer 
      Height          =   4575
      Index           =   1
      Left            =   5880
      Top             =   120
      Width           =   6135
      _ExtentX        =   0
      _ExtentY        =   0
      Begin PhotoDemon.pdSlider sltShadowWidth 
         Height          =   705
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   5700
         _ExtentX        =   10054
         _ExtentY        =   1270
         Caption         =   "tonal width"
         Max             =   100
         Value           =   75
         NotchPosition   =   2
         NotchValueCustom=   75
      End
      Begin PhotoDemon.pdSlider sltShadowRadius 
         Height          =   705
         Left            =   240
         TabIndex        =   3
         Top             =   1200
         Width           =   5700
         _ExtentX        =   10054
         _ExtentY        =   1270
         Caption         =   "radius"
         Max             =   200
         Value           =   25
         NotchPosition   =   2
         NotchValueCustom=   25
      End
      Begin PhotoDemon.pdSlider sltHighlightWidth 
         Height          =   705
         Left            =   240
         TabIndex        =   9
         Top             =   2760
         Width           =   5700
         _ExtentX        =   10054
         _ExtentY        =   1270
         Caption         =   "tonal width"
         Max             =   100
         Value           =   75
         NotchPosition   =   2
         NotchValueCustom=   75
      End
      Begin PhotoDemon.pdSlider sltHighlightRadius 
         Height          =   705
         Left            =   240
         TabIndex        =   5
         Top             =   3600
         Width           =   5700
         _ExtentX        =   10054
         _ExtentY        =   1270
         Caption         =   "radius"
         Max             =   200
         Value           =   25
         NotchPosition   =   2
         NotchValueCustom=   25
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   285
         Index           =   1
         Left            =   0
         Top             =   2340
         Width           =   5835
         _ExtentX        =   0
         _ExtentY        =   0
         Caption         =   "highlights"
         FontSize        =   12
         ForeColor       =   4210752
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   285
         Index           =   0
         Left            =   0
         Top             =   0
         Width           =   5715
         _ExtentX        =   0
         _ExtentY        =   0
         Caption         =   "shadows"
         FontSize        =   12
         ForeColor       =   4210752
      End
   End
End
Attribute VB_Name = "FormShadowHighlight"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Shadow / Midtone / Highlight Adjustment Tool
'Copyright 2013-2026 by Tanner Helland
'Created: 17/February/13
'Last updated: 20/July/17
'Last update: migrate to XML parameters
'
'This tool provides detailed control over the shadow and/or highlight regions of an image.  A combination of
' heuristics and user-editable parameters allow for brightening and/or darkening any luminance range in the
' source image.
'
'Note that the bulk of the image processing does not occur here, but in the separate AdjustDIBShadowHighlight
' function (currently inside the Filters_Layers module).
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

Private Sub btsOptions_Click(ByVal buttonIndex As Long)
    picContainer(buttonIndex).Visible = True
    picContainer(1 - buttonIndex).Visible = False
End Sub

'OK button
Private Sub cmdBar_OKClick()
    Process "Shadows and highlights", , GetLocalParamString(), UNDO_Layer
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub cmdBar_ResetClick()
    sltShadowWidth.Value = 75
    sltHighlightWidth.Value = 75
    sltShadowRadius.Value = 25
    sltHighlightRadius.Value = 25
End Sub

'Correct white balance by stretching the histogram and ignoring pixels above or below the 0.05% threshold
Public Sub ApplyShadowHighlight(ByVal effectParams As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)

    If (Not toPreview) Then Message "Adjusting shadows, midtones, and highlights..."
    
    Dim shadowAmount As Double, midtoneContrast As Double, highlightAmount As Double
    Dim shadowWidth As Long, shadowRadius As Double, highlightWidth As Long, highlightRadius As Double
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString effectParams
    
    With cParams
        shadowAmount = .GetDouble("shadowamount", sltShadowAmount)
        midtoneContrast = .GetDouble("midtonecontrast", sltMidtoneContrast)
        highlightAmount = .GetDouble("highlightamount", sltHighlightAmount)
        shadowWidth = .GetLong("shadowwidth", 50)
        shadowRadius = .GetDouble("shadowradius", 5#)
        highlightWidth = .GetLong("highlightwidth", 50)
        highlightRadius = .GetDouble("highlightradius", 5#)
    End With
    
    'Create a local array and point it at the pixel data of the current image
    Dim dstSA As SafeArray2D
    EffectPrep.PrepImageData dstSA, toPreview, dstPic
    
    AdjustDIBShadowHighlight shadowAmount, midtoneContrast, highlightAmount, shadowWidth, shadowRadius * curDIBValues.previewModifier, highlightWidth, highlightRadius * curDIBValues.previewModifier, workingDIB, toPreview
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering using the data inside workingDIB
    EffectPrep.FinalizeImageData toPreview, dstPic
    
End Sub

Private Sub Form_Load()
    
    cmdBar.SetPreviewStatus False
    
    'Set up the basic/advanced panels
    btsOptions.AddItem "basic", 0
    btsOptions.AddItem "advanced", 1
    btsOptions.ListIndex = 0
    btsOptions_Click 0
    
    'Apply translations and visual themes
    ApplyThemeAndTranslations Me, True, True
    cmdBar.SetPreviewStatus True
    UpdatePreview
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub UpdatePreview()
    If cmdBar.PreviewsAllowed Then Me.ApplyShadowHighlight GetLocalParamString(), True, pdFxPreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub pdFxPreview_ViewportChanged()
    UpdatePreview
End Sub

Private Sub sltHighlightAmount_Change()
    UpdatePreview
End Sub

Private Sub sltHighlightRadius_Change()
    UpdatePreview
End Sub

Private Sub sltHighlightWidth_Change()
    UpdatePreview
End Sub

Private Sub sltMidtoneContrast_Change()
    UpdatePreview
End Sub

Private Sub sltShadowAmount_Change()
    UpdatePreview
End Sub

Private Sub sltShadowRadius_Change()
    UpdatePreview
End Sub

Private Sub sltShadowWidth_Change()
    UpdatePreview
End Sub

Private Function GetLocalParamString() As String
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    
    With cParams
        .AddParam "shadowamount", sltShadowAmount
        .AddParam "midtonecontrast", sltMidtoneContrast
        .AddParam "highlightamount", sltHighlightAmount
        .AddParam "shadowwidth", sltShadowWidth
        .AddParam "shadowradius", sltShadowRadius
        .AddParam "highlightwidth", sltHighlightWidth
        .AddParam "highlightradius", sltHighlightRadius
    End With
    
    GetLocalParamString = cParams.GetParamString()
    
End Function
