VERSION 5.00
Begin VB.Form FormRelief 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   " Relief"
   ClientHeight    =   6540
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11550
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
   ScaleWidth      =   770
   Begin PhotoDemon.pdCommandBar cmdBar 
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5790
      Width           =   11550
      _ExtentX        =   20373
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
   Begin PhotoDemon.pdSlider sltDistance 
      Height          =   705
      Left            =   6000
      TabIndex        =   2
      Top             =   2400
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   1244
      Caption         =   "thickness"
      Min             =   -10
      SigDigits       =   2
      Value           =   1
      DefaultValue    =   1
   End
   Begin PhotoDemon.pdSlider sltAngle 
      Height          =   705
      Left            =   6000
      TabIndex        =   3
      Top             =   1320
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   1244
      Caption         =   "angle"
      Min             =   -180
      Max             =   180
      SigDigits       =   1
   End
   Begin PhotoDemon.pdSlider sltDepth 
      Height          =   705
      Left            =   6000
      TabIndex        =   4
      Top             =   3480
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   1244
      Caption         =   "depth"
      Min             =   0.1
      SigDigits       =   2
      Value           =   1
      DefaultValue    =   1
   End
End
Attribute VB_Name = "FormRelief"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Relief Artistic Effect Dialog
'Copyright 2003-2026 by Tanner Helland
'Created: sometime 2003
'Last updated: 21/February/21
'Last update: large performance improvements
'
'This dialog applied a relief-style filter to an image.  Some kind of relief filter has existed
' in PD for a long time, but the 6.4 release saw much-needed improvements in the form of selectable
' angle, depth, and thickness.  Interpolation is used to process all relief calculations, so the
' result looks great for any angle and/or depth combination.  Edge handling is also handled much
' better than past versions.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'OK button
Private Sub cmdBar_OKClick()
    Process "Relief", , GetLocalParamString(), UNDO_Layer
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub cmdBar_ResetClick()
    sltDepth.Value = 1
    sltDistance.Value = 1
End Sub

Private Sub Form_Load()
    cmdBar.SetPreviewStatus False
    ApplyThemeAndTranslations Me, True, True
    cmdBar.SetPreviewStatus True
    UpdatePreview
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'Apply a relief filter, which gives the image a pseudo-3D appearance
Public Sub ApplyReliefEffect(ByVal effectParams As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)
    Filters_Edge.Filter_Edge_Relief effectParams, Nothing, toPreview, dstPic
End Sub

'Render a new preview
Private Sub UpdatePreview()
    If cmdBar.PreviewsAllowed Then ApplyReliefEffect GetLocalParamString(), True, pdFxPreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub pdFxPreview_ViewportChanged()
    UpdatePreview
End Sub

Private Sub sltAngle_Change()
    UpdatePreview
End Sub

Private Sub sltDepth_Change()
    UpdatePreview
End Sub

Private Sub sltDistance_Change()
    UpdatePreview
End Sub

Private Function GetLocalParamString() As String
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    
    With cParams
        .AddParam "distance", sltDistance.Value
        .AddParam "angle", sltAngle.Value
        .AddParam "depth", sltDepth.Value
    End With
    
    GetLocalParamString = cParams.GetParamString()
    
End Function
