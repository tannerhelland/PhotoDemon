VERSION 5.00
Begin VB.Form FormConvert24bpp 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   " Remove alpha channel"
   ClientHeight    =   6540
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11310
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
   ScaleWidth      =   754
   Begin PhotoDemon.pdCommandBar cmdBar 
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5790
      Width           =   11310
      _ExtentX        =   19950
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
   Begin PhotoDemon.pdColorSelector csBackground 
      Height          =   1215
      Left            =   6000
      TabIndex        =   2
      Top             =   2160
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   2143
      Caption         =   "background color:"
   End
End
Attribute VB_Name = "FormConvert24bpp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Convert image to 24bpp (remove alpha channel) interface
'Copyright 2013-2026 by Tanner Helland
'Created: 14/June/13
'Last updated: 14/June/13
'Last update: initial build
'
'PhotoDemon has long provided the ability to convert a 32bpp image to 24bpp, but the lack of an interface meant it would
' always composite against white.  Now the user can select any background color they want.
'
'Other than that, there really isn't much to this form.  It's possibly the smallest tool dialog in PhotoDemon code-wise!
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

Private Sub cmdBar_OKClick()
    Process "Remove alpha channel", , GetLocalParamString(), UNDO_Layer
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub csBackground_ColorChanged()
    UpdatePreview
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

Private Sub pdFxPreview_ColorSelected()
    csBackground.Color = pdFxPreview.SelectedColor
    UpdatePreview
End Sub

'Render a new preview
Private Sub UpdatePreview()
    
    If cmdBar.PreviewsAllowed Then
    
        Dim tmpSA As SafeArray2D
        EffectPrep.PrepImageData tmpSA, True, pdFxPreview
        
        Dim newBackColor As Long
        newBackColor = csBackground.Color
        workingDIB.CompositeBackgroundColor Colors.ExtractRed(newBackColor), Colors.ExtractGreen(newBackColor), Colors.ExtractBlue(newBackColor)
        
        EffectPrep.FinalizeImageData True, pdFxPreview, True
        
    End If
    
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub pdFxPreview_ViewportChanged()
    UpdatePreview
End Sub

Public Sub RemoveLayerTransparency(ByVal processParameters As String)
    
    Message "Removing transparency..."
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString processParameters
    
    Dim newBackColor As Long
    newBackColor = cParams.GetLong("backcolor", vbWhite)
    
    'Ask the current DIB to convert itself to 24bpp mode
    PDImages.GetActiveImage.GetActiveDIB.CompositeBackgroundColor Colors.ExtractRed(newBackColor), Colors.ExtractGreen(newBackColor), Colors.ExtractBlue(newBackColor)
    PDImages.GetActiveImage.GetActiveDIB.SetInitialAlphaPremultiplicationState True
    
    'Notify the parent of the target layer of the change
    PDImages.GetActiveImage.NotifyImageChanged UNDO_Layer, PDImages.GetActiveImage.GetActiveLayerIndex
    
    Message "Finished."
    
    'Redraw the main window
    Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    
End Sub

Private Function GetLocalParamString() As String
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    
    With cParams
        .AddParam "backcolor", csBackground.Color
    End With
    
    GetLocalParamString = cParams.GetParamString()
    
End Function
