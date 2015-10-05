VERSION 5.00
Begin VB.Form FormConvert24bpp 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Remove alpha channel"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11820
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
   ScaleWidth      =   788
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.commandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5790
      Width           =   11820
      _ExtentX        =   20849
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
      ColorSelection  =   -1  'True
   End
   Begin PhotoDemon.colorSelector colorPicker 
      Height          =   1215
      Left            =   6000
      TabIndex        =   2
      Top             =   2160
      Width           =   5535
      _ExtentX        =   9763
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
'Copyright 2013-2015 by Tanner Helland
'Created: 14/June/13
'Last updated: 14/June/13
'Last update: initial build
'
'PhotoDemon has long provided the ability to convert a 32bpp image to 24bpp, but the lack of an interface meant it would
' always composite against white.  Now the user can select any background color they want.
'
'Other than that, there really isn't much to this form.  It's possibly the smallest tool dialog in PhotoDemon code-wise!
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

Private Sub cmdBar_OKClick()
    Process "Remove alpha channel", , buildParams(colorPicker.Color), UNDO_LAYER
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    updatePreview
End Sub

Private Sub colorPicker_ColorChanged()
    updatePreview
End Sub

Private Sub Form_Activate()
        
    'Apply translations and visual themes
    MakeFormPretty Me
    
    'Render a preview of the emboss/engrave effect
    updatePreview
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub fxPreview_ColorSelected()
    colorPicker.Color = fxPreview.SelectedColor
    updatePreview
End Sub

'Render a new preview
Private Sub updatePreview()
    If cmdBar.previewsAllowed Then
        Dim tmpSA As SAFEARRAY2D
        prepImageData tmpSA, True, fxPreview
        workingDIB.convertTo24bpp colorPicker.Color
        finalizeImageData True, fxPreview
    End If
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub fxPreview_ViewportChanged()
    updatePreview
End Sub


