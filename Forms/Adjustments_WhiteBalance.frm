VERSION 5.00
Begin VB.Form FormWhiteBalance 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   " White balance"
   ClientHeight    =   6510
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11655
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
   ScaleHeight     =   434
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   777
   Begin PhotoDemon.pdCommandBar cmdBar 
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5760
      Width           =   11655
      _ExtentX        =   20558
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
   Begin PhotoDemon.pdSlider sltStrength 
      Height          =   705
      Left            =   6000
      TabIndex        =   2
      Top             =   2520
      Width           =   5445
      _ExtentX        =   9604
      _ExtentY        =   1244
      Caption         =   "strength"
      Min             =   0.01
      Max             =   5
      SigDigits       =   2
      Value           =   0.05
      DefaultValue    =   0.05
   End
End
Attribute VB_Name = "FormWhiteBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'White Balance Handler
'Copyright 2012-2026 by Tanner Helland
'Created: 03/July/12
'Last updated: 24/August/13
'Last update: added command bar
'
'White balance handler.  Unlike other programs, which shove this under the Levels dialog as an "auto levels"
' function, I consider it worthy of its own interface.  The reason is - white balance is an important function.
' It's arguably more useful than the Levels dialog, especially to a casual user, because it automatically
' calculates levels according to a reliable, often-accurate algorithm.  Rather than forcing the user through the
' Levels dialog (because really, how many people know that Auto Levels is actually White Balance in photography
' parlance?), PhotoDemon provides a full implementation of custom white balance handling.
'
'The value box on the form is the percentage of pixels ignored at the top and bottom of the histogram.
' 0.05 is the recommended default.  I've specified 5.0 as the maximum, but there's no reason it couldn't be set
' higher... just be forewarned that higher values (obviously) blow out the picture with increasing strength.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'OK button
Private Sub cmdBar_OKClick()
    Process "White balance", , GetLocalParamString(), UNDO_Layer
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub cmdBar_ResetClick()
    sltStrength.Value = 0.05
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

Private Sub sltStrength_Change()
    UpdatePreview
End Sub

Private Sub UpdatePreview()
    If cmdBar.PreviewsAllowed Then Filters_Adjustments.AutoWhiteBalance GetLocalParamString(), True, pdFxPreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub pdFxPreview_ViewportChanged()
    UpdatePreview
End Sub

Private Function GetLocalParamString() As String
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.AddParam "threshold", sltStrength.Value
    
    GetLocalParamString = cParams.GetParamString()
    
End Function
