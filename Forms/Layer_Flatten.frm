VERSION 5.00
Begin VB.Form FormLayerFlatten 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Flatten image"
   ClientHeight    =   3300
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9630
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
   ScaleHeight     =   220
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   642
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.pdColorSelector clsBackground 
      Height          =   975
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   1931
      Caption         =   "background color"
   End
   Begin PhotoDemon.pdButtonStrip btsFlatten 
      Height          =   1095
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   1931
      Caption         =   "flatten behavior"
      FontSize        =   11
   End
   Begin PhotoDemon.pdCommandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   2550
      Width           =   9630
      _ExtentX        =   16986
      _ExtentY        =   1323
   End
End
Attribute VB_Name = "FormLayerFlatten"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Flatten Image Dialog
'Copyright 2015-2026 by Tanner Helland
'Created: 20/May/16
'Last updated: 20/May/16
'Last update: provide dialog for setting flatten options
'
'PD has supported flattening as long as it's supported layers.  However, there is some confusion over how a
' "Flatten" option should work.  Some software follows the Photoshop convention, where Flatten always replaces
' transparency with white, and there are no settings to change this.  Other software follows the Paint.NET
' convention, where transparency is preserved in the flattened image.
'
'PD originally defaulted to the Paint.NET model, but to reduce confusion, I've since added this dialog,
' so the user can specify exactly which flatten behavior they prefer.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

Private Sub btsFlatten_Click(ByVal buttonIndex As Long)
    UpdateComponentVisibility
End Sub

Private Sub UpdateComponentVisibility()
    clsBackground.Visible = (btsFlatten.ListIndex = 2)
End Sub

'OK button
Private Sub cmdBar_OKClick()
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    
    With cParams
        .AddParam "removetransparency", (btsFlatten.ListIndex = 2)
        .AddParam "backgroundcolor", clsBackground.Color
    End With
    
    Process "Flatten image", , cParams.GetParamString, UNDO_Image
    
End Sub

Private Sub cmdBar_ResetClick()
    btsFlatten.ListIndex = 0
    clsBackground.Color = vbWhite
End Sub

Private Sub Form_Load()
    
    'Populate the primary button strip
    btsFlatten.AddItem "auto", 0
    btsFlatten.AddItem "keep transparency", 1
    btsFlatten.AddItem "remove transparency", 2
    btsFlatten.ListIndex = 0
    UpdateComponentVisibility
    
    'Apply translations and visual themes
    ApplyThemeAndTranslations Me
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub
