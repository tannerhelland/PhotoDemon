VERSION 5.00
Begin VB.Form FormNewLayer 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Add new layer"
   ClientHeight    =   6075
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
   ScaleHeight     =   405
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   642
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.pdDropDown cboPosition 
      Height          =   375
      Left            =   480
      TabIndex        =   8
      Top             =   4200
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   661
   End
   Begin PhotoDemon.pdTextBox txtLayerName 
      Height          =   345
      Left            =   480
      TabIndex        =   1
      Top             =   480
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   609
      FontSize        =   11
   End
   Begin PhotoDemon.pdCommandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5325
      Width           =   9630
      _ExtentX        =   16986
      _ExtentY        =   1323
   End
   Begin PhotoDemon.pdRadioButton optLayer 
      CausesValidation=   0   'False
      Height          =   330
      Index           =   0
      Left            =   480
      TabIndex        =   2
      Top             =   1560
      Width           =   8850
      _ExtentX        =   15610
      _ExtentY        =   582
      Caption         =   "transparent"
      Value           =   -1  'True
   End
   Begin PhotoDemon.pdRadioButton optLayer 
      CausesValidation=   0   'False
      Height          =   330
      Index           =   1
      Left            =   480
      TabIndex        =   3
      Top             =   1920
      Width           =   8850
      _ExtentX        =   15610
      _ExtentY        =   582
      Caption         =   "black"
   End
   Begin PhotoDemon.pdRadioButton optLayer 
      CausesValidation=   0   'False
      Height          =   330
      Index           =   2
      Left            =   480
      TabIndex        =   4
      Top             =   2280
      Width           =   8850
      _ExtentX        =   15610
      _ExtentY        =   582
      Caption         =   "white"
   End
   Begin PhotoDemon.pdRadioButton optLayer 
      CausesValidation=   0   'False
      Height          =   330
      Index           =   3
      Left            =   480
      TabIndex        =   5
      Top             =   2640
      Width           =   8850
      _ExtentX        =   15610
      _ExtentY        =   582
      Caption         =   "custom color"
   End
   Begin PhotoDemon.pdColorSelector csBackground 
      Height          =   615
      Left            =   840
      TabIndex        =   6
      Top             =   3000
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   1085
      curColor        =   16749332
   End
   Begin PhotoDemon.pdCheckBox chkAutoSelectLayer 
      Height          =   300
      Left            =   480
      TabIndex        =   7
      Top             =   4680
      Width           =   8820
      _ExtentX        =   6059
      _ExtentY        =   582
      Caption         =   "make the new layer the active layer"
   End
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   285
      Index           =   2
      Left            =   240
      Top             =   3840
      Width           =   9240
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "position"
      FontSize        =   12
      ForeColor       =   4210752
   End
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   285
      Index           =   1
      Left            =   240
      Top             =   1200
      Width           =   9165
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "background"
      FontSize        =   12
      ForeColor       =   4210752
   End
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   285
      Index           =   0
      Left            =   240
      Top             =   120
      Width           =   9105
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "name"
      FontSize        =   12
      ForeColor       =   4210752
   End
End
Attribute VB_Name = "FormNewLayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'New Layer dialog
'Copyright 2014-2026 by Tanner Helland
'Created: 08/June/14
'Last updated: 09/July/14
'Last update: added option for position and auto-selecting new layer
'
'Basic "add new layer" dialog.  Layer name, color, and position can be specified directly from the dialog,
' and the command bar allows for saving/loading presets just like every other tool.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

Private Sub cmdBar_OKClick()
    Process "Add new layer", False, GetLocalParamString(), UNDO_Image_VectorSafe
End Sub

Private Sub cmdBar_RandomizeClick()
    txtLayerName.Text = g_Language.TranslateMessage("Layer %1", Int(Rnd * 10000))
    optLayer(3).Value = True
End Sub

Private Sub cmdBar_ResetClick()
    txtLayerName.Text = g_Language.TranslateMessage("(enter name here)")
    csBackground.Color = RGB(60, 160, 255)
End Sub

Private Sub Form_Load()

    'Populate the position drop-down box
    cboPosition.Clear
    cboPosition.AddItem "default (above current layer)"
    cboPosition.AddItem "below current layer"
    cboPosition.AddItem "top of layer stack"
    cboPosition.AddItem "bottom of layer stack"
    cboPosition.ListIndex = 0

    ApplyThemeAndTranslations Me

End Sub

Private Function GetLocalParamString() As String
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    With cParams
        .AddParam "targetlayer", PDImages.GetActiveImage.GetActiveLayerIndex
        .AddParam "layertype", PDL_Image
        
        'Retrieve the layer type from the active radio button
        Dim newLayerType As Long, i As Long
        For i = 0 To optLayer.Count - 1
            If optLayer(i).Value Then
                newLayerType = i
                Exit For
            End If
        Next i
        
        .AddParam "layersubtype", newLayerType
        .AddParam "layercolor", csBackground.Color
        .AddParam "layerposition", cboPosition.ListIndex
        .AddParam "activatelayer", chkAutoSelectLayer.Value
        .AddParam "layername", txtLayerName
    End With
    
    GetLocalParamString = cParams.GetParamString()
    
End Function
