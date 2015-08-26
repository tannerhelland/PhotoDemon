VERSION 5.00
Begin VB.Form FormNewLayer 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Add new layer"
   ClientHeight    =   6075
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9630
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
   ScaleHeight     =   405
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   642
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin PhotoDemon.pdTextBox txtLayerName 
      Height          =   345
      Left            =   480
      TabIndex        =   11
      Top             =   480
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   609
      FontSize        =   11
   End
   Begin VB.ComboBox cboPosition 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   330
      Left            =   480
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   4200
      Width           =   8895
   End
   Begin PhotoDemon.commandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5325
      Width           =   9630
      _ExtentX        =   16986
      _ExtentY        =   1323
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14802140
   End
   Begin PhotoDemon.smartOptionButton optLayer 
      CausesValidation=   0   'False
      Height          =   330
      Index           =   0
      Left            =   480
      TabIndex        =   3
      Top             =   1560
      Width           =   8850
      _ExtentX        =   15610
      _ExtentY        =   582
      Caption         =   "transparent"
      Value           =   -1  'True
   End
   Begin PhotoDemon.smartOptionButton optLayer 
      CausesValidation=   0   'False
      Height          =   330
      Index           =   1
      Left            =   480
      TabIndex        =   4
      Top             =   1920
      Width           =   8850
      _ExtentX        =   15610
      _ExtentY        =   582
      Caption         =   "black"
   End
   Begin PhotoDemon.smartOptionButton optLayer 
      CausesValidation=   0   'False
      Height          =   330
      Index           =   2
      Left            =   480
      TabIndex        =   5
      Top             =   2280
      Width           =   8850
      _ExtentX        =   15610
      _ExtentY        =   582
      Caption         =   "white"
   End
   Begin PhotoDemon.smartOptionButton optLayer 
      CausesValidation=   0   'False
      Height          =   330
      Index           =   3
      Left            =   480
      TabIndex        =   6
      Top             =   2640
      Width           =   8850
      _ExtentX        =   15610
      _ExtentY        =   582
      Caption         =   "custom color"
   End
   Begin PhotoDemon.colorSelector colorPicker 
      Height          =   615
      Left            =   840
      TabIndex        =   7
      Top             =   3000
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   1085
      curColor        =   16749332
   End
   Begin PhotoDemon.smartCheckBox chkAutoSelectLayer 
      Height          =   300
      Left            =   480
      TabIndex        =   10
      Top             =   4680
      Width           =   8820
      _ExtentX        =   6059
      _ExtentY        =   582
      Caption         =   "make the new layer the active layer"
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "position"
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
      Index           =   2
      Left            =   240
      TabIndex        =   8
      Top             =   3840
      Width           =   840
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "background"
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
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   1245
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "name"
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
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   585
   End
End
Attribute VB_Name = "FormNewLayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'New Layer dialog
'Copyright 2014-2015 by Tanner Helland
'Created: 08/June/14
'Last updated: 09/July/14
'Last update: added option for position and auto-selecting new layer
'
'Basic "add new layer" dialog.  Layer name, color, and position can be specified directly from the dialog,
' and the command bar allows for saving/loading presets just like every other tool.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

Private Sub cmdBar_OKClick()

    'Retrieve the layer type from the active command button
    Dim newLayerType As Long
    
    Dim i As Long
    For i = 0 To optLayer.Count - 1
        If optLayer(i).Value Then
            newLayerType = i
            Exit For
        End If
    Next i
    
    Process "Add new layer", False, buildParams(pdImages(g_CurrentImage).getActiveLayerIndex, PDL_IMAGE, newLayerType, colorPicker.Color, cboPosition.ListIndex, CBool(chkAutoSelectLayer), txtLayerName), UNDO_IMAGE_VECTORSAFE
    
End Sub

Private Sub cmdBar_RandomizeClick()
    txtLayerName.Text = g_Language.TranslateMessage("Layer %1", Int(Rnd * 10000))
    optLayer(3).Value = True
End Sub

Private Sub cmdBar_ResetClick()
    txtLayerName.Text = g_Language.TranslateMessage("(enter name here)")
    colorPicker.Color = RGB(60, 160, 255)
End Sub

Private Sub Form_Activate()
    makeFormPretty Me
End Sub

Private Sub Form_Load()

    'Populate the position drop-down box
    cboPosition.Clear
    cboPosition.AddItem "default (above current layer)"
    cboPosition.AddItem "below current layer"
    cboPosition.AddItem "top of layer stack"
    cboPosition.AddItem "bottom of layer stack"
    cboPosition.ListIndex = 0

End Sub
