VERSION 5.00
Begin VB.Form FormNewImage 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Create new image"
   ClientHeight    =   8160
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
   ScaleHeight     =   544
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   642
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin PhotoDemon.pdComboBox cboTemplates 
      Height          =   360
      Left            =   480
      TabIndex        =   8
      Top             =   480
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   635
   End
   Begin PhotoDemon.commandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   7410
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
   End
   Begin PhotoDemon.smartOptionButton optLayer 
      CausesValidation=   0   'False
      Height          =   330
      Index           =   0
      Left            =   480
      TabIndex        =   2
      Top             =   4800
      Width           =   8820
      _ExtentX        =   2434
      _ExtentY        =   582
      Caption         =   "transparent"
      Value           =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PhotoDemon.smartOptionButton optLayer 
      CausesValidation=   0   'False
      Height          =   330
      Index           =   1
      Left            =   480
      TabIndex        =   3
      Top             =   5160
      Width           =   8820
      _ExtentX        =   1455
      _ExtentY        =   582
      Caption         =   "black"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PhotoDemon.smartOptionButton optLayer 
      CausesValidation=   0   'False
      Height          =   330
      Index           =   2
      Left            =   480
      TabIndex        =   4
      Top             =   5520
      Width           =   8820
      _ExtentX        =   1508
      _ExtentY        =   582
      Caption         =   "white"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PhotoDemon.smartOptionButton optLayer 
      CausesValidation=   0   'False
      Height          =   330
      Index           =   3
      Left            =   480
      TabIndex        =   5
      Top             =   5880
      Width           =   8820
      _ExtentX        =   2752
      _ExtentY        =   582
      Caption         =   "custom color:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PhotoDemon.colorSelector colorPicker 
      Height          =   615
      Left            =   840
      TabIndex        =   6
      Top             =   6240
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   1085
      curColor        =   16749332
   End
   Begin PhotoDemon.smartResize ucResize 
      Height          =   2850
      Left            =   360
      TabIndex        =   9
      Top             =   1440
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   5027
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DisablePercentOption=   -1  'True
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "size:"
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
      TabIndex        =   10
      Top             =   1080
      Width           =   480
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "templates:"
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
      TabIndex        =   7
      Top             =   120
      Width           =   1110
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "background:"
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
      TabIndex        =   1
      Top             =   4440
      Width           =   1335
   End
End
Attribute VB_Name = "FormNewImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'New Image Dialog
'Copyright ©2014-2014 by Tanner Helland
'Created: 29/December/14
'Last updated: 29/December/14
'Last update: initial build
'
'Basic "create new image" dialog.  Image size and background can be specified directly from the dialog,
' and the command bar allows for saving/loading presets just like every other tool.  A few templates are
' provided for convenience.
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
    
    'Process "Add new layer", False, buildParams(pdImages(g_CurrentImage).getActiveLayerIndex, newLayerType, colorPicker.Color, cboPosition.ListIndex, CBool(chkAutoSelectLayer), txtLayerName), UNDO_IMAGE
    
End Sub

Private Sub cmdBar_RandomizeClick()
    'txtLayerName.Text = g_Language.TranslateMessage("Layer %1", Int(Rnd * 10000))
    optLayer(3).Value = True
End Sub

Private Sub cmdBar_ResetClick()
    'txtLayerName.Text = g_Language.TranslateMessage("(enter name here)")
    colorPicker.Color = RGB(60, 160, 255)
End Sub

Private Sub Form_Load()

    'Populate the position drop-down box
    'cboPosition.Clear
    'cboPosition.AddItem "default (above current layer)"
    'cboPosition.AddItem "below current layer"
    'cboPosition.AddItem "top of layer stack"
    'cboPosition.AddItem "bottom of layer stack"
    'cboPosition.ListIndex = 0

End Sub
