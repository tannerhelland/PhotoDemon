VERSION 5.00
Begin VB.Form FormNewLayer 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Add new layer"
   ClientHeight    =   5340
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
   ScaleHeight     =   356
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   642
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtLayerName 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   480
      TabIndex        =   2
      Text            =   "(enter name here)"
      Top             =   600
      Width           =   8895
   End
   Begin PhotoDemon.commandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   4590
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
      Height          =   375
      Index           =   0
      Left            =   480
      TabIndex        =   4
      Top             =   1800
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   661
      Caption         =   "transparent"
      Value           =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PhotoDemon.smartOptionButton optLayer 
      CausesValidation=   0   'False
      Height          =   375
      Index           =   1
      Left            =   480
      TabIndex        =   5
      Top             =   2280
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   661
      Caption         =   "black"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PhotoDemon.smartOptionButton optLayer 
      CausesValidation=   0   'False
      Height          =   375
      Index           =   2
      Left            =   480
      TabIndex        =   6
      Top             =   2760
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   661
      Caption         =   "white"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PhotoDemon.smartOptionButton optLayer 
      CausesValidation=   0   'False
      Height          =   375
      Index           =   3
      Left            =   480
      TabIndex        =   7
      Top             =   3240
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   661
      Caption         =   "custom color:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
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
      TabIndex        =   8
      Top             =   3720
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   1085
      curColor        =   16749332
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "new layer type:"
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
      TabIndex        =   3
      Top             =   1320
      Width           =   1635
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "new layer name:"
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
      Width           =   1770
   End
End
Attribute VB_Name = "FormNewLayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'New Layer dialog
'Copyright ©2013-2014 by Tanner Helland
'Created: 08/June/14
'Last updated: 09/June/14
'Last update: wrapped up initial build
'
'Basic "add new layer" dialog.  Layer name and color can be specified directly from the dialog, and the command bar
' allows for saving/loading presets just like every other tool.
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
    
    Process "Add new layer", False, buildParams(pdImages(g_CurrentImage).getActiveLayerIndex, newLayerType, colorPicker.Color, txtLayerName), UNDO_IMAGE
    
End Sub

Private Sub cmdBar_RandomizeClick()
    txtLayerName.Text = g_Language.TranslateMessage("Layer %1", Int(Rnd * 10000))
    optLayer(3).Value = True
End Sub

Private Sub cmdBar_ResetClick()
    txtLayerName.Text = g_Language.TranslateMessage("(enter name here)")
    colorPicker.Color = RGB(60, 160, 255)
End Sub

Private Sub txtLayerName_GotFocus()
    AutoSelectText txtLayerName
End Sub
