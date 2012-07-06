VERSION 5.00
Begin VB.Form FormPCX 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Save PCX Options"
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5025
   BeginProperty Font 
      Name            =   "Arial"
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
   ScaleHeight     =   190
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox ChkOptimize 
      Appearance      =   0  'Flat
      Caption         =   "Optimize Palette"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   2640
      MouseIcon       =   "VBP_FormPCX.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   960
      Value           =   1  'Checked
      Width           =   1935
   End
   Begin VB.CheckBox ChkRLE 
      Appearance      =   0  'Flat
      Caption         =   "RLE Compressed"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   2640
      MouseIcon       =   "VBP_FormPCX.frx":0152
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   600
      Value           =   1  'Checked
      Width           =   1935
   End
   Begin VB.OptionButton OptColorDepth 
      Appearance      =   0  'Flat
      Caption         =   "1-bit (2 colors)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   0
      Left            =   360
      MouseIcon       =   "VBP_FormPCX.frx":02A4
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Tag             =   "1"
      Top             =   600
      Width           =   1455
   End
   Begin VB.OptionButton OptColorDepth 
      Appearance      =   0  'Flat
      Caption         =   "4-bit (16 colors)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   1
      Left            =   360
      MouseIcon       =   "VBP_FormPCX.frx":03F6
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Tag             =   "4"
      Top             =   960
      Width           =   1455
   End
   Begin VB.OptionButton OptColorDepth 
      Appearance      =   0  'Flat
      Caption         =   "8-bit (256 colors)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   2
      Left            =   360
      MouseIcon       =   "VBP_FormPCX.frx":0548
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Tag             =   "8"
      Top             =   1320
      Width           =   2055
   End
   Begin VB.OptionButton OptColorDepth 
      Appearance      =   0  'Flat
      Caption         =   "24-bit (16 million colors)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   3
      Left            =   360
      MouseIcon       =   "VBP_FormPCX.frx":069A
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Tag             =   "24"
      Top             =   1680
      Value           =   -1  'True
      Width           =   2055
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      MouseIcon       =   "VBP_FormPCX.frx":07EC
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   2280
      Width           =   1125
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      MouseIcon       =   "VBP_FormPCX.frx":093E
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   2280
      Width           =   1125
   End
   Begin VB.Label lblOptions 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Additional Options:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   2520
      TabIndex        =   9
      Top             =   240
      Width           =   1590
   End
   Begin VB.Label lblColorDepth 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Color Depth:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   240
      TabIndex        =   8
      Top             =   240
      Width           =   1035
   End
End
Attribute VB_Name = "FormPCX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'PCX Export Interface
'©2000-2012 Tanner Helland
'Created: 6/16/06
'Last updated: 05/July/12
'Last update: code changes to match interface redesign
'
'Form for allowing the user to set some PCX export options.  Interacts heavily
' with Alfred Koppold's "SavePCX" class
'
'***************************************************************************

Option Explicit

Dim cdIndex As Long

'CANCEL button
Private Sub CmdCancel_Click()
    Unload Me
End Sub

'OK button
Private Sub CmdOK_Click()
    Me.Visible = False
    Message "Preparing image..."
    If ChkOptimize.Enabled = True And ChkOptimize.Value = 1 Then
        Message "Optimizing palette..."
        Select Case cdIndex
        Case 0 'Black and white
            Process BWFloydSteinberg
        Case 1 '16 color palette
            Process ReduceColors, REDUCECOLORS_MANUAL_ERRORDIFFUSION, 2, 3, 2, True
        Case 2 '256 color palette
            If FreeImageEnabled = True Then
                Process ReduceColors, REDUCECOLORS_AUTO, FIQ_WUQUANT
            Else
                Process ReduceColors, REDUCECOLORS_MANUAL_ERRORDIFFUSION, 6, 7, 6, True
            End If
        End Select
    End If
    
    PhotoDemon_SaveImage CurrentImage, SaveFileName, False, CLng(OptColorDepth(cdIndex).Tag), CLng(ChkRLE.Value)

End Sub

Private Sub Form_Load()
    cdIndex = 3
End Sub

Private Sub OptColorDepth_Click(Index As Integer)
    cdIndex = Index
    If Index = 3 Then ChkOptimize.Enabled = False Else ChkOptimize.Enabled = True
End Sub
