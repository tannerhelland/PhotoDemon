VERSION 5.00
Begin VB.Form frmUpdate 
   BackColor       =   &H80000005&
   Caption         =   "PhotoDemon Update Generator"
   ClientHeight    =   6915
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12630
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   461
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   842
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAction 
      Caption         =   "Assemble nightly build files"
      Height          =   615
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   12135
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Step 1: copy all relevant nightly build files into a dedicated /nightly folder"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00DC7032&
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8040
   End
End
Attribute VB_Name = "frmUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAction_Click(Index As Integer)
    
    Select Case Index
    
        'Assemble nightly build files
        Case 0
            AssembleNightlyBuild
            
    
    End Select
    
End Sub

'Copy the relevant nightly build files from their default VB project location, to a dedicated /Nightly folder.
' This greatly simplifies the pdPackage generation step, as we can handle the dedicated /Nightly folder the same way
' we handle the /Stable and /Beta folders.
Private Sub AssembleNightlyBuild()

    'This list of relevant files is hardcoded.

End Sub
