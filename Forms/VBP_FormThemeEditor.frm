VERSION 5.00
Begin VB.Form FormThemeEditor 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Theme editor"
   ClientHeight    =   7050
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10290
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
   ScaleHeight     =   470
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   686
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdTextBoxTesting 
      Caption         =   "resize control"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   3480
      TabIndex        =   2
      Top             =   4320
      Width           =   3015
   End
   Begin VB.CommandButton cmdTextBoxTesting 
      Caption         =   "toggle visibility"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   4320
      Width           =   3015
   End
   Begin PhotoDemon.pdLabel pdLabelTitle 
      Height          =   285
      Index           =   0
      Left            =   120
      Top             =   120
      Width           =   3030
      _ExtentX        =   5345
      _ExtentY        =   503
      Caption         =   "Unicode text box for testing:"
      FontSize        =   12
      Layout          =   2
   End
   Begin PhotoDemon.pdTextBox pdTextBox1 
      Height          =   2175
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   3836
   End
   Begin PhotoDemon.pdLabel pdLabelTitle 
      Height          =   285
      Index           =   1
      Left            =   120
      Top             =   3840
      Width           =   2580
      _ExtentX        =   4551
      _ExtentY        =   503
      Caption         =   "text box testing options:"
      FontSize        =   12
      Layout          =   2
   End
End
Attribute VB_Name = "FormThemeEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdTextBoxTesting_Click(Index As Integer)
    
    Select Case Index
        
        '0 - toggle visibility
        Case 0
            pdTextBox1.Visible = Not pdTextBox1.Visible
            
        '1 - resize text box
        Case 1
            Randomize Timer
            pdTextBox1.Move pdTextBox1.Left, pdTextBox1.Top, pdTextBox1.Width + ((Rnd * 10) - 5), pdTextBox1.Height + ((Rnd * 10) - 5)
        
    End Select
    
End Sub
