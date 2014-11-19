VERSION 5.00
Begin VB.Form FormThemeEditor 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Theme editor"
   ClientHeight    =   7185
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10605
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
   ScaleHeight     =   479
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   707
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.pdComboBox pdComboBox1 
      Height          =   360
      Left            =   4800
      TabIndex        =   7
      Top             =   6000
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   635
   End
   Begin VB.CommandButton cmdTextBoxTesting 
      Caption         =   "select all"
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
      Index           =   4
      Left            =   240
      TabIndex        =   6
      Top             =   5760
      Width           =   2055
   End
   Begin VB.CommandButton cmdTextBoxTesting 
      Caption         =   "toggle multiline"
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
      Index           =   3
      Left            =   2400
      TabIndex        =   5
      Top             =   5040
      Width           =   2055
   End
   Begin VB.CommandButton cmdTextBoxTesting 
      Caption         =   "append random chars"
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
      Index           =   2
      Left            =   240
      TabIndex        =   4
      Top             =   5040
      Width           =   2055
   End
   Begin PhotoDemon.pdLabel pdLabelVerify 
      Height          =   1695
      Left            =   4800
      Top             =   4200
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   2990
      Caption         =   ""
      FontSize        =   9
      Layout          =   1
   End
   Begin VB.CommandButton cmdTextBoxFake 
      Caption         =   "useless command button with TabIndex 0. (UC is TabIndex 1.)"
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
      Left            =   240
      TabIndex        =   0
      Top             =   6480
      Width           =   6135
   End
   Begin VB.CommandButton cmdTextBoxTesting 
      Caption         =   "resize edit box"
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
      Left            =   2400
      TabIndex        =   3
      Top             =   4320
      Width           =   2055
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
      TabIndex        =   2
      Top             =   4320
      Width           =   2055
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
      Height          =   3255
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   5741
      Multiline       =   -1  'True
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
   Begin PhotoDemon.pdLabel pdLabelTitle 
      Height          =   285
      Index           =   2
      Left            =   4800
      Top             =   3840
      Width           =   5715
      _ExtentX        =   6059
      _ExtentY        =   503
      Caption         =   "Unicode label for testing output:"
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

Private Sub cmdTextBoxFake_GotFocus()
    Debug.Print "Focus set to TabStop item " & cmdTextBoxFake.TabIndex
End Sub

Private Sub cmdTextBoxTesting_Click(Index As Integer)
    
    Select Case Index
        
        '0 - toggle visibility
        Case 0
            pdTextBox1.Visible = Not pdTextBox1.Visible
            
        '1 - resize text box
        Case 1
            Randomize Timer
            pdTextBox1.Move pdTextBox1.Left, pdTextBox1.Top, pdTextBox1.Width + ((Rnd * 10) - 5), pdTextBox1.Height + ((Rnd * 10) - 5)
            
        '2 - set random chars
        Case 2
            Dim tmpString As String
            tmpString = pdTextBox1.Text
            
            Dim i As Long
            For i = 0 To 40
                tmpString = tmpString & ChrW(Rnd * 2000)
            Next i
            
            pdTextBox1.Text = tmpString
            
        '3 - toggle multiline
        Case 3
            pdTextBox1.Multiline = Not pdTextBox1.Multiline
            
        '4 - select all
        Case 4
            pdTextBox1.SelectAll
        
    End Select
    
End Sub

Private Sub cmdTextBoxTesting_GotFocus(Index As Integer)
    Debug.Print "Focus set to TabStop item " & cmdTextBoxTesting(Index).TabIndex
End Sub

Private Sub Form_Activate()
    pdComboBox1.AddItem "1", 0
    pdComboBox1.AddItem "2", 1
    pdComboBox1.AddItem "3", 2
    pdComboBox1.ListIndex = 1
End Sub

Private Sub pdTextBox1_Change()
    pdLabelVerify.Caption = pdTextBox1.Text
End Sub

Private Sub pdTextBox1_GotFocus()
    Debug.Print "Focus set to UC - TabStop item " & pdTextBox1.TabIndex
End Sub

Private Sub pdTextBox1_LostFocus()
    Debug.Print "Focus lost from UC - TabStop item " & pdTextBox1.TabIndex
End Sub
