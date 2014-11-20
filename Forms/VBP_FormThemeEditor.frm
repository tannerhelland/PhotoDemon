VERSION 5.00
Begin VB.Form FormThemeEditor 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Theme editor"
   ClientHeight    =   7365
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
   ScaleHeight     =   491
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   707
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.pdComboBox pdComboBox1 
      Height          =   360
      Left            =   4800
      TabIndex        =   7
      Top             =   6000
      Width           =   5655
      _extentx        =   9975
      _extenty        =   635
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
      _extentx        =   9975
      _extenty        =   2990
      caption         =   ""
      fontsize        =   9
      layout          =   1
   End
   Begin VB.CommandButton cmdTextBoxFake 
      Caption         =   "useless command button with TabIndex 0"
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
      Width           =   4215
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
      _extentx        =   5345
      _extenty        =   503
      caption         =   "Unicode text box for testing:"
      fontsize        =   12
      layout          =   2
   End
   Begin PhotoDemon.pdTextBox pdTextBox1 
      Height          =   2055
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   9975
      _extentx        =   17595
      _extenty        =   5741
      multiline       =   -1  'True
   End
   Begin PhotoDemon.pdLabel pdLabelTitle 
      Height          =   285
      Index           =   1
      Left            =   120
      Top             =   3840
      Width           =   2580
      _extentx        =   4551
      _extenty        =   503
      caption         =   "text box testing options:"
      fontsize        =   12
      layout          =   2
   End
   Begin PhotoDemon.pdLabel pdLabelTitle 
      Height          =   285
      Index           =   2
      Left            =   4800
      Top             =   3840
      Width           =   5715
      _extentx        =   6059
      _extenty        =   503
      caption         =   "Unicode label for testing output:"
      fontsize        =   12
      layout          =   2
   End
   Begin PhotoDemon.pdTextBox pdTextBox2 
      Height          =   315
      Index           =   0
      Left            =   240
      TabIndex        =   8
      Top             =   2640
      Width           =   9975
      _extentx        =   17595
      _extenty        =   556
   End
   Begin PhotoDemon.pdTextBox pdTextBox2 
      Height          =   315
      Index           =   1
      Left            =   240
      TabIndex        =   9
      Top             =   3000
      Width           =   9975
      _extentx        =   17595
      _extenty        =   556
   End
   Begin PhotoDemon.pdTextBox pdTextBox2 
      Height          =   315
      Index           =   2
      Left            =   240
      TabIndex        =   10
      Top             =   3360
      Width           =   9975
      _extentx        =   17595
      _extenty        =   556
   End
   Begin PhotoDemon.pdComboBox pdComboBox2 
      Height          =   360
      Index           =   0
      Left            =   4800
      TabIndex        =   11
      Top             =   6480
      Width           =   5655
      _extentx        =   9975
      _extenty        =   635
   End
   Begin PhotoDemon.pdComboBox pdComboBox2 
      Height          =   360
      Index           =   1
      Left            =   4800
      TabIndex        =   12
      Top             =   6960
      Width           =   5655
      _extentx        =   9975
      _extenty        =   635
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
