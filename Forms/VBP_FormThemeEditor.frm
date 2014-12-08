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
   Begin VB.CommandButton cmdTextBoxTesting 
      Caption         =   "add new drop-down entries"
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
      TabIndex        =   13
      Top             =   5760
      Width           =   2055
   End
   Begin VB.CommandButton cmdTextBoxTesting 
      Caption         =   "delete drop-down entries"
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
      Index           =   5
      Left            =   2400
      TabIndex        =   12
      Top             =   5760
      Width           =   2055
   End
   Begin PhotoDemon.pdComboBox pdComboBox1 
      Height          =   360
      Left            =   4800
      TabIndex        =   6
      Top             =   6000
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   635
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
      Caption         =   "toggle enablement"
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
      Height          =   2055
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
      Width           =   1650
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
   Begin PhotoDemon.pdTextBox pdTextBox2 
      Height          =   315
      Index           =   0
      Left            =   240
      TabIndex        =   7
      Top             =   2640
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   556
   End
   Begin PhotoDemon.pdTextBox pdTextBox2 
      Height          =   315
      Index           =   1
      Left            =   240
      TabIndex        =   8
      Top             =   3000
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   556
   End
   Begin PhotoDemon.pdTextBox pdTextBox2 
      Height          =   315
      Index           =   2
      Left            =   240
      TabIndex        =   9
      Top             =   3360
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   556
      Enabled         =   0   'False
   End
   Begin PhotoDemon.pdComboBox pdComboBox2 
      Height          =   360
      Index           =   0
      Left            =   4800
      TabIndex        =   10
      Top             =   6480
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   635
   End
   Begin PhotoDemon.pdComboBox pdComboBox2 
      Height          =   360
      Index           =   1
      Left            =   4800
      TabIndex        =   11
      Top             =   6960
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   635
      Enabled         =   0   'False
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
            pdTextBox1.Enabled = Not pdTextBox1.Enabled
            pdComboBox2(1).Enabled = Not pdComboBox2(1).Enabled
            
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
            
        '4 - add new drop-down entries
        Case 4
            pdComboBox1.AddItem pdComboBox1.ListCount
        
        '5 - delete drop-down entries
        Case 5
            pdComboBox1.RemoveItem pdComboBox1.ListCount - 1
        
        
    End Select
    
End Sub

Private Sub cmdTextBoxTesting_GotFocus(Index As Integer)
    Debug.Print "Focus set to TabStop item " & cmdTextBoxTesting(Index).TabIndex
End Sub

Private Sub Form_Activate()
    pdComboBox1.AddItem "0", 0
    pdComboBox1.AddItem "1", 1
    pdComboBox1.AddItem "2", 2
    pdComboBox1.ListIndex = 1
End Sub

Private Sub Form_Load()

    pdComboBox2(0).AddItem "1 - added at Form_Load", 0
    pdComboBox2(0).AddItem "2 - added at Form_Load", 1
    pdComboBox2(0).AddItem "3 - added at Form_Load", 2
    pdComboBox2(0).AddItem "4 - added at Form_Load, also ListIndex", 3
    pdComboBox2(0).AddItem ChrW$(&H6B22) & ChrW$(&H8FCE) & ChrW$(&H6B22) & "abc", 4
    pdTextBox2(1).Text = ChrW$(&H6B22) & ChrW$(&H8FCE) & ChrW$(&H6B22) & "abc"
    pdComboBox2(0).ListIndex = 3
    
    makeFormPretty Me
    
End Sub

Private Sub pdComboBox1_Click()
    Debug.Print "pdComboBox1 clicked: " & pdComboBox1.ListIndex
End Sub

Private Sub pdComboBox2_Click(Index As Integer)
    Debug.Print "pdComboBox2(" & CStr(Index) & ") clicked: " & pdComboBox2(Index).ListIndex
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
