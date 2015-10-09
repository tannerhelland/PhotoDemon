VERSION 5.00
Begin VB.Form FormThemeEditor 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Theme editor"
   ClientHeight    =   8985
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   13260
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
   ScaleHeight     =   599
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   884
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.pdLabel lblHScroll 
      Height          =   255
      Index           =   0
      Left            =   10320
      Top             =   1320
      Width           =   1335
      _extentx        =   2355
      _extenty        =   503
      alignment       =   2
      caption         =   ""
   End
   Begin PhotoDemon.pdScrollBar hScroll1 
      Height          =   255
      Index           =   0
      Left            =   10320
      TabIndex        =   14
      Top             =   600
      Width           =   2775
      _extentx        =   4895
      _extenty        =   450
      backcolor       =   0
      orientationhorizontal=   -1  'True
   End
   Begin PhotoDemon.pdButton cmdTest 
      Height          =   615
      Index           =   0
      Left            =   240
      TabIndex        =   8
      Top             =   4320
      Width           =   2055
      _extentx        =   3625
      _extenty        =   1085
      caption         =   "toggle visibility"
   End
   Begin PhotoDemon.pdLabel lblExplanation 
      Height          =   855
      Left            =   120
      Top             =   7560
      Width           =   10335
      _extentx        =   18230
      _extenty        =   1508
      caption         =   $"Tools_ThemeEditor.frx":0000
      fontsize        =   12
      layout          =   3
   End
   Begin PhotoDemon.pdComboBox pdComboBox1 
      Height          =   375
      Left            =   4800
      TabIndex        =   2
      Top             =   6000
      Width           =   5655
      _extentx        =   9975
      _extenty        =   635
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
      Width           =   1650
      _extentx        =   2910
      _extenty        =   503
      caption         =   "testing options:"
      fontsize        =   12
      layout          =   2
   End
   Begin PhotoDemon.pdLabel pdLabelTitle 
      Height          =   285
      Index           =   2
      Left            =   4800
      Top             =   3840
      Width           =   4455
      _extentx        =   7858
      _extenty        =   503
      caption         =   "Other Unicode controls for testing output:"
      fontsize        =   12
      layout          =   2
   End
   Begin PhotoDemon.pdTextBox pdTextBox2 
      Height          =   315
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   2640
      Width           =   9975
      _extentx        =   17595
      _extenty        =   556
   End
   Begin PhotoDemon.pdTextBox pdTextBox2 
      Height          =   315
      Index           =   1
      Left            =   240
      TabIndex        =   4
      Top             =   3000
      Width           =   9975
      _extentx        =   17595
      _extenty        =   556
   End
   Begin PhotoDemon.pdTextBox pdTextBox2 
      Height          =   315
      Index           =   2
      Left            =   240
      TabIndex        =   5
      Top             =   3360
      Width           =   9975
      _extentx        =   17595
      _extenty        =   556
      enabled         =   0   'False
   End
   Begin PhotoDemon.pdComboBox pdComboBox2 
      Height          =   375
      Index           =   0
      Left            =   4800
      TabIndex        =   6
      Top             =   6480
      Width           =   5655
      _extentx        =   9975
      _extenty        =   635
   End
   Begin PhotoDemon.pdComboBox pdComboBox2 
      Height          =   375
      Index           =   1
      Left            =   4800
      TabIndex        =   7
      Top             =   6960
      Width           =   5655
      _extentx        =   9975
      _extenty        =   635
      enabled         =   0   'False
   End
   Begin PhotoDemon.pdButton cmdTest 
      Height          =   615
      Index           =   1
      Left            =   2400
      TabIndex        =   9
      Top             =   4320
      Width           =   2055
      _extentx        =   3625
      _extenty        =   1085
      caption         =   "toggle enablement"
   End
   Begin PhotoDemon.pdButton cmdTest 
      Height          =   615
      Index           =   2
      Left            =   240
      TabIndex        =   10
      Top             =   5040
      Width           =   2055
      _extentx        =   3625
      _extenty        =   1085
      caption         =   "append random chars"
   End
   Begin PhotoDemon.pdButton cmdTest 
      Height          =   615
      Index           =   3
      Left            =   2400
      TabIndex        =   11
      Top             =   5040
      Width           =   2055
      _extentx        =   3625
      _extenty        =   1085
      caption         =   "toggle multiline"
   End
   Begin PhotoDemon.pdButton cmdTest 
      Height          =   615
      Index           =   4
      Left            =   240
      TabIndex        =   12
      Top             =   5760
      Width           =   2055
      _extentx        =   3625
      _extenty        =   1085
      caption         =   "add drop-down entries"
   End
   Begin PhotoDemon.pdButton cmdTest 
      Height          =   615
      Index           =   5
      Left            =   2400
      TabIndex        =   13
      Top             =   5760
      Width           =   2055
      _extentx        =   3625
      _extenty        =   1085
      caption         =   "delete drop-down entries"
   End
   Begin PhotoDemon.pdLabel pdLabelTitle 
      Height          =   285
      Index           =   3
      Left            =   10320
      Top             =   120
      Width           =   1875
      _extentx        =   5345
      _extenty        =   503
      caption         =   "Unicode text box for testing:"
      fontsize        =   12
      layout          =   2
   End
   Begin PhotoDemon.pdScrollBar hScroll1 
      Height          =   255
      Index           =   1
      Left            =   10320
      TabIndex        =   15
      Top             =   960
      Width           =   2775
      _extentx        =   4895
      _extenty        =   450
      backcolor       =   0
      max             =   100
      orientationhorizontal=   -1  'True
   End
   Begin PhotoDemon.pdScrollBar vScroll1 
      Height          =   2775
      Index           =   0
      Left            =   11760
      TabIndex        =   16
      Top             =   1320
      Width           =   255
      _extentx        =   4895
      _extenty        =   450
      backcolor       =   0
   End
   Begin PhotoDemon.pdScrollBar vScroll1 
      Height          =   2775
      Index           =   1
      Left            =   12120
      TabIndex        =   17
      Top             =   1320
      Width           =   255
      _extentx        =   4895
      _extenty        =   450
      backcolor       =   0
      max             =   100
   End
   Begin PhotoDemon.pdLabel lblHScroll 
      Height          =   255
      Index           =   1
      Left            =   10320
      Top             =   1680
      Width           =   1335
      _extentx        =   2355
      _extenty        =   503
      alignment       =   2
      caption         =   ""
   End
   Begin PhotoDemon.pdLabel lblVScroll 
      Height          =   255
      Index           =   0
      Left            =   10320
      Top             =   2520
      Width           =   1335
      _extentx        =   2355
      _extenty        =   503
      alignment       =   2
      caption         =   ""
   End
   Begin PhotoDemon.pdLabel lblVScroll 
      Height          =   255
      Index           =   1
      Left            =   10320
      Top             =   2880
      Width           =   1335
      _extentx        =   2355
      _extenty        =   503
      alignment       =   2
      caption         =   ""
   End
End
Attribute VB_Name = "FormThemeEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdTest_GotFocusAPI(Index As Integer)
    Debug.Print "Focus set to TabStop item " & cmdTest(Index).TabIndex
End Sub

Private Sub cmdTest_Click(Index As Integer)
    
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
            pdComboBox1.AddItem pdComboBox1.ListCount, , True
                    
        '5 - delete drop-down entries
        Case 5
            pdComboBox1.RemoveItem pdComboBox1.ListCount - 1
        
        
    End Select
    
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
    
    MakeFormPretty Me
    
End Sub

Private Sub hScroll1_Scroll(Index As Integer, ByVal eventIsCritical As Boolean)
    lblHScroll(Index).Caption = CStr(hScroll1(Index).Value)
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

Private Sub vScroll1_Scroll(Index As Integer, ByVal eventIsCritical As Boolean)
    lblVScroll(Index).Caption = CStr(vScroll1(Index).Value)
End Sub
