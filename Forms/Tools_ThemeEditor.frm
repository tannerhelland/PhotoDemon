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
   Begin PhotoDemon.buttonStrip btsToggleTest 
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   873
   End
   Begin PhotoDemon.pdLabel lblExplanation 
      Height          =   855
      Left            =   120
      Top             =   8040
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   1508
      Caption         =   $"Tools_ThemeEditor.frx":0000
      FontSize        =   12
      Layout          =   3
   End
   Begin PhotoDemon.pdLabel pdLabelTitle 
      Height          =   285
      Index           =   0
      Left            =   120
      Top             =   120
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   503
      Caption         =   "Toggle theme (for testing purposes only!):"
      FontSize        =   12
   End
End
Attribute VB_Name = "FormThemeEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btsToggleTest_Click(ByVal buttonIndex As Long)

    If (buttonIndex = 0) Then
        g_Themer.LoadThemeFile "Default_Light.xml"
    Else
        g_Themer.LoadThemeFile "Default_Dark.xml"
    End If
    
    Interface.ApplyThemeAndTranslations Me
    
    'Eventually, form backcolor will be moved into the theming code, but for now, apply it manually
    Me.BackColor = Colors.GetRGBLongFromHex(g_Themer.LookUpColor("Default", "Background"))
    
End Sub

Private Sub Form_Load()
    btsToggleTest.AddItem "Light theme", 0
    btsToggleTest.AddItem "Dark theme", 1
    Interface.ApplyThemeAndTranslations Me
End Sub
