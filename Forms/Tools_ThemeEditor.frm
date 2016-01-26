VERSION 5.00
Begin VB.Form FormThemeEditor 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Theme editor"
   ClientHeight    =   9105
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
   ScaleHeight     =   607
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   884
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.buttonStripVertical btsvTest 
      Height          =   2535
      Left            =   120
      TabIndex        =   2
      Top             =   3000
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   4471
      Caption         =   "I'm a vertical button strip"
   End
   Begin PhotoDemon.buttonStrip btsTest 
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   1920
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   1720
      Caption         =   "I'm a horizontal button strip"
   End
   Begin PhotoDemon.buttonStrip btsToggleTest 
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12975
      _ExtentX        =   15266
      _ExtentY        =   1720
      Caption         =   "toggle theme (please don't exit without clicking LIGHT THEME; otherwise PD may look funky!):"
   End
   Begin PhotoDemon.pdLabel lblExplanation 
      Height          =   855
      Index           =   0
      Left            =   120
      Top             =   8040
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   1508
      Alignment       =   2
      Caption         =   $"Tools_ThemeEditor.frx":0000
      FontSize        =   12
      Layout          =   1
   End
   Begin PhotoDemon.pdLabel lblExplanation 
      Height          =   330
      Index           =   1
      Left            =   120
      Top             =   1200
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   582
      Alignment       =   2
      Caption         =   "(Note: if you edit a theme file externally, you can toggle the button above to force PD to reload the updated file.)"
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
    
    Dim i As Long
    
    For i = 0 To 4
        btsTest.AddItem "Button " & CStr(i + 1)
        btsvTest.AddItem "Button " & CStr(i + 1)
    Next i
    
    Interface.ApplyThemeAndTranslations Me
    
End Sub
