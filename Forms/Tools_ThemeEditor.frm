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
   Begin PhotoDemon.pdPenSelector pdpsTest 
      Height          =   1095
      Left            =   3120
      TabIndex        =   6
      Top             =   5160
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   1931
      Caption         =   "Pen selector test"
   End
   Begin PhotoDemon.pdGradientSelector pdgsTest 
      Height          =   1095
      Left            =   3120
      TabIndex        =   5
      Top             =   3960
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   1931
      Caption         =   "Gradient selector test"
   End
   Begin PhotoDemon.pdButtonStrip pdbsEnableTest 
      Height          =   615
      Left            =   9240
      TabIndex        =   4
      Top             =   7320
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   1085
   End
   Begin PhotoDemon.pdBrushSelector pdbsTest 
      Height          =   1095
      Left            =   3120
      TabIndex        =   3
      Top             =   2760
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   1931
      Caption         =   "Brush selector test"
   End
   Begin PhotoDemon.pdHyperlink pdhlTest 
      Height          =   375
      Index           =   0
      Left            =   120
      Top             =   7440
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      Caption         =   "I'm a basic hyperlink"
   End
   Begin PhotoDemon.pdButtonStripVertical btsvTest 
      Height          =   3495
      Left            =   120
      TabIndex        =   2
      Top             =   2760
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   4471
      Caption         =   "I'm a vertical button strip"
   End
   Begin PhotoDemon.pdButtonStrip btsTest 
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   1680
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   1720
      Caption         =   "I'm a horizontal button strip"
   End
   Begin PhotoDemon.pdButtonStrip btsToggleTest 
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   1720
      Caption         =   "toggle theme (please click LIGHT THEME before exiting):"
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
      Caption         =   "(Note: if you edit a theme file externally, you can toggle the button(s) above to force PD to refresh its theme cache.)"
   End
   Begin PhotoDemon.pdHyperlink pdhlTest 
      Height          =   375
      Index           =   1
      Left            =   2400
      Top             =   7440
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   661
      Caption         =   "I'm a hyperlink with weird formatting"
      FontBold        =   -1  'True
      FontItalic      =   -1  'True
      FontSize        =   12
   End
   Begin PhotoDemon.pdButtonStrip btsColorTest 
      Height          =   975
      Left            =   6720
      TabIndex        =   7
      Top             =   120
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   1720
      Caption         =   "toggle accent color (please click BLUE before exiting):"
   End
End
Attribute VB_Name = "FormThemeEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btsColorTest_Click(ByVal buttonIndex As Long)
    LoadRelevantThemeFile
End Sub

Private Sub btsToggleTest_Click(ByVal buttonIndex As Long)
    LoadRelevantThemeFile
End Sub

'Given the current combination of light/dark theme and accent color, load a new theme file
Private Sub LoadRelevantThemeFile()
    
    'First, figure out which base theme to load
    Dim baseThemeFile As String
    
    If (btsToggleTest.ListIndex = 0) Then
        baseThemeFile = "Default_Light.xml"
    Else
        baseThemeFile = "Default_Dark.xml"
    End If
    
    'Next, figure out which accent file to load
    Dim colorAccentFile As String
    Select Case btsColorTest.ListIndex
        
        Case 0
            colorAccentFile = "Blue.xml"
        
        Case 1
            colorAccentFile = "Green.xml"
        
        Case 2
            colorAccentFile = "Purple.xml"
        
    End Select
    
    colorAccentFile = "Colors_" & colorAccentFile
    
    'Load and apply the new theme
    g_Themer.LoadThemeFile baseThemeFile, colorAccentFile
    
    Interface.ApplyThemeAndTranslations Me
    
    'Eventually, form backcolor will be moved into the theming code, but for now, apply it manually
    Me.BackColor = Colors.GetRGBLongFromHex(g_Themer.LookUpColor("Default", "Background"))
    

End Sub

Private Sub Form_Load()
    
    btsToggleTest.AddItem "Light theme", 0
    btsToggleTest.AddItem "Dark theme", 1
    
    btsColorTest.AddItem "Blue", 0
    btsColorTest.AddItem "Green", 1
    btsColorTest.AddItem "Purple", 2
    
    Dim i As Long
    
    For i = 0 To 4
        btsTest.AddItem "Button " & CStr(i + 1)
        btsvTest.AddItem "Button " & CStr(i + 1)
    Next i
    
    pdbsEnableTest.AddItem "Enable all", 0
    pdbsEnableTest.AddItem "Disable all", 1
    pdbsEnableTest.ListIndex = 0
    
    Interface.ApplyThemeAndTranslations Me
    
End Sub

Private Sub pdbsEnableTest_Click(ByVal buttonIndex As Long)
    
    Dim enableSetting As Boolean
    enableSetting = CBool(buttonIndex = 0)
    
    Dim eControl As Control
    For Each eControl In Me.Controls
        On Error Resume Next
            If eControl.hWnd <> pdbsEnableTest.hWnd Then
                eControl.Enabled = enableSetting
            End If
        On Error GoTo 0
    Next
    
End Sub
