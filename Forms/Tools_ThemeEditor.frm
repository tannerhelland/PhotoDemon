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
   Begin PhotoDemon.pdTitle pdTitleTest 
      Height          =   375
      Left            =   10560
      TabIndex        =   18
      Top             =   2760
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
      Caption         =   "title test"
   End
   Begin PhotoDemon.pdSlider pdSliderTest 
      Height          =   735
      Index           =   0
      Left            =   7680
      TabIndex        =   16
      Top             =   4440
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   1296
      Caption         =   "first slider test"
      Min             =   -5
      Max             =   5
   End
   Begin PhotoDemon.pdSpinner pdSpinnerTest 
      Height          =   375
      Left            =   7680
      TabIndex        =   14
      Top             =   3960
      Width           =   2775
      _ExtentX        =   6165
      _ExtentY        =   661
      SigDigits       =   2
      Value           =   5
   End
   Begin PhotoDemon.pdTextBox pdTextTest 
      Height          =   255
      Index           =   0
      Left            =   7680
      TabIndex        =   13
      Top             =   2760
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   450
      Text            =   "Sample text goes here"
   End
   Begin PhotoDemon.pdScrollBar pdScrollTest 
      Height          =   255
      Index           =   0
      Left            =   10560
      TabIndex        =   10
      Top             =   1680
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   450
      Min             =   1
      Max             =   5
      Value           =   3
      OrientationHorizontal=   -1  'True
   End
   Begin PhotoDemon.pdButton pdButtonTest 
      Height          =   465
      Index           =   0
      Left            =   7680
      TabIndex        =   8
      Top             =   1680
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   820
      Caption         =   "Button"
   End
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
   Begin PhotoDemon.pdButton pdButtonTest 
      Height          =   465
      Index           =   1
      Left            =   7680
      TabIndex        =   9
      Top             =   2190
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   820
      Caption         =   "Button w/ image"
   End
   Begin PhotoDemon.pdScrollBar pdScrollTest 
      Height          =   255
      Index           =   1
      Left            =   10560
      TabIndex        =   11
      Top             =   2040
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   450
      Max             =   20
      Value           =   10
      OrientationHorizontal=   -1  'True
      VisualStyle     =   1
   End
   Begin PhotoDemon.pdScrollBar pdScrollTest 
      Height          =   255
      Index           =   2
      Left            =   10560
      TabIndex        =   12
      Top             =   2400
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   450
      Max             =   1000
      Value           =   500
      OrientationHorizontal=   -1  'True
   End
   Begin PhotoDemon.pdTextBox pdTextTest 
      Height          =   735
      Index           =   1
      Left            =   7680
      TabIndex        =   15
      Top             =   3120
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   1296
      Multiline       =   -1  'True
      Text            =   "Sample text goes here"
   End
   Begin PhotoDemon.pdSlider pdSliderTest 
      Height          =   735
      Index           =   1
      Left            =   7680
      TabIndex        =   17
      Top             =   5520
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   1296
      Caption         =   "second slider test"
      Max             =   1000
      SliderTrackStyle=   4
      NotchValueCustom=   250
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
    
    pdButtonTest(1).AssignImage "TF_NEW"
    
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
