VERSION 5.00
Begin VB.Form FormPreferences 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " "
   ClientHeight    =   6150
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8085
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
   ScaleHeight     =   410
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   539
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picCanvasImg 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   14040
      Picture         =   "VBP_FormPreferences.frx":0000
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   2
      Top             =   1800
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   6600
      TabIndex        =   1
      Top             =   5400
      Width           =   1245
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   5280
      TabIndex        =   0
      Top             =   5400
      Width           =   1245
   End
   Begin PhotoDemon.jcbutton cmdCategory 
      Height          =   1140
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   2011
      ButtonStyle     =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Interface"
      ForeColor       =   4210752
      ForeColorHover  =   4194304
      Mode            =   2
      Value           =   -1  'True
      HandPointer     =   -1  'True
      PictureNormal   =   "VBP_FormPreferences.frx":008A
      PictureAlign    =   6
      PictureEffectOnDown=   0
      DisabledPictureMode=   1
      CaptionEffects  =   0
      TooltipType     =   1
      TooltipTitle    =   "Interface Preferences"
   End
   Begin PhotoDemon.jcbutton cmdCategory 
      Height          =   1140
      Index           =   2
      Left            =   240
      TabIndex        =   4
      Top             =   2640
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   2011
      ButtonStyle     =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Updates"
      ForeColor       =   4210752
      ForeColorHover  =   4194304
      Mode            =   2
      HandPointer     =   -1  'True
      PictureNormal   =   "VBP_FormPreferences.frx":10DC
      PictureAlign    =   6
      PictureEffectOnDown=   0
      DisabledPictureMode=   1
      CaptionEffects  =   0
      TooltipType     =   1
      TooltipTitle    =   "Update Preferences"
   End
   Begin PhotoDemon.jcbutton cmdCategory 
      Height          =   1140
      Index           =   1
      Left            =   240
      TabIndex        =   5
      Top             =   1440
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   2011
      ButtonStyle     =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Tools"
      ForeColor       =   4210752
      ForeColorHover  =   4194304
      Mode            =   2
      HandPointer     =   -1  'True
      PictureNormal   =   "VBP_FormPreferences.frx":212E
      PictureAlign    =   6
      PictureEffectOnDown=   0
      DisabledPictureMode=   1
      CaptionEffects  =   0
      TooltipType     =   1
      TooltipTitle    =   "Tool Preferences"
   End
   Begin PhotoDemon.jcbutton cmdCategory 
      Height          =   1140
      Index           =   3
      Left            =   240
      TabIndex        =   6
      Top             =   3840
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   2011
      ButtonStyle     =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Advanced"
      ForeColor       =   4210752
      ForeColorHover  =   4194304
      Mode            =   2
      HandPointer     =   -1  'True
      PictureNormal   =   "VBP_FormPreferences.frx":3180
      PictureAlign    =   6
      PictureEffectOnDown=   0
      DisabledPictureMode=   1
      CaptionEffects  =   0
      TooltipType     =   1
      TooltipTitle    =   "Advanced Settings"
   End
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4695
      Index           =   0
      Left            =   2280
      MousePointer    =   1  'Arrow
      ScaleHeight     =   313
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   369
      TabIndex        =   7
      Top             =   240
      Width           =   5535
      Begin VB.PictureBox picAlphaTwo 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   4800
         MouseIcon       =   "VBP_FormPreferences.frx":41D2
         MousePointer    =   99  'Custom
         ScaleHeight     =   20
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   37
         TabIndex        =   33
         TabStop         =   0   'False
         ToolTipText     =   "Click to change the second checkerboard background color for alpha channels"
         Top             =   1755
         Width           =   585
      End
      Begin VB.PictureBox picAlphaOne 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   4080
         MouseIcon       =   "VBP_FormPreferences.frx":4324
         MousePointer    =   99  'Custom
         ScaleHeight     =   20
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   37
         TabIndex        =   31
         TabStop         =   0   'False
         ToolTipText     =   "Click to change the second checkerboard background color for alpha channels"
         Top             =   1755
         Width           =   585
      End
      Begin VB.ComboBox cmbAlphaCheck 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   360
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   1755
         Width           =   3615
      End
      Begin VB.CheckBox chkConfirmUnsaved 
         Appearance      =   0  'Flat
         Caption         =   "warn me if the image has unsaved changes"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   360
         TabIndex        =   22
         ToolTipText     =   "Check this if you want to be warned when you try to close an image with unsaved changes"
         Top             =   3510
         Width           =   5055
      End
      Begin VB.ComboBox cmbLargeImages 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   360
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   2640
         Width           =   5055
      End
      Begin VB.ComboBox cmbCanvas 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   360
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   960
         Width           =   4335
      End
      Begin VB.PictureBox picCanvasColor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   4800
         MouseIcon       =   "VBP_FormPreferences.frx":4476
         MousePointer    =   99  'Custom
         ScaleHeight     =   20
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   37
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Click to change the image window background color"
         Top             =   960
         Width           =   585
      End
      Begin VB.Label lblAlphaCheck 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "alpha (transparency) background:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   240
         Left            =   240
         TabIndex        =   32
         Top             =   1440
         Width           =   2910
      End
      Begin VB.Label lblClosingUnsavedImages 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "when image files are closed:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   240
         Left            =   240
         TabIndex        =   23
         Top             =   3150
         Width           =   2475
      End
      Begin VB.Label lblImgOpen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "when image files are opened: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   240
         Left            =   240
         TabIndex        =   13
         Top             =   2280
         Width           =   2625
      End
      Begin VB.Label lblCanvasFX 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "image window background:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   240
         Left            =   240
         TabIndex        =   12
         Top             =   645
         Width           =   2370
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "interface and theme preferences"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   0
         Width           =   3390
      End
   End
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4695
      Index           =   3
      Left            =   2280
      MousePointer    =   1  'Arrow
      ScaleHeight     =   313
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   369
      TabIndex        =   24
      Top             =   240
      Width           =   5535
      Begin VB.TextBox TxtTempPath 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   28
         Text            =   "automatically generated at run-time"
         ToolTipText     =   "Folder used for temporary files"
         Top             =   1560
         Width           =   4575
      End
      Begin VB.CommandButton CmdTmpPath 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   5040
         TabIndex        =   27
         ToolTipText     =   "Click to open a browse-for-folder dialog"
         Top             =   1560
         Width           =   375
      End
      Begin VB.CheckBox ChkLogMessages 
         Appearance      =   0  'Flat
         Caption         =   "log program messages to file (for debugging purposes)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   240
         TabIndex        =   26
         ToolTipText     =   $"VBP_FormPreferences.frx":45C8
         Top             =   600
         Width           =   5175
      End
      Begin VB.Label lblTempFolder 
         BackStyle       =   0  'Transparent
         Caption         =   "temporary file folder (used to hold Undo/Redo data):"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   1200
         Width           =   4575
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "advanced settings"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   3
         Left            =   120
         TabIndex        =   25
         Top             =   0
         Width           =   1875
      End
   End
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4695
      Index           =   2
      Left            =   2280
      MousePointer    =   1  'Arrow
      ScaleHeight     =   313
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   369
      TabIndex        =   17
      Top             =   240
      Width           =   5535
      Begin VB.CheckBox ChkPromptPluginDownload 
         Appearance      =   0  'Flat
         Caption         =   "if core plugins cannot be located, offer to download them"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   240
         TabIndex        =   20
         ToolTipText     =   $"VBP_FormPreferences.frx":46BA
         Top             =   1080
         Width           =   5295
      End
      Begin VB.CheckBox chkProgramUpdates 
         Appearance      =   0  'Flat
         Caption         =   "automatically check for software updates every 10 days"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   495
         Left            =   240
         TabIndex        =   19
         ToolTipText     =   "If this is disabled, you can visit tannerhelland.com/photodemon to manually download the latest version of PhotoDemon"
         Top             =   480
         Width           =   5295
      End
      Begin VB.Label lblExplanation 
         BackStyle       =   0  'Transparent
         Caption         =   "(disclaimer populated at run-time)"
         ForeColor       =   &H00808080&
         Height          =   3015
         Left            =   240
         TabIndex        =   21
         Top             =   1560
         Width           =   5295
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "update preferences"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   2
         Left            =   120
         TabIndex        =   18
         Top             =   0
         Width           =   2010
      End
   End
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4695
      Index           =   1
      Left            =   2280
      MousePointer    =   1  'Arrow
      ScaleHeight     =   313
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   369
      TabIndex        =   14
      Top             =   240
      Width           =   5535
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "tool preferences"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   16
         Top             =   0
         Width           =   1695
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "There are not currently any tool settings."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   240
         Left            =   240
         TabIndex        =   15
         Top             =   645
         Width           =   3510
      End
   End
   Begin VB.Line lneVertical 
      BorderColor     =   &H8000000D&
      X1              =   136
      X2              =   136
      Y1              =   8
      Y2              =   344
   End
End
Attribute VB_Name = "FormPreferences"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Program Preferences Handler
'Copyright ©2000-2012 by Tanner Helland
'Created: 8/November/02
'Last updated: 22/October/12
'Last update: revamped entire interface; settings are now sorted by category.
'
'Module for interfacing with the user's desired program preferences.  Handles
' reading from and copying to the program's ".INI" file.
'
'Note that this form interacts heavily with the INIProcessor module.
'
'***************************************************************************

Option Explicit

'Used to see if the user physically clicked a combo box, or if VB selected it on its own
Dim userInitiatedColorSelection As Boolean
Dim userInitiatedAlphaSelection As Boolean

'Alpha channel checkerboard selection
Private Sub cmbAlphaCheck_Click()

    'Only respond to user-generated events
    If userInitiatedAlphaSelection = False Then Exit Sub

    'Redraw the sample picture boxes based on the value the user has selected
    AlphaCheckMode = cmbAlphaCheck.ListIndex
    Select Case cmbAlphaCheck.ListIndex
    
        'Case 0 - Highlights
        Case 0
            AlphaCheckOne = RGB(255, 255, 255)
            AlphaCheckTwo = RGB(204, 204, 204)
        
        'Case 1 - Midtones
        Case 1
            AlphaCheckOne = RGB(153, 153, 153)
            AlphaCheckTwo = RGB(102, 102, 102)
        
        'Case 2 - Shadows
        Case 2
            AlphaCheckOne = RGB(51, 51, 51)
            AlphaCheckTwo = RGB(0, 0, 0)
        
        'Case 3 - Custom
        Case 3
            AlphaCheckOne = RGB(255, 204, 246)
            AlphaCheckTwo = RGB(255, 255, 255)
        
    End Select

    'Change the picture boxes to match the current selection
    picAlphaOne.BackColor = AlphaCheckOne
    picAlphaTwo.BackColor = AlphaCheckTwo

End Sub

'Canvas background selection
Private Sub cmbCanvas_Click()
    
    'Only respond to user-generated events
    If userInitiatedColorSelection = False Then Exit Sub
    
    'Redraw the sample picture box value based on the value the user has selected
    Select Case cmbCanvas.ListIndex
        Case 0
            CanvasBackground = -1
        Case 1
            CanvasBackground = vb3DLight
        Case 2
            CanvasBackground = vb3DShadow
        Case 3
            'Prompt with a color selection box
            Dim retColor As Long
    
            Dim CD1 As cCommonDialog
            Set CD1 = New cCommonDialog
    
            retColor = picCanvasColor.BackColor
    
            CD1.VBChooseColor retColor, True, True, False, Me.HWnd
    
            'If a color was selected, change the picture box and associated combo box to match
            If retColor > 0 Then CanvasBackground = retColor Else CanvasBackground = picCanvasColor.BackColor
            
    End Select
    
    DrawSampleCanvasBackground
    
End Sub

'CANCEL button
Private Sub CmdCancel_Click()
    Unload Me
End Sub

'When the category is changed, only display the controls in that category
Private Sub cmdCategory_Click(Index As Integer)
    
    Static catID As Long
    For catID = 0 To cmdCategory.Count - 1
        If catID = Index Then picContainer(catID).Visible = True Else picContainer(catID).Visible = False
    Next catID
    
End Sub

'OK button
Private Sub CmdOK_Click()
    
    'Store whether the user wants to be prompted when closing unsaved images
    If chkConfirmUnsaved.Value = vbChecked Then
        ConfirmClosingUnsaved = True
        WriteToIni "General Preferences", "ConfirmClosingUnsaved", 1
        FormMain.cmdClose.ToolTip = "Close the current image." & vbCrLf & vbCrLf & "If the current image has not been saved, you will" & vbCrLf & " receive a prompt to save it before it closes."
    Else
        ConfirmClosingUnsaved = False
        WriteToIni "General Preferences", "ConfirmClosingUnsaved", 0
        FormMain.cmdClose.ToolTip = "Close the current image." & vbCrLf & vbCrLf & "Because you have turned off save prompts (via Edit -> Preferences)," & vbCrLf & " you WILL NOT receive a prompt to save this image before it closes."
    End If
    
    'Store whether PhotoDemon is allowed to check for updates
    If chkProgramUpdates.Value = vbChecked Then
        WriteToIni "General Preferences", "CheckForUpdates", 1
    Else
        WriteToIni "General Preferences", "CheckForUpdates", 0
    End If
    
    'Store whether PhotoDemon is allowed to offer the automatic download of missing core plugins
    If ChkPromptPluginDownload.Value = vbChecked Then
        WriteToIni "General Preferences", "PromptForPluginDownload", 1
    Else
        WriteToIni "General Preferences", "PromptForPluginDownload", 0
    End If
    
    'Store whether we'll log system messages or not
    If ChkLogMessages.Value = vbChecked Then
        LogProgramMessages = True
        WriteToIni "General Preferences", "LogProgramMessages", 1
    Else
        LogProgramMessages = False
        WriteToIni "General Preferences", "LogProgramMessages", 0
    End If
    
    'Store the canvas background preference
    Select Case cmbCanvas.ListIndex
        
        'Checkerboard pattern
        Case 0
            CanvasBackground = -1
            WriteToIni "General Preferences", "CanvasBackground", "-1"
            
        'Color only
        Case Else
            CanvasBackground = picCanvasColor.BackColor
            WriteToIni "General Preferences", "CanvasBackground", CStr(CLng(picCanvasColor.BackColor))
            
    End Select
    
    'Store the alpha checkerboard preference
    WriteToIni "General Preferences", "AlphaCheckMode", CStr(cmbAlphaCheck.ListIndex)
    WriteToIni "General Preferences", "AlphaCheckOne", CStr(CLng(picAlphaOne.BackColor))
    WriteToIni "General Preferences", "AlphaCheckTwo", CStr(CLng(picAlphaTwo.BackColor))
    
    'Now run a loop to draw the checkerboard effect on every window
    Dim tForm As Form
    Message "Saving preferences..."
    For Each tForm In VB.Forms
        If tForm.Name = "FormImage" Then PrepareViewport tForm
    Next
    Message "Finished."
    
    'Remember whether or not to autozoom large images
    AutosizeLargeImages = cmbLargeImages.ListIndex
    WriteToIni "General Preferences", "AutosizeLargeImages", CStr(AutosizeLargeImages)
    
    'Verify the temporary path
    If LCase(TxtTempPath.Text) <> LCase(TempPath) Then
        TempPath = TxtTempPath.Text
        WriteToIni "Paths", "TempPath", TxtTempPath.Text
    End If
    
    Unload Me
    
End Sub

'When the "..." button is clicked, prompt the user with a "browse for folder" dialog
Private Sub CmdTmpPath_Click()
    Dim tString As String
    tString = BrowseForFolder(Me.HWnd)
    If tString <> "" Then TxtTempPath.Text = FixPath(tString)
End Sub

'When the form is loaded, populate the various checkboxes and textboxes with the values from the INI file
Private Sub Form_Load()
    
    Me.Caption = PROGRAMNAME & " Preferences"
    
    'Load all relevant values from the INI file, and populate their corresponding controls with the user's current settings
    
    'Start with the canvas background (which also requires populating the canvas background combo box)
    userInitiatedColorSelection = False
    cmbCanvas.AddItem "Checkerboard pattern", 0
    cmbCanvas.AddItem "System theme: light", 1
    cmbCanvas.AddItem "System theme: dark", 2
    cmbCanvas.AddItem "Custom color (click box to customize)", 3
        
    'Select the proper combo box value based on the CanvasBackground variable
    If CanvasBackground = -1 Then
        'Checkerboard pattern
        cmbCanvas.ListIndex = 0
    ElseIf CanvasBackground = vb3DLight Then
        'System theme: light
        cmbCanvas.ListIndex = 1
    ElseIf CanvasBackground = vb3DShadow Then
        'System theme: dark
        cmbCanvas.ListIndex = 2
    Else
        'Custom color
        cmbCanvas.ListIndex = 3
    End If
    
    'Draw the current canvas background to the sample picture box
    DrawSampleCanvasBackground
    userInitiatedColorSelection = True
    
    'Next, get the values for alpha-channel checkerboard rendering
    userInitiatedAlphaSelection = False
    
    cmbAlphaCheck.AddItem "Highlight checks", 0
    cmbAlphaCheck.AddItem "Midtone checks", 1
    cmbAlphaCheck.AddItem "Shadow checks", 2
    cmbAlphaCheck.AddItem "Custom (click boxes to customize)", 3
    
    cmbAlphaCheck.ListIndex = AlphaCheckMode
    
    picAlphaOne.BackColor = AlphaCheckOne
    picAlphaTwo.BackColor = AlphaCheckTwo
    
    userInitiatedAlphaSelection = True
    
    'Assign the check box for logging program messages
    If LogProgramMessages = True Then ChkLogMessages.Value = vbChecked Else ChkLogMessages.Value = vbUnchecked
    
    'Assign the check box for prompting about unsaved images
    If ConfirmClosingUnsaved = True Then chkConfirmUnsaved.Value = vbChecked Else chkConfirmUnsaved.Value = vbUnchecked
    
    'Display the current temporary file path
    TxtTempPath.Text = TempPath
    
    'We have to pull the "offer to download plugins" value from the INI file, since we don't track
    ' it internally (it's only accessed when PhotoDemon is first loaded)
    Dim tmpString As String
    tmpString = GetFromIni("General Preferences", "PromptForPluginDownload")
    If Val(tmpString) = 1 Then ChkPromptPluginDownload.Value = vbChecked Else ChkPromptPluginDownload.Value = vbUnchecked
    
    'Same for checking for software updates
    tmpString = GetFromIni("General Preferences", "CheckForUpdates")
    If Val(tmpString) = 1 Then chkProgramUpdates.Value = vbChecked Else chkProgramUpdates.Value = vbUnchecked
    
    'Populate the "what to do when loading large images" combo box
    cmbLargeImages.AddItem "Automatically zoom out so the images fit on-screen", 0
    cmbLargeImages.AddItem "Load images at 100% zoom regardless of size", 1
    
    tmpString = GetFromIni("General Preferences", "AutosizeLargeImages")
    cmbLargeImages.ListIndex = Val(tmpString)
    
    'Populate the multi-line tooltips for the category command buttons
    'Interface
    cmdCategory(0).ToolTip = "Interface preferences include default setting for canvas backgrounds," & vbCrLf & "transparency checkerboards, and the program's visual theme."
    'Tools
    cmdCategory(1).ToolTip = "Tool preferences currently includes customizable options for the selection tool." & vbCrLf & "In the future, PhotoDemon will gain paint tools, and those settings will appear" & vbCrLf & "here as well."
    'Updates
    cmdCategory(2).ToolTip = "Update preferences control how frequently PhotoDemon checks for" & vbCrLf & "updated versions, and how it handles the download of missing plugins."
    'Advanced
    cmdCategory(3).ToolTip = "Advanced preferences can be safely ignored by regular users." & vbCrLf & "Testers and developers may, however, find these settings useful."
    
    'Populate the network access disclaimer in the "Update" panel
    lblExplanation.Caption = PROGRAMNAME & " provides two non-essential features that require Internet access: checking for software updates, and offering to download core plugins (FreeImage, EZTwain, and ZLib) if they aren't present in the \Data\Plugins subdirectory." _
    & vbCrLf & vbCrLf & "The developers of " & PROGRAMNAME & " take privacy very seriously, so no information - statistical or otherwise - is uploaded by these features. Checking for software updates involves downloading a single ""updates.txt"" file containing the latest PhotoDemon version number. Similarly, downloading missing plugins simply involves downloading one or more of the compressed plugin files from the " & PROGRAMNAME & " server." _
    & vbCrLf & vbCrLf & "If you choose to disable these features, note that you can always visit tannerhelland.com/photodemon to manually download the most recent version of the program."
    
    'Finally, hide the inactive category panels
    picContainer(1).Visible = False
    picContainer(2).Visible = False
    picContainer(3).Visible = False
    
    'Assign the system hand cursor to all relevant objects
    makeFormPretty Me
    
    'Note: at present, this doesn't seem to be working, and I'm not sure why.  It has something to do with
    ' picture boxes contained within other picture boxes.  Because of this, I've manually set the mouse icon
    ' to an old-school hand cursor (which is all VB will accept).
    'setHandCursor picCanvasColor
    'setHandCursor picAlphaOne
    'setHandCursor picAlphaTwo
    
    'For some reason, the container picture boxes automatically acquire the pointer of children objects.
    ' Manually force those cursors to arrows to prevent this.
    setArrowCursorToObject picContainer(0)
    setArrowCursorToObject picContainer(1)
    setArrowCursorToObject picContainer(2)
    setArrowCursorToObject picContainer(3)
        
End Sub

'Draw a sample of the current background to the PicCanvasColor picture box
Private Sub DrawSampleCanvasBackground()

    '-1 indicates the user wants a checkboard background pattern
    If CanvasBackground = -1 Then

        Dim stepIntervalX As Long, stepIntervalY As Long
        stepIntervalX = Me.picCanvasImg.ScaleWidth
        stepIntervalY = Me.picCanvasImg.ScaleHeight

        For x = 0 To picCanvasColor.ScaleWidth Step stepIntervalX
        For y = 0 To picCanvasColor.ScaleHeight Step stepIntervalY
            BitBlt Me.picCanvasColor.hDC, x, y, stepIntervalX, stepIntervalY, Me.picCanvasImg.hDC, 0, 0, vbSrcCopy
        Next y
        Next x
        picCanvasColor.Picture = picCanvasColor.Image
        
        Me.picCanvasColor.Enabled = False
        
    'Any other value is treated as an RGB long
    Else
    
        Me.picCanvasColor.Picture = LoadPicture("")
        Me.picCanvasColor.BackColor = CanvasBackground
    
        Me.picCanvasColor.Enabled = True
    
    End If
    
End Sub

'Allow the user to change the first checkerboard color for alpha channels
Private Sub picAlphaOne_Click()
    
    Dim retColor As Long
    
    Dim CD1 As cCommonDialog
    Set CD1 = New cCommonDialog
    
    retColor = picAlphaOne.BackColor
    
    'Display a Windows color selection box
    CD1.VBChooseColor retColor, True, True, False, Me.HWnd
    
    'If a color was selected, change the picture box and associated combo box to match
    If retColor > 0 Then
    
        AlphaCheckOne = retColor
        picAlphaOne.BackColor = retColor
        
        userInitiatedAlphaSelection = False
        cmbAlphaCheck.ListIndex = 3   '3 corresponds to "custom colors"
        userInitiatedAlphaSelection = True
                
    End If
    
End Sub

'Allow the user to change the second checkerboard color for alpha channels
Private Sub picAlphaTwo_Click()
    
    Dim retColor As Long
    
    Dim CD1 As cCommonDialog
    Set CD1 = New cCommonDialog
    
    retColor = picAlphaTwo.BackColor
    
    'Display a Windows color selection box
    CD1.VBChooseColor retColor, True, True, False, Me.HWnd
    
    'If a color was selected, change the picture box and associated combo box to match
    If retColor > 0 Then
    
        AlphaCheckTwo = retColor
        picAlphaTwo.BackColor = retColor
        
        userInitiatedAlphaSelection = False
        cmbAlphaCheck.ListIndex = 3   '3 corresponds to "custom colors"
        userInitiatedAlphaSelection = True
                
    End If
    
End Sub

'Clicking the sample color box allows the user to pick a new color
Private Sub picCanvasColor_Click()
    
    Dim retColor As Long
    
    Dim CD1 As cCommonDialog
    Set CD1 = New cCommonDialog
    
    retColor = picCanvasColor.BackColor
    
    'Display a Windows color selection box
    CD1.VBChooseColor retColor, True, True, False, Me.HWnd
    
    'If a color was selected, change the picture box and associated combo box to match
    If retColor > 0 Then
    
        CanvasBackground = retColor
        
        userInitiatedColorSelection = False
        If CanvasBackground = vb3DLight Then
            'System theme: light
            cmbCanvas.ListIndex = 1
        ElseIf CanvasBackground = vb3DShadow Then
            'System theme: dark
            cmbCanvas.ListIndex = 2
        Else
            'Custom color
            cmbCanvas.ListIndex = 3
        End If
        userInitiatedColorSelection = True
        
        DrawSampleCanvasBackground
        
    End If
    
End Sub
