VERSION 5.00
Begin VB.Form FormBlackAndWhite 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Black/White Color Conversion"
   ClientHeight    =   4155
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6135
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
   MouseIcon       =   "VBP_FormBlackAndWhite.frx":0000
   ScaleHeight     =   277
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   409
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.HScrollBar hsThreshold 
      Height          =   255
      Left            =   3120
      Max             =   254
      Min             =   1
      MouseIcon       =   "VBP_FormBlackAndWhite.frx":0152
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   1920
      Value           =   128
      Width           =   2775
   End
   Begin VB.TextBox txtThreshold 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   4230
      TabIndex        =   1
      Text            =   "128"
      Top             =   1560
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.CheckBox chkStretch 
      Appearance      =   0  'Flat
      Caption         =   "Stretch histogram before conversion"
      ForeColor       =   &H00400000&
      Height          =   495
      Left            =   1440
      MouseIcon       =   "VBP_FormBlackAndWhite.frx":02A4
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   2640
      Value           =   1  'Checked
      Width           =   3015
   End
   Begin VB.ListBox LstConvert 
      ForeColor       =   &H00400000&
      Height          =   2010
      Left            =   240
      MouseIcon       =   "VBP_FormBlackAndWhite.frx":03F6
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   225
      Width           =   2535
   End
   Begin VB.CommandButton CmdCancel 
      BackColor       =   &H00C0C0C0&
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4800
      MouseIcon       =   "VBP_FormBlackAndWhite.frx":0548
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   3600
      Width           =   1125
   End
   Begin VB.CommandButton CmdOK 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3600
      MaskColor       =   &H00000000&
      MouseIcon       =   "VBP_FormBlackAndWhite.frx":069A
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   3600
      Width           =   1125
   End
   Begin VB.Label lblHeader 
      BackStyle       =   0  'Transparent
      Caption         =   "Description:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   3000
      TabIndex        =   7
      Top             =   240
      Width           =   2415
   End
   Begin VB.Label lblDesc 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "<No item selected>"
      ForeColor       =   &H00400000&
      Height          =   1815
      Left            =   3120
      TabIndex        =   6
      Top             =   600
      Width           =   2775
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "FormBlackAndWhite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Black-White Color Reduction Form
'Copyright ©2000-2012 by Tanner Helland
'Created: some time 2002
'Last updated: 19/June/12
'Last update: reordered the algorithms and improved descriptions
'
'The meat of this form is in the module with the same name...look there for
'real algorithm info.
'
'***************************************************************************

Option Explicit

'CANCEL button
Private Sub CmdCancel_Click()
    Unload Me
End Sub

'OK button
Private Sub CmdOK_Click()
    'Error checking for threshold runs
    If EntryValid(txtThreshold, hsThreshold.Min, hsThreshold.Max) = False Then
        AutoSelectText txtThreshold
        Exit Sub
    End If
    
    Me.Visible = False
    'Allow the user to stretch the range of the image before converting to black and white
    If chkStretch.Value = vbChecked Then Process StretchHistogram
    
    Message "Processing 1-bit conversion..."
    Select Case LstConvert.ListIndex
        Case 0
            Process BWNearestColor2
        Case 1
            Process BWImpressionist
        Case 2
            Process BWNearestColor
        Case 3
            Process Threshold, val(txtThreshold.Text)
        Case 4
            Process BWOrderedDither
        Case 5
            Process BWDiffusionDither
        Case 6
            Process BWEnhancedDither
        Case 7
            Process BWFloydSteinberg
    End Select
    Unload Me
End Sub

Private Sub Form_Load()
    LstConvert.AddItem "Component Color"
    LstConvert.AddItem "Impressionist"
    LstConvert.AddItem "Nearest Color"
    LstConvert.AddItem "Threshold"
    LstConvert.AddItem "Ordered (Dithered)"
    LstConvert.AddItem "Error Diffusion (Dithered)"
    LstConvert.AddItem "Santos Enhanced (Dithered)"
    LstConvert.AddItem "Floyd-Steinberg (Dithered)"
    LstConvert.ListIndex = 7
    UpdateDescription
End Sub

Private Sub hsThreshold_Change()
    txtThreshold.Text = hsThreshold.Value
End Sub

Private Sub hsThreshold_Scroll()
    txtThreshold.Text = hsThreshold.Value
End Sub

Private Sub LstConvert_Click()
    UpdateDescription
End Sub

Private Sub LstConvert_KeyPress(KeyAscii As Integer)
    UpdateDescription
End Sub

Private Sub UpdateDescription()

    Dim l As String
    l = LstConvert.List(LstConvert.ListIndex)
    If l = "Component Color" Then
        lblDesc = "Non-dithered algorithm. A pixel is set to black only if each color component (red, green, and blue) is closer to black than white."
    ElseIf l = "Impressionist" Then
        lblDesc = "Non-dithered algorithm.  Special ranges are used for reduction: (0-24%, 50-74%) luminance is set to black, (25-49%, 75-100%) luminance is set to white."
    ElseIf l = "Nearest Color" Then
        lblDesc = "Non-dithered algorithm.  Colors with luminance below 50% are set to black, all others to white."
    ElseIf l = "Threshold" Then
        lblDesc = "Non-dithered algorithm.  Colors below a user-specified luminance threshold are set to black, all others to white."
    ElseIf l = "Error Diffusion (Dithered)" Then
        lblDesc = "Simple dithering algorithm.  Color conversion errors are bled from left-to-right in an attempt to better model color gradients."
    ElseIf l = "Floyd-Steinberg (Dithered)" Then
        lblDesc = "Complex dithering algorithm.  Conversion errors are bled in four directions to minimize visual artifacts. (Note: this is specifically a reduced-color bleed variant of the original Floyd-Steinberg algorithm.)"
    ElseIf l = "Ordered (Dithered)" Then
        lblDesc = "Simple dithering algorithm.  Geometric patterns are used to spread conversion errors across 4x4 patches of pixels."
    Else
        lblDesc = "Unique dithering method based off Manuel Augusto Santos's original ""Fast Graphics Filters"" implementation.  (Download link available from the Help -> About " & PROGRAMNAME & " menu)"
    End If
    
    'If threshold was selected, display the textbox and slider
    If LstConvert.ListIndex = 3 Then
        txtThreshold.Visible = True
        hsThreshold.Visible = True
    Else
        txtThreshold.Visible = False
        hsThreshold.Visible = False
    End If
    
End Sub

Private Sub txtThreshold_Change()
    If EntryValid(txtThreshold, hsThreshold.Min, hsThreshold.Max, False, False) Then
        hsThreshold.Value = val(txtThreshold)
    End If
End Sub

Private Sub txtThreshold_GotFocus()
    AutoSelectText txtThreshold
End Sub
