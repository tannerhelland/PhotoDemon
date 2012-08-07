VERSION 5.00
Begin VB.Form FormJPEG 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " JPEG Options"
   ClientHeight    =   2325
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4560
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
   ScaleHeight     =   155
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   304
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.HScrollBar hsQuality 
      Height          =   285
      LargeChange     =   5
      Left            =   960
      Max             =   100
      Min             =   1
      MouseIcon       =   "VBP_FormJPEG.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1200
      Value           =   90
      Width           =   3375
   End
   Begin VB.ComboBox CmbSaveQuality 
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
      Height          =   330
      Left            =   360
      MouseIcon       =   "VBP_FormJPEG.frx":0152
      MousePointer    =   99  'Custom
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   600
      Width           =   3975
   End
   Begin VB.TextBox txtQuality 
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
      Height          =   285
      Left            =   360
      TabIndex        =   2
      Text            =   "90"
      Top             =   1200
      Width           =   495
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3120
      MouseIcon       =   "VBP_FormJPEG.frx":02A4
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   1800
      Width           =   1125
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1920
      MouseIcon       =   "VBP_FormJPEG.frx":03F6
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   1800
      Width           =   1125
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "JPEG Save Quality:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   240
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   1605
   End
End
Attribute VB_Name = "FormJPEG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'JPEG Export interface
'Copyright ©2001-2012 by Tanner Helland
'Created: 5/8/00
'Last updated: 12/June/12
'Last update: added FreeImage option for saving JPEGs.  John's class is great, but it's largely untested.
'             FreeImage also offers more advanced options for compression, and it's faster to boot.  So,
'             if FreeImage is found it will be used, otherwise the job falls to John's class.
'
'Form for giving the user a couple options for exporting JPEGs.
'
'***************************************************************************

Option Explicit


'QUALITY combo box - when adjusted, enable or disable the custom slider
Private Sub CmbSaveQuality_Click()
    ShowAdditionalControls
End Sub

Private Sub CmbSaveQuality_KeyDown(KeyCode As Integer, Shift As Integer)
    ShowAdditionalControls
End Sub

'CANCEL button
Private Sub CmdCancel_Click()
    saveDialogCanceled = True
    Message "Image save canceled. "
    Unload Me
End Sub

'OK button
Private Sub CmdOK_Click()
    
    Me.Visible = False
    
    Message "Preparing image..."

    'Determine the compression quality for the quantization tables
    Dim jpegQuality As Long
    Select Case CmbSaveQuality.ListIndex
        Case 0
            jpegQuality = 99
        Case 1
            jpegQuality = 92
        Case 2
            jpegQuality = 80
        Case 3
            jpegQuality = 65
        Case 4
            jpegQuality = 40
        Case 5
            If EntryValid(txtQuality, hsQuality.Min, hsQuality.Max) Then
                jpegQuality = hsQuality.Value
            Else
                AutoSelectText txtQuality
                Exit Sub
            End If
    End Select
    
    Me.Visible = False
    
    'I implement two separate save functions for JPEG images.  While I greatly appreciate John Korejwa's native
    ' VB JPEG encoder, it's not nearly as robust, or fast, or tested as the FreeImage implementation.  As such,
    ' PhotoDemon will default to FreeImage if it's available; otherwise John's JPEG class will be used.
    
    'Also, if the save function fails for some reason, return the save dialog as canceled (so the user can try again)
    saveDialogCanceled = Not PhotoDemon_SaveImage(CurrentImage, SaveFileName, False, jpegQuality)

    Unload Me
    
End Sub

Private Sub Form_Load()
    'I've found that pre-existing combo box entries are more user-friendly
    CmbSaveQuality.AddItem "Perfect (99)", 0
    CmbSaveQuality.AddItem "Excellent (92)", 1
    CmbSaveQuality.AddItem "Good (80)", 2
    CmbSaveQuality.AddItem "Average (65)", 3
    CmbSaveQuality.AddItem "Poor (40)", 4
    CmbSaveQuality.AddItem "Custom value", 5
    CmbSaveQuality.ListIndex = 1
    Message "Waiting for user to specify JPEG export options... "
    
    'Mark the form as having NOT been canceled
    saveDialogCanceled = False
End Sub

Private Sub hsQuality_Change()
    txtQuality.Text = hsQuality.Value
End Sub

Private Sub hsQuality_Scroll()
    txtQuality.Text = hsQuality.Value
End Sub

Private Sub txtQuality_Change()
    If EntryValid(txtQuality, hsQuality.Min, hsQuality.Max, False, False) Then hsQuality.Value = val(txtQuality)
End Sub

Private Sub txtQuality_GotFocus()
    AutoSelectText txtQuality
End Sub

Private Sub ShowAdditionalControls()
    If CmbSaveQuality.ListIndex = 5 Then
        txtQuality.Visible = True
        hsQuality.Visible = True
    Else
        txtQuality.Visible = False
        hsQuality.Visible = False
    End If
End Sub
