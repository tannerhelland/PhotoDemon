VERSION 5.00
Begin VB.Form FormJPEG 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " JPEG Options"
   ClientHeight    =   2445
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
   ScaleHeight     =   163
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   304
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.HScrollBar hsQuality 
      Height          =   285
      LargeChange     =   5
      Left            =   360
      Max             =   100
      Min             =   1
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
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   600
      Width           =   3975
   End
   Begin VB.TextBox txtQuality 
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
      Height          =   330
      Left            =   3840
      TabIndex        =   2
      Text            =   "90"
      Top             =   1185
      Width           =   495
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3240
      TabIndex        =   1
      Top             =   1920
      Width           =   1125
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2040
      TabIndex        =   0
      Top             =   1920
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
'Last updated: 09/September/12
'Last update: moved decision-making about which JPEG export method to use to the PhotoDemon_SaveImage function.
'              It makes more sense there, because then ANY function that needs to save can benefit from the wisdom of
'              the automatic JPEG export mechanism selection.
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
    Dim JPEGQuality As Long
    Select Case CmbSaveQuality.ListIndex
        Case 0
            JPEGQuality = 99
        Case 1
            JPEGQuality = 92
        Case 2
            JPEGQuality = 80
        Case 3
            JPEGQuality = 65
        Case 4
            JPEGQuality = 40
        Case 5
            If EntryValid(txtQuality, hsQuality.Min, hsQuality.Max) Then
                JPEGQuality = hsQuality.Value
            Else
                AutoSelectText txtQuality
                Exit Sub
            End If
    End Select
    
    Me.Visible = False
        
    'Pass control to PhotoDemon_SaveImage
    ' (if the save function fails for some reason, return the save dialog as canceled so the user can try again)
    saveDialogCanceled = Not PhotoDemon_SaveImage(CurrentImage, SaveFileName, False, JPEGQuality)

    Unload Me
    
End Sub

'LOAD form
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
    
    'Assign the system hand cursor to all relevant objects
    setHandCursorForAll Me
    
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
