VERSION 5.00
Begin VB.Form FormJPEG 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " JPEG Export Options"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7245
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
   ScaleHeight     =   328
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   483
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cmbSubsample 
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
      Left            =   3480
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   3255
      Width           =   3495
   End
   Begin VB.CheckBox chkSubsample 
      Appearance      =   0  'Flat
      Caption         =   " use custom subsampling:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   480
      TabIndex        =   9
      Top             =   3240
      Width           =   2895
   End
   Begin VB.CheckBox chkProgressive 
      Appearance      =   0  'Flat
      Caption         =   " use progressive encoding"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   480
      TabIndex        =   8
      Top             =   2760
      Width           =   6375
   End
   Begin VB.CheckBox chkOptimize 
      Appearance      =   0  'Flat
      Caption         =   " optimize (takes slightly longer, but results in a smaller file)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   480
      TabIndex        =   7
      Top             =   2280
      Value           =   1  'Checked
      Width           =   6375
   End
   Begin VB.HScrollBar hsQuality 
      Height          =   285
      LargeChange     =   5
      Left            =   480
      Max             =   99
      Min             =   1
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1200
      Value           =   90
      Width           =   5655
   End
   Begin VB.ComboBox CmbSaveQuality 
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
      Left            =   480
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   600
      Width           =   6495
   End
   Begin VB.TextBox txtQuality 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   345
      Left            =   6240
      MaxLength       =   2
      TabIndex        =   2
      Text            =   "90"
      Top             =   1170
      Width           =   735
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   5640
      TabIndex        =   1
      Top             =   4200
      Width           =   1245
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   4320
      TabIndex        =   0
      Top             =   4200
      Width           =   1245
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000D&
      X1              =   8
      X2              =   472
      Y1              =   264
      Y2              =   264
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "advanced settings:"
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
      Left            =   360
      TabIndex        =   6
      Top             =   1800
      Width           =   1965
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "image quality:"
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
      Left            =   360
      TabIndex        =   4
      Top             =   240
      Width           =   1515
   End
End
Attribute VB_Name = "FormJPEG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'JPEG Export interface
'Copyright ©2000-2012 by Tanner Helland
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

'QUALITY combo box - when adjusted, change the scroll bar to match
Private Sub CmbSaveQuality_Click()
    
    Select Case CmbSaveQuality.ListIndex
    
        Case 0
            hsQuality.Value = 99
            
        Case 1
            hsQuality.Value = 92
            
        Case 2
            hsQuality = 80
            
        Case 3
            hsQuality = 65
            
        Case 4
            hsQuality = 40
            
    End Select
    
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
    
    'Determine any extra flags based on the advanced settings
    Dim JPEGFlags As Long
    JPEGFlags = 0
    
    'Optimize
    If CBool(chkOptimize) Then JPEGFlags = JPEGFlags Or JPEG_OPTIMIZE
    
    'Progressive scan
    If CBool(chkProgressive) Then JPEGFlags = JPEGFlags Or JPEG_PROGRESSIVE
    
    'Subsampling
    If CBool(chkSubsample) Then
    
        Select Case cmbSubsample.ListIndex
        
            Case 0
                JPEGFlags = JPEGFlags Or JPEG_SUBSAMPLING_444
            Case 1
                JPEGFlags = JPEGFlags Or JPEG_SUBSAMPLING_422
            Case 2
                JPEGFlags = JPEGFlags Or JPEG_SUBSAMPLING_420
            Case 3
                JPEGFlags = JPEGFlags Or JPEG_SUBSAMPLING_411
                
        End Select
        
    End If
    
    Me.Visible = False
        
    'Pass control to PhotoDemon_SaveImage
    ' (if the save function fails for some reason, return the save dialog as canceled so the user can try again)
    saveDialogCanceled = Not PhotoDemon_SaveImage(CurrentImage, SaveFileName, False, JPEGQuality, JPEGFlags)

    Unload Me
    
End Sub

'LOAD form
Private Sub Form_Load()
    
    'Make sure that the proper cursor is set
    Screen.MousePointer = 0
    
    'I've found that pre-existing combo box entries are more user-friendly
    CmbSaveQuality.AddItem "Perfect (99)", 0
    CmbSaveQuality.AddItem "Excellent (92)", 1
    CmbSaveQuality.AddItem "Good (80)", 2
    CmbSaveQuality.AddItem "Average (65)", 3
    CmbSaveQuality.AddItem "Low (40)", 4
    CmbSaveQuality.AddItem "Custom value", 5
    CmbSaveQuality.ListIndex = 1
    Message "Waiting for user to specify JPEG export options... "
    
    'Populate the custom subsampling combo box as well
    cmbSubsample.AddItem "4:4:4 (best quality)", 0
    cmbSubsample.AddItem "4:2:2 (good quality)", 1
    cmbSubsample.AddItem "4:2:0 (moderate quality)", 2
    cmbSubsample.AddItem "4:1:1 (low quality)", 3
    cmbSubsample.ListIndex = 2
    
    'If FreeImage is not available, disable all the advanced settings
    If Not imageFormats.FreeImageEnabled Then
        chkOptimize.Enabled = False
        chkProgressive.Enabled = False
        chkSubsample.Enabled = False
        cmbSubsample.AddItem "n/a", 4
        cmbSubsample.ListIndex = 4
        lblTitle(1).Caption = "advanced settings require the FreeImage plugin"
    End If
    
    'Mark the form as having NOT been canceled
    saveDialogCanceled = False
            
    'Assign the system hand cursor to all relevant objects
    makeFormPretty Me
    
End Sub

Private Sub hsQuality_Change()
    txtQuality.Text = hsQuality.Value
    updateComboBox
End Sub

Private Sub hsQuality_Scroll()
    txtQuality.Text = hsQuality.Value
    updateComboBox
End Sub

Private Sub txtQuality_Change()
    If EntryValid(txtQuality, hsQuality.Min, hsQuality.Max, False, False) Then hsQuality.Value = Val(txtQuality)
End Sub

Private Sub txtQuality_GotFocus()
    AutoSelectText txtQuality
End Sub

'Used to keep the "image quality" text box, scroll bar, and combo box in sync
Private Sub updateComboBox()

    Select Case hsQuality.Value
    
        Case 40
            If CmbSaveQuality.ListIndex <> 4 Then CmbSaveQuality.ListIndex = 4
            
        Case 65
            If CmbSaveQuality.ListIndex <> 3 Then CmbSaveQuality.ListIndex = 3
            
        Case 80
            If CmbSaveQuality.ListIndex <> 2 Then CmbSaveQuality.ListIndex = 2
            
        Case 92
            If CmbSaveQuality.ListIndex <> 1 Then CmbSaveQuality.ListIndex = 1
            
        Case 99
            If CmbSaveQuality.ListIndex <> 0 Then CmbSaveQuality.ListIndex = 0
            
        Case Else
            If CmbSaveQuality.ListIndex <> 5 Then CmbSaveQuality.ListIndex = 5
            
    End Select
    
End Sub
