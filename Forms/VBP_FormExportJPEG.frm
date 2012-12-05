VERSION 5.00
Begin VB.Form dialog_ExportJPEG 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " JPEG Export Options"
   ClientHeight    =   5385
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7335
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
   ScaleHeight     =   359
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   489
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdShowHide 
      Caption         =   "<<  Hide advanced settings"
      Height          =   495
      Left            =   240
      TabIndex        =   12
      Top             =   4680
      Width           =   2685
   End
   Begin VB.CheckBox chkThumbnail 
      Appearance      =   0  'Flat
      Caption         =   " embed thumbnail image"
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
      Left            =   600
      TabIndex        =   11
      ToolTipText     =   $"VBP_FormExportJPEG.frx":0000
      Top             =   2520
      Width           =   6375
   End
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
      Left            =   3600
      Style           =   2  'Dropdown List
      TabIndex        =   10
      ToolTipText     =   "Subsampling affects the way the JPEG encoder compresses image luminance.  4:2:0 (moderate) is the default value."
      Top             =   3495
      Width           =   3375
   End
   Begin VB.CheckBox chkSubsample 
      Appearance      =   0  'Flat
      Caption         =   " use specific subsampling:"
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
      Left            =   600
      TabIndex        =   9
      ToolTipText     =   "Subsampling affects the way the JPEG encoder compresses image luminance.  4:2:0 (moderate) is the default value."
      Top             =   3480
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
      Left            =   600
      TabIndex        =   8
      ToolTipText     =   $"VBP_FormExportJPEG.frx":008E
      Top             =   3000
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
      Left            =   600
      TabIndex        =   7
      ToolTipText     =   $"VBP_FormExportJPEG.frx":015E
      Top             =   2040
      Value           =   1  'Checked
      Width           =   6375
   End
   Begin VB.HScrollBar hsQuality 
      Height          =   330
      LargeChange     =   5
      Left            =   2520
      Max             =   99
      Min             =   1
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   645
      Value           =   90
      Width           =   3735
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
      Left            =   600
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   630
      Width           =   1815
   End
   Begin VB.TextBox txtQuality 
      Alignment       =   2  'Center
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
      Left            =   6360
      MaxLength       =   2
      TabIndex        =   2
      Text            =   "90"
      Top             =   630
      Width           =   735
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   5880
      TabIndex        =   1
      Top             =   4680
      Width           =   1245
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   4560
      TabIndex        =   0
      Top             =   4680
      Width           =   1245
   End
   Begin VB.Line lineSeparator 
      BorderColor     =   &H8000000F&
      X1              =   8
      X2              =   480
      Y1              =   296
      Y2              =   296
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "advanced JPEG settings:"
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
      Top             =   1560
      Width           =   2580
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
Attribute VB_Name = "dialog_ExportJPEG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'JPEG Export Dialog
'Copyright ©2000-2012 by Tanner Helland
'Created: 5/8/00
'Last updated: 03/December/12
'Last update: converted this into a true "dialog", in that it can be called from anywhere, and it will return
'              "OK" or "Cancel" (as type VBMsgBoxResult) if the user hit OK or Cancel.  If OK was pressed, three
'              global variables - g_JPEGQuality, g_JPEGFlags, and g_JPEGThumbnail - will be set with the user's
'              answers.  These can then be queried by the calling function as needed.
'
'Dialog for preseting the user a number of options for related to JPEG exporting.  The various advanced features
' rely on FreeImage for implementation, and will be disabled if FreeImage cannot be found.
'
'***************************************************************************

Option Explicit

'The user input from the dialog
Private userAnswer As VbMsgBoxResult

'The pdImage object being exported
Private imageBeingExported As pdImage

'The user's answer is returned via this property
Public Property Get DialogResult() As VbMsgBoxResult
    DialogResult = userAnswer
End Property

'This form can be notified of the image being exported.  This may be used in the future to provide a preview.
Public Property Let srcImage(srcImage As pdImage)
    imageBeingExported = srcImage
End Property

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
    
    userAnswer = vbCancel
    Me.Hide
    
End Sub

'OK button
Private Sub CmdOK_Click()
        
    'Determine the compression quality for the quantization tables
    Select Case CmbSaveQuality.ListIndex
        Case 0
            g_JPEGQuality = 99
        Case 1
            g_JPEGQuality = 92
        Case 2
            g_JPEGQuality = 80
        Case 3
            g_JPEGQuality = 65
        Case 4
            g_JPEGQuality = 40
        Case 5
            If EntryValid(txtQuality, hsQuality.Min, hsQuality.Max) Then
                g_JPEGQuality = hsQuality.Value
            Else
                AutoSelectText txtQuality
                Exit Sub
            End If
    End Select
        
    'Determine any extra flags based on the advanced settings
    g_JPEGFlags = 0
        
    'Optimize
    If CBool(chkOptimize) Then g_JPEGFlags = g_JPEGFlags Or JPEG_OPTIMIZE
        
    'Progressive scan
    If CBool(chkProgressive) Then g_JPEGFlags = g_JPEGFlags Or JPEG_PROGRESSIVE
        
    'Subsampling
    If CBool(chkSubsample) Then
    
        Select Case cmbSubsample.ListIndex
            
            Case 0
                g_JPEGFlags = g_JPEGFlags Or JPEG_SUBSAMPLING_444
            Case 1
                g_JPEGFlags = g_JPEGFlags Or JPEG_SUBSAMPLING_422
            Case 2
                g_JPEGFlags = g_JPEGFlags Or JPEG_SUBSAMPLING_420
            Case 3
                g_JPEGFlags = g_JPEGFlags Or JPEG_SUBSAMPLING_411
                    
        End Select
            
    End If
        
    'Finally, determine whether or not a thumbnail version of the file should be embedded inside
    If CBool(chkThumbnail) Then g_JPEGThumbnail = 1 Else g_JPEGThumbnail = 0
     
    userAnswer = vbOK
    Me.Hide
    
End Sub

'Show or hide the advanced settings per the user's command
Private Sub cmdShowHide_Click()

    toggleAdvancedSettings
    
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

'Show or hide the advanced settings per the user's command
Private Sub toggleAdvancedSettings()

    If cmdShowHide.Caption = "<<  Hide advanced settings" Then
    
        'Re-caption the button
        cmdShowHide.Caption = "Show advanced settings  >>"
    
        'Hide all advanced options
        lblTitle(1).Visible = False
        chkOptimize.Visible = False
        chkThumbnail.Visible = False
        chkProgressive.Visible = False
        chkSubsample.Visible = False
        cmbSubsample.Visible = False
    
        'Move all other controls accordingly
        lineSeparator.y1 = hsQuality.Top + 48
        lineSeparator.y2 = lineSeparator.y1
        cmdShowHide.Top = lineSeparator.y1 + 16
        CmdOK.Top = cmdShowHide.Top
        CmdCancel.Top = CmdOK.Top
    
    Else
    
        'Re-caption the button
        cmdShowHide.Caption = "<<  Hide advanced settings"
    
        'Show all advanced options
        lblTitle(1).Visible = True
        chkOptimize.Visible = True
        chkThumbnail.Visible = True
        chkProgressive.Visible = True
        chkSubsample.Visible = True
        cmbSubsample.Visible = True
        
        'Move all other controls accordingly
        lineSeparator.y1 = chkSubsample.Top + 48
        lineSeparator.y2 = lineSeparator.y1
        cmdShowHide.Top = lineSeparator.y1 + 16
        CmdOK.Top = cmdShowHide.Top
        CmdCancel.Top = CmdOK.Top
    
    End If

    'Change the form size to match
    Dim formSizeDiff As Long
    Me.ScaleMode = vbTwips
    formSizeDiff = Me.Height - Me.ScaleHeight
    
    Me.Height = formSizeDiff + cmdShowHide.Top + cmdShowHide.Height + Abs(lineSeparator.y1 - cmdShowHide.Top)
    
    Me.ScaleMode = vbPixels

End Sub

'The ShowDialog routine presents the user with this form.
Public Sub ShowDialog(Optional ByVal showAdvanced As Boolean = False)

    'Provide a default answer of "cancel" (in the event that the user clicks the "x" button in the top-right)
    userAnswer = vbCancel
    
    'Make sure that the proper cursor is set
    Screen.MousePointer = 0
    
    'Populate the quality drop-down box with presets corresponding to the JPEG format
    CmbSaveQuality.Clear
    CmbSaveQuality.AddItem " Perfect (99)", 0
    CmbSaveQuality.AddItem " Excellent (92)", 1
    CmbSaveQuality.AddItem " Good (80)", 2
    CmbSaveQuality.AddItem " Average (65)", 3
    CmbSaveQuality.AddItem " Low (40)", 4
    CmbSaveQuality.AddItem " Custom value", 5
    CmbSaveQuality.ListIndex = 1
    Message "Waiting for user to specify JPEG export options... "
        
    'Populate the custom subsampling combo box as well
    cmbSubsample.Clear
    cmbSubsample.AddItem " 4:4:4 (best quality)", 0
    cmbSubsample.AddItem " 4:2:2 (good quality)", 1
    cmbSubsample.AddItem " 4:2:0 (moderate quality)", 2
    cmbSubsample.AddItem " 4:1:1 (low quality)", 3
    cmbSubsample.ListIndex = 2
    
    'If FreeImage is not available, disable all the advanced settings
    If Not imageFormats.FreeImageEnabled Then
        chkOptimize.Enabled = False
        chkProgressive.Enabled = False
        chkSubsample.Enabled = False
        chkThumbnail.Enabled = False
        cmbSubsample.AddItem "n/a", 4
        cmbSubsample.ListIndex = 4
        cmbSubsample.Enabled = False
        lblTitle(1).Caption = "advanced settings require the FreeImage plugin"
    End If
        
    'Hide the advanced settings unless the user has specifically requested otherwise
    If Not showAdvanced Then toggleAdvancedSettings
                
    'Assign the system hand cursor to all relevant objects
    makeFormPretty Me
    
    'If fancy fonts are being used, increase the horizontal scroll bar height by one pixel equivalent (to make it fit better)
    If useFancyFonts Then hsQuality.Height = 23 Else hsQuality.Height = 22

    'Display the dialog
    Me.Show vbModal, FormMain

End Sub
