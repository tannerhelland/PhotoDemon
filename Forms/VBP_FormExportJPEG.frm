VERSION 5.00
Begin VB.Form dialog_ExportJPEG 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " JPEG Export Options"
   ClientHeight    =   5715
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7440
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
   ScaleHeight     =   381
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   496
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.smartCheckBox chkOptimize 
      Height          =   540
      Left            =   600
      TabIndex        =   8
      Top             =   1920
      Width           =   3105
      _ExtentX        =   5477
      _ExtentY        =   953
      Caption         =   "optimize compression tables"
      Value           =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      TabIndex        =   0
      Top             =   5070
      Width           =   1365
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5910
      TabIndex        =   1
      Top             =   5070
      Width           =   1365
   End
   Begin VB.CommandButton cmdShowHide 
      Caption         =   " Hide advanced settings"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   5070
      Width           =   2685
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
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   5
      ToolTipText     =   "Subsampling affects the way the JPEG encoder compresses image luminance.  4:2:0 (moderate) is the default value."
      Top             =   3930
      Width           =   5415
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
      TabIndex        =   2
      Top             =   630
      Width           =   2055
   End
   Begin PhotoDemon.smartCheckBox chkThumbnail 
      Height          =   540
      Left            =   600
      TabIndex        =   9
      Top             =   2400
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   953
      Caption         =   "embed thumbnail image"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PhotoDemon.smartCheckBox chkProgressive 
      Height          =   540
      Left            =   600
      TabIndex        =   10
      Top             =   2880
      Width           =   2835
      _ExtentX        =   5001
      _ExtentY        =   953
      Caption         =   "use progressive encoding"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PhotoDemon.smartCheckBox chkSubsample 
      Height          =   540
      Left            =   600
      TabIndex        =   11
      ToolTipText     =   "Subsampling affects the way the JPEG encoder compresses image luminance.  4:2:0 (moderate) is the default value."
      Top             =   3360
      Width           =   2820
      _ExtentX        =   4974
      _ExtentY        =   953
      Caption         =   "use specific subsampling:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PhotoDemon.sliderTextCombo sltQuality 
      Height          =   495
      Left            =   2640
      TabIndex        =   12
      Top             =   585
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   873
      Min             =   1
      Max             =   99
      Value           =   90
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblBackground 
      Height          =   855
      Left            =   -120
      TabIndex        =   7
      Top             =   4920
      Width           =   7575
   End
   Begin VB.Line lineSeparator 
      BorderColor     =   &H8000000F&
      X1              =   8
      X2              =   480
      Y1              =   328
      Y2              =   328
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
      TabIndex        =   4
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
      TabIndex        =   3
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
'Copyright ©2000-2013 by Tanner Helland
'Created: 5/8/00
'Last updated: 03/December/12
'Last update: converted this into a true "dialog", in that it can be called from anywhere, and it will return
'              "OK" or "Cancel" (as type vbMsgBoxResult) if the user hit OK or Cancel.  If OK was pressed, three
'              global variables - g_JPEGQuality, g_JPEGFlags, and g_JPEGThumbnail - will be set with the user's
'              answers.  These can then be queried by external functions as needed.
'
'Dialog for preseting the user a number of options for related to JPEG exporting.  The various advanced features
' rely on FreeImage for implementation, and will be disabled if FreeImage cannot be found.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'The user input from the dialog
Private userAnswer As VbMsgBoxResult

'The pdImage object being exported
Private imageBeingExported As pdImage

'Whether to show or hide the advanced settings
Private showAdvanced As Boolean

'Custom tooltip class allows for things like multiline, theming, and multiple monitor support
Dim m_ToolTip As clsToolTip

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
            sltQuality.Value = 99
                
        Case 1
            sltQuality.Value = 92
                
        Case 2
            sltQuality = 80
                
        Case 3
            sltQuality = 65
                
        Case 4
            sltQuality = 40
                
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
    If sltQuality.IsValid Then
        g_JPEGQuality = sltQuality.Value
    Else
        Exit Sub
    End If
            
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

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'Used to keep the "image quality" text box, scroll bar, and combo box in sync
Private Sub updateComboBox()
    
    Select Case sltQuality.Value
        
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

    showAdvanced = Not showAdvanced

    If showAdvanced Then
    
        'Re-caption the button
        cmdShowHide.Caption = g_Language.TranslateMessage("Show advanced settings") & "  >>"
    
        'Hide all advanced options
        lblTitle(1).Visible = False
        chkOptimize.Visible = False
        chkThumbnail.Visible = False
        chkProgressive.Visible = False
        chkSubsample.Visible = False
        cmbSubsample.Visible = False
    
        'Move all other controls accordingly
        lineSeparator.y1 = sltQuality.Top + fixDPI(48)
        lineSeparator.y2 = lineSeparator.y1
        lblBackground.Top = lineSeparator.y1
        cmdShowHide.Top = lineSeparator.y1 + fixDPI(10)
        CmdOK.Top = cmdShowHide.Top
        cmdCancel.Top = CmdOK.Top
    
    Else
    
        'Re-caption the button
        cmdShowHide.Caption = "<<  " & g_Language.TranslateMessage("Hide advanced settings")
    
        'Show all advanced options
        lblTitle(1).Visible = True
        chkOptimize.Visible = True
        chkThumbnail.Visible = True
        chkProgressive.Visible = True
        chkSubsample.Visible = True
        cmbSubsample.Visible = True
        
        'Move all other controls accordingly
        lineSeparator.y1 = cmbSubsample.Top + fixDPI(48)
        lineSeparator.y2 = lineSeparator.y1
        lblBackground.Top = lineSeparator.y1
        cmdShowHide.Top = lineSeparator.y1 + fixDPI(10)
        CmdOK.Top = cmdShowHide.Top
        cmdCancel.Top = CmdOK.Top
    
    End If
    
    'Change the form size to match
    Dim formSizeDiff As Long
    Me.ScaleMode = vbTwips
    formSizeDiff = Me.Height - Me.ScaleHeight
    
    Me.Height = formSizeDiff + cmdShowHide.Top + cmdShowHide.Height + Abs(lineSeparator.y1 - cmdShowHide.Top)
    
    Me.ScaleMode = vbPixels

End Sub

'The ShowDialog routine presents the user with this form.
Public Sub showDialog(Optional ByVal showAdvanced As Boolean = False)

    showAdvanced = False

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
    If Not g_ImageFormats.FreeImageEnabled Then
        chkOptimize.Enabled = False
        chkProgressive.Enabled = False
        chkSubsample.Enabled = False
        chkThumbnail.Enabled = False
        cmbSubsample.AddItem "n/a", 4
        cmbSubsample.ListIndex = 4
        cmbSubsample.Enabled = False
        lblTitle(1).Caption = g_Language.TranslateMessage("advanced settings require the FreeImage plugin")
    End If
        
    'Apply some tooltips manually (so the translation engine can find them)
    chkOptimize.ToolTipText = g_Language.TranslateMessage("Optimization is highly recommended.  This option allows the JPEG encoder to compute an optimal Huffman coding table for the file.  It does not affect image quality - only file size.")
    chkThumbnail.ToolTipText = g_Language.TranslateMessage("Embedded thumbnails increase file size, but they help previews of the image appear more quickly in other software (e.g. Windows Explorer).")
    chkProgressive.ToolTipText = g_Language.TranslateMessage("Progressive encoding is sometimes used for JPEG files that will be used on the Internet.  It saves the image in three steps, which can be used to gradually fade-in the image on a slow Internet connection.")
    
    'Hide the advanced settings unless the user has specifically requested otherwise
    'If Not showAdvanced Then toggleAdvancedSettings
    toggleAdvancedSettings
                
    'Assign the system hand cursor to all relevant objects
    Set m_ToolTip = New clsToolTip
    makeFormPretty Me, m_ToolTip
    
    'Display the dialog
    showPDDialog vbModal, Me

End Sub

Private Sub sltQuality_Change()
    updateComboBox
End Sub
