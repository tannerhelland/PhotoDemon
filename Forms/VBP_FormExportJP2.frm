VERSION 5.00
Begin VB.Form dialog_ExportJP2 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " JPEG 2000 Export Options"
   ClientHeight    =   3180
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
   ScaleHeight     =   212
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   489
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   4320
      TabIndex        =   0
      Top             =   2550
      Width           =   1365
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   5790
      TabIndex        =   1
      Top             =   2550
      Width           =   1365
   End
   Begin VB.HScrollBar hsQuality 
      Height          =   330
      LargeChange     =   5
      Left            =   600
      Max             =   256
      Min             =   1
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1245
      Value           =   16
      Width           =   5295
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
      Width           =   6135
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
      Left            =   6000
      MaxLength       =   3
      TabIndex        =   2
      Text            =   "90"
      Top             =   1230
      Width           =   735
   End
   Begin VB.Label lblBackground 
      Height          =   855
      Left            =   0
      TabIndex        =   8
      Top             =   2400
      Width           =   7335
   End
   Begin VB.Label lblBefore 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "high quality, large file"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   600
      TabIndex        =   7
      Top             =   1680
      Width           =   1545
   End
   Begin VB.Label lblAfter 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "low quality, small file"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   4410
      TabIndex        =   6
      Top             =   1680
      Width           =   1470
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "image compression ratio:"
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
      Width           =   2700
   End
End
Attribute VB_Name = "dialog_ExportJP2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'JPEG-2000 (JP2) Export Dialog
'Copyright ©2011-2013 by Tanner Helland
'Created: 04/December/12
'Last updated: 04/December/12
'Last update: abandoned my attempt to merge this with the JPEG export form; it's way easier (and less code, surprisingly)
'             to just give this its own dialog.
'
'Dialog for presenting the user a number of options related to JPEG-2000 exporting.  Obviously this feature
' relies on FreeImage, and JPEG-2000 support will be disabled if FreeImage cannot be found.
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
            hsQuality.Value = 1
                
        Case 1
            hsQuality.Value = 16
                
        Case 2
            hsQuality = 32
                
        Case 3
            hsQuality = 64
                
        Case 4
            hsQuality = 256
                
    End Select
    
End Sub

'CANCEL button
Private Sub CmdCancel_Click()
    
    userAnswer = vbCancel
    Me.Hide
    
End Sub

'OK button
Private Sub cmdOK_Click()
        
    'Determine the compression ratio for the JPEG2000 wavelet transformation
    Select Case CmbSaveQuality.ListIndex
            
        Case 0
            g_JP2Compression = 1
        Case 1
            g_JP2Compression = 16
        Case 2
            g_JP2Compression = 32
        Case 3
            g_JP2Compression = 64
        Case 4
            g_JP2Compression = 256
        Case 5
            If EntryValid(txtQuality, hsQuality.Min, hsQuality.Max) Then
                g_JP2Compression = Abs(hsQuality.Value)
            Else
                AutoSelectText txtQuality
                Exit Sub
            End If
    End Select
     
    userAnswer = vbOK
    Me.Hide
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
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

'Used to keep the "compression ratio" text box, scroll bar, and combo box in sync
Private Sub updateComboBox()
    
    Select Case hsQuality.Value
        
        Case 1
            If CmbSaveQuality.ListIndex <> 0 Then CmbSaveQuality.ListIndex = 0
                
        Case 16
            If CmbSaveQuality.ListIndex <> 1 Then CmbSaveQuality.ListIndex = 1
                
        Case 32
            If CmbSaveQuality.ListIndex <> 2 Then CmbSaveQuality.ListIndex = 2
                
        Case 64
            If CmbSaveQuality.ListIndex <> 3 Then CmbSaveQuality.ListIndex = 3
                
        Case 256
            If CmbSaveQuality.ListIndex <> 4 Then CmbSaveQuality.ListIndex = 4
                
        Case Else
            If CmbSaveQuality.ListIndex <> 5 Then CmbSaveQuality.ListIndex = 5
                
    End Select
    
End Sub

'The ShowDialog routine presents the user with this form.
Public Sub ShowDialog()

    'Provide a default answer of "cancel" (in the event that the user clicks the "x" button in the top-right)
    userAnswer = vbCancel
    
    'Make sure that the proper cursor is set
    Screen.MousePointer = 0
    
    'Populate the quality drop-down box with presets corresponding to the JPEG-2000 file format
    CmbSaveQuality.Clear
    CmbSaveQuality.AddItem " Lossless (1:1)", 0
    CmbSaveQuality.AddItem " Low compression, good image quality (16:1)", 1
    CmbSaveQuality.AddItem " Moderate compression, medium image quality (32:1)", 2
    CmbSaveQuality.AddItem " High compression, poor image quality (64:1)", 3
    CmbSaveQuality.AddItem " Super compression, very poor image quality (256:1)", 4
    CmbSaveQuality.AddItem " Custom ratio (X:1)", 5
    CmbSaveQuality.ListIndex = 0
    
    Message "Waiting for user to specify JPEG-2000 export options... "
        
    'Assign the system hand cursor to all relevant objects
    makeFormPretty Me
    
    'If fancy fonts are being used, increase the horizontal scroll bar height by one pixel equivalent (to make it fit better)
    If g_UseFancyFonts Then hsQuality.Height = 23 Else hsQuality.Height = 22

    'Display the dialog
    Me.Show vbModal, FormMain

End Sub
