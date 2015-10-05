VERSION 5.00
Begin VB.Form dialog_ExportJP2 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " JPEG 2000 Export Options"
   ClientHeight    =   6585
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   12135
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
   ScaleHeight     =   439
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   809
   ShowInTaskbar   =   0   'False
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
      Left            =   6240
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   2550
      Width           =   5535
   End
   Begin PhotoDemon.sliderTextCombo sltQuality 
      Height          =   405
      Left            =   6120
      TabIndex        =   4
      Top             =   3120
      Width           =   5775
      _ExtentX        =   15055
      _ExtentY        =   873
      Min             =   1
      Max             =   256
      Value           =   16
      NotchPosition   =   1
   End
   Begin PhotoDemon.commandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   5
      Top             =   5835
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   1323
      BackColor       =   14802140
      dontAutoUnloadParent=   -1  'True
   End
   Begin PhotoDemon.fxPreviewCtl fxPreview 
      Height          =   5625
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9922
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
      Left            =   6240
      TabIndex        =   3
      Top             =   3600
      Width           =   1545
   End
   Begin VB.Label lblAfter 
      Alignment       =   1  'Right Justify
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
      Left            =   9480
      TabIndex        =   2
      Top             =   3600
      Width           =   1470
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "image compression ratio"
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
      Left            =   6000
      TabIndex        =   1
      Top             =   2160
      Width           =   2610
   End
End
Attribute VB_Name = "dialog_ExportJP2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'JPEG-2000 (JP2) Export Dialog
'Copyright 2012-2015 by Tanner Helland
'Created: 04/December/12
'Last updated: 14/February/14
'Last update: reworked layout to incorporate preview UC and more closely mimic the JPEG dialog
'
'Dialog for presenting the user a number of options related to JPEG-2000 exporting.  Obviously this feature
' relies on FreeImage, and JPEG-2000 support will be disabled if FreeImage cannot be found.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'The user input from the dialog
Private userAnswer As VbMsgBoxResult

'This form can be notified of the image being exported.  This may be used in the future to provide a preview.
Public imageBeingExported As pdImage

'When rendering the preview, we don't want to always re-request a copy of the main image.  Instead, we
' store one in this DIB (at the size of the preview) and simply re-use it when we need to render a preview.
Private origImageCopy As pdDIB

'The user's answer is returned via this property
Public Property Get DialogResult() As VbMsgBoxResult
    DialogResult = userAnswer
End Property

'QUALITY combo box - when adjusted, change the scroll bar to match
Private Sub CmbSaveQuality_Click()
    
    Select Case CmbSaveQuality.ListIndex
        
        Case 0
            sltQuality = 1
                
        Case 1
            sltQuality = 16
                
        Case 2
            sltQuality = 32
                
        Case 3
            sltQuality = 64
                
        Case 4
            sltQuality = 256
                
    End Select
    
End Sub

Private Sub cmdBar_CancelClick()
    userAnswer = vbCancel
    Me.Hide
End Sub

Private Sub cmdBar_OKClick()

    'Determine the compression ratio for the JPEG2000 wavelet transformation
    If sltQuality.IsValid Then
        g_JP2Compression = Abs(sltQuality)
    Else
        Exit Sub
    End If
     
    userAnswer = vbOK
    Me.Hide

End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    updatePreview
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub fxPreview_ViewportChanged()
    updatePreview
End Sub

Private Sub sltQuality_Change()
    updateComboBox
    updatePreview
End Sub

'Used to keep the "compression ratio" text box, scroll bar, and combo box in sync
Private Sub updateComboBox()
    
    Select Case sltQuality.Value
        
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
Public Sub showDialog()

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
    
    'Apply translations and visual themes
    MakeFormPretty Me
    
    'Make a copy of the current image
    Set origImageCopy = New pdDIB
    imageBeingExported.getCompositedImage origImageCopy, True
    
    'Update the preview
    updatePreview
    
    'Display the dialog
    ShowPDDialog vbModal, Me, True

End Sub

'Render a new JPEG-2000 preview
Private Sub updatePreview()

    If cmdBar.previewsAllowed And g_ImageFormats.FreeImageEnabled And sltQuality.IsValid Then
        
        'Start by retrieving the relevant portion of the image, according to the preview window
        Dim tmpSafeArray As SAFEARRAY2D
        previewNonStandardImage tmpSafeArray, origImageCopy, fxPreview
        
        'The public workingDIB object now contains the relevant portion of the preview window.  Use that to
        ' obtain a JPEG-ified version of the image data.
        fillDIBWithJP2Version workingDIB, workingDIB, Abs(sltQuality.Value)
        
        'Paint the final image to screen and release all temporary objects
        finalizeNonstandardPreview fxPreview
        
    End If

End Sub
