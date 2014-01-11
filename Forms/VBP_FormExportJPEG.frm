VERSION 5.00
Begin VB.Form dialog_ExportJPEG 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " JPEG Export Options"
   ClientHeight    =   8580
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9240
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
   ScaleHeight     =   572
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   616
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picPreview 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4695
      Left            =   240
      ScaleHeight     =   311
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   583
      TabIndex        =   10
      Top             =   120
      Width           =   8775
   End
   Begin PhotoDemon.commandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   7830
      Width           =   9240
      _ExtentX        =   16298
      _ExtentY        =   1323
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
   Begin PhotoDemon.smartCheckBox chkOptimize 
      Height          =   540
      Left            =   480
      TabIndex        =   5
      Top             =   6240
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
      Left            =   840
      Style           =   2  'Dropdown List
      TabIndex        =   4
      ToolTipText     =   "Subsampling affects the way the JPEG encoder compresses image luminance.  4:2:0 (moderate) is the default value."
      Top             =   7320
      Width           =   4215
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
      TabIndex        =   1
      Top             =   5325
      Width           =   2775
   End
   Begin PhotoDemon.smartCheckBox chkThumbnail 
      Height          =   540
      Left            =   5160
      TabIndex        =   6
      Top             =   6240
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
      Left            =   5160
      TabIndex        =   7
      Top             =   6720
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
      Left            =   480
      TabIndex        =   8
      ToolTipText     =   "Subsampling affects the way the JPEG encoder compresses image luminance.  4:2:0 (moderate) is the default value."
      Top             =   6720
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
      Left            =   3480
      TabIndex        =   9
      Top             =   5265
      Width           =   5655
      _ExtentX        =   9975
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
      Left            =   240
      TabIndex        =   3
      Top             =   5880
      Width           =   1965
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "JPEG quality:"
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
      Left            =   240
      TabIndex        =   2
      Top             =   4920
      Width           =   1410
   End
End
Attribute VB_Name = "dialog_ExportJPEG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'JPEG Export Dialog
'Copyright ©2000-2014 by Tanner Helland
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

'This form can be notified of the image being exported.  This may be used in the future to provide a preview.
Public imageBeingExported As pdImage

'Custom tooltip class allows for things like multiline, theming, and multiple monitor support
Dim m_ToolTip As clsToolTip

'When rendering the preview, we don't want to always re-request a copy of the main image.  Instead, we
' store one in this DIB (at the size of the preview) and simply re-use it when we need to render a preview.
Private origImageCopy As pdDIB
Private previewWidth As Long, previewHeight As Long

'As a further optimizations, we keep a persistent copy of the image in FreeImage format; FreeImage is used to save the
' JPEG in-memory, then render it back out to the picture box.  As the JPEG encoding/decoding is an intensive process,
' anything we can do to alleviate its burden is helpful.
Private fi_DIB As Long

'The user's answer is returned via this property
Public Property Get DialogResult() As VbMsgBoxResult
    DialogResult = userAnswer
End Property

Private Sub chkOptimize_Click()
    updatePreview
End Sub

Private Sub chkProgressive_Click()
    updatePreview
End Sub

Private Sub chkSubsample_Click()
    updatePreview
End Sub

Private Sub chkThumbnail_Click()
    updatePreview
End Sub

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

Private Sub cmbSubsample_Click()
    updatePreview
End Sub

Private Sub cmdBar_CancelClick()
    userAnswer = vbCancel
    Me.Hide
End Sub

Private Sub cmdBar_OKClick()
    
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
    If CBool(chkSubsample) Then g_JPEGFlags = g_JPEGFlags Or getSubsampleConstantFromComboBox()
    
    'Finally, determine whether or not a thumbnail version of the file should be embedded inside
    If CBool(chkThumbnail) Then g_JPEGThumbnail = 1 Else g_JPEGThumbnail = 0
     
    userAnswer = vbOK
    Me.Hide
    
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    updatePreview
End Sub

Private Sub cmdBar_ResetClick()
    
    'Default save quality is "Excellent"
    CmbSaveQuality.ListIndex = 1
    
    'By default, the only advanced setting is Optimize compression tables
    chkOptimize.Value = vbChecked
    chkThumbnail.Value = vbUnchecked
    chkProgressive.Value = vbUnchecked
    chkSubsample.Value = vbUnchecked

End Sub

Private Sub Form_Activate()
    'Draw a preview of the effect
    updatePreview
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
    
    'Release any remaining FreeImage handles
    If fi_DIB <> 0 Then FreeImage_Unload fi_DIB
    If Not origImageCopy Is Nothing Then Set origImageCopy = Nothing
    
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
    
    updatePreview
    
End Sub

'The ShowDialog routine presents the user with this form.
Public Sub showDialog()

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
    
    'Make a copy of the main image; we'll use this to render the preview image
    Set origImageCopy = New pdDIB
    convertAspectRatio imageBeingExported.Width, imageBeingExported.Height, picPreview.Width, picPreview.Height, previewWidth, previewHeight
    origImageCopy.createFromExistingDIB imageBeingExported.getActiveDIB, previewWidth, previewHeight
    If origImageCopy.getDIBColorDepth = 32 Then origImageCopy.convertTo24bpp
    
    'FreeImage is required to perform the JPEG transformation.  We could use GDI+, but FreeImage is
    ' much easier to interface with.
    If g_ImageFormats.FreeImageEnabled Then
    
        'Convert our DIB into FreeImage-format; we will maintain this copy to improve JPEG preview performance.
        fi_DIB = FreeImage_CreateFromDC(origImageCopy.getDIBDC)
        
    'If FreeImage is not available, notify the user.
    Else
        
        Dim tmpDIB As pdDIB
        Set tmpDIB = New pdDIB
        tmpDIB.createBlank picPreview.ScaleWidth, picPreview.ScaleHeight
    
        Dim notifyFont As pdFont
        Set notifyFont = New pdFont
        notifyFont.setFontFace g_InterfaceFont
        notifyFont.setFontSize 14
        notifyFont.setFontColor 0
        notifyFont.setFontBold True
        notifyFont.setTextAlignment vbCenter
        notifyFont.createFontObject
        notifyFont.attachToDC tmpDIB.getDIBDC
    
        notifyFont.fastRenderText tmpDIB.getDIBWidth \ 2, tmpDIB.getDIBHeight \ 2, g_Language.TranslateMessage("JPEG previews require the FreeImage plugin.")
        tmpDIB.renderToPictureBox picPreview
        Set tmpDIB = Nothing
        
    End If
    
    'Assign the system hand cursor to all relevant objects
    Set m_ToolTip = New clsToolTip
    makeFormPretty Me, m_ToolTip
    
    'Update the preview
    updatePreview
    
    'Display the dialog
    showPDDialog vbModal, Me

End Sub

Private Sub sltQuality_Change()
    updateComboBox
End Sub

Private Sub updatePreview()
    If cmdBar.previewsAllowed And g_ImageFormats.FreeImageEnabled And sltQuality.IsValid Then
                
        'Only some of the JPEG settings actually affect the appearance of the saved image.  Specifically, only
        ' save quality and subsampling technique matter.  Convert those into FreeImage-compatible settings now.
        Dim jpegFlags As Long
        jpegFlags = sltQuality.Value
        
        If CBool(chkSubsample) Then jpegFlags = jpegFlags Or getSubsampleConstantFromComboBox()
        
        'Now comes the conversion, which is handled exclusively by FreeImage.  Basically, we ask it to save the image
        ' in JPEG format to a byte array; we then hand that byte array back to it and request a decompression.
        Dim jpegArray() As Byte
        Dim fi_Check As Long
        fi_Check = FreeImage_SaveToMemoryEx(FIF_JPEG, fi_DIB, jpegArray, jpegFlags, False)
        
        Dim tmpFI_DIB As Long
        tmpFI_DIB = FreeImage_LoadFromMemoryEx(jpegArray, FILO_JPEG_FAST)
        
        'Copy the newly decompressed JPEG into our original pdDIB object.
        SetDIBitsToDevice origImageCopy.getDIBDC, 0, 0, origImageCopy.getDIBWidth, origImageCopy.getDIBHeight, 0, 0, 0, origImageCopy.getDIBHeight, ByVal FreeImage_GetBits(tmpFI_DIB), ByVal FreeImage_GetInfo(tmpFI_DIB), 0&
        
        'Paint the final image to screen and release all temporary objects
        origImageCopy.renderToPictureBox picPreview
        FreeImage_Unload tmpFI_DIB
        Erase jpegArray
    
    End If
End Sub

Private Function getSubsampleConstantFromComboBox() As Long
    
    Select Case cmbSubsample.ListIndex
            
        Case 0
            getSubsampleConstantFromComboBox = JPEG_SUBSAMPLING_444
        Case 1
            getSubsampleConstantFromComboBox = JPEG_SUBSAMPLING_422
        Case 2
            getSubsampleConstantFromComboBox = JPEG_SUBSAMPLING_420
        Case 3
            getSubsampleConstantFromComboBox = JPEG_SUBSAMPLING_411
                    
    End Select
    
End Function
