VERSION 5.00
Begin VB.Form dialog_ExportJPEG 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " JPEG Export Options"
   ClientHeight    =   6585
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   13125
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
   ScaleWidth      =   875
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.commandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5835
      Width           =   13125
      _ExtentX        =   23151
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
      Left            =   6000
      TabIndex        =   5
      Top             =   1560
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
      Left            =   6360
      Style           =   2  'Dropdown List
      TabIndex        =   4
      ToolTipText     =   "Subsampling affects the way the JPEG encoder compresses image luminance.  4:2:0 (moderate) is the default value."
      Top             =   3600
      Width           =   6615
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
      Left            =   6000
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   600
      Width           =   2775
   End
   Begin PhotoDemon.smartCheckBox chkThumbnail 
      Height          =   540
      Left            =   6000
      TabIndex        =   6
      Top             =   2040
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
      Left            =   6000
      TabIndex        =   7
      Top             =   2520
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
      Left            =   6000
      TabIndex        =   8
      ToolTipText     =   "Subsampling affects the way the JPEG encoder compresses image luminance.  4:2:0 (moderate) is the default value."
      Top             =   3000
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
      Left            =   8880
      TabIndex        =   9
      Top             =   540
      Width           =   4215
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
   Begin PhotoDemon.fxPreviewCtl fxPreview 
      Height          =   5625
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9922
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
      Left            =   5880
      TabIndex        =   3
      Top             =   1200
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
      Left            =   5880
      TabIndex        =   2
      Top             =   120
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
'Last updated: 15/January/14
'Last update: replaced the standalone picture box preview with PD's dedicated preview control.  This allows the
'              user to pan around the image at their own leisure, and investigate specific elements of the
'              compressed JPEG image.
'
'Dialog for presenting the user various options related to JPEG exporting.  The advanced features all currently
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
    
    'FreeImage is required to perform the JPEG transformation.  We could use GDI+, but FreeImage is
    ' much easier to interface with.  If FreeImage is not available, warn the user.
    If Not g_ImageFormats.FreeImageEnabled Then
        
        Dim tmpDIB As pdDIB
        Set tmpDIB = New pdDIB
        tmpDIB.createBlank fxPreview.getPreviewWidth, fxPreview.getPreviewHeight
    
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
        fxPreview.setOriginalImage tmpDIB
        fxPreview.setFXImage tmpDIB
        'tmpDIB.renderToPictureBox picPreview
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

Private Sub fxPreview_ViewportChanged()
    updatePreview
End Sub

Private Sub sltQuality_Change()
    updateComboBox
End Sub

Private Sub updatePreview()

    If cmdBar.previewsAllowed And g_ImageFormats.FreeImageEnabled And sltQuality.IsValid Then
        
        'Start by retrieving the relevant portion of the image, according to the preview window
        Dim tmpSafeArray As SAFEARRAY2D
        previewNonStandardImage tmpSafeArray, imageBeingExported.getCompositedImage, fxPreview
        
        'workingDIB may be 32bpp at present.  Convert it to 24bpp if necessary.
        If workingDIB.getDIBColorDepth = 32 Then workingDIB.convertTo24bpp
        
        'The public workingDIB object now contains the relevant portion of the preview window.  Pass that to
        ' FreeImage, which will make a copy for itself.
        Dim fi_DIB As Long
        fi_DIB = FreeImage_CreateFromDC(workingDIB.getDIBDC)
                
        'Only some of the JPEG settings actually affect the appearance of the saved image.  Specifically, only
        ' save quality and subsampling technique matter.  Convert those into FreeImage-compatible settings now.
        Dim jpegFlags As Long
        jpegFlags = sltQuality.Value
        
        If CBool(chkSubsample) Then jpegFlags = jpegFlags Or getSubsampleConstantFromComboBox()
        
        'Now comes the conversion, which is handled exclusively by FreeImage.  Basically, we ask it to save the image
        ' in JPEG format to a byte array; we then hand that byte array back to it and request a decompression.
        Dim jpegArray() As Byte
        Dim fi_Check As Long
        fi_Check = FreeImage_SaveToMemoryEx(FIF_JPEG, fi_DIB, jpegArray, jpegFlags, True)
        
        fi_DIB = FreeImage_LoadFromMemoryEx(jpegArray, FILO_JPEG_FAST)
        
        'Copy the newly decompressed JPEG into our original pdDIB object.
        SetDIBitsToDevice workingDIB.getDIBDC, 0, 0, workingDIB.getDIBWidth, workingDIB.getDIBHeight, 0, 0, 0, workingDIB.getDIBHeight, ByVal FreeImage_GetBits(fi_DIB), ByVal FreeImage_GetInfo(fi_DIB), 0&
        
        'Paint the final image to screen and release all temporary objects
        finalizeNonstandardPreview fxPreview
        
        FreeImage_Unload fi_DIB
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
