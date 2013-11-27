VERSION 5.00
Begin VB.Form dialog_ExportJP2 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " JPEG 2000 Export Options"
   ClientHeight    =   7815
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9255
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
   ScaleHeight     =   521
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   617
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
      TabIndex        =   5
      Top             =   120
      Width           =   8775
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
      TabIndex        =   0
      Top             =   5430
      Width           =   8295
   End
   Begin PhotoDemon.sliderTextCombo sltQuality 
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   6000
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   873
      Min             =   1
      Max             =   256
      Value           =   16
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
   Begin PhotoDemon.commandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   6
      Top             =   7065
      Width           =   9255
      _ExtentX        =   16325
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
      Left            =   480
      TabIndex        =   3
      Top             =   6480
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
      Left            =   6240
      TabIndex        =   2
      Top             =   6480
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
      Left            =   240
      TabIndex        =   1
      Top             =   5040
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
'Copyright ©2012-2013 by Tanner Helland
'Created: 04/December/12
'Last updated: 22/November/13
'Last update: added live previews!
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

'Custom tooltip class allows for things like multiline, theming, and multiple monitor support
Dim m_ToolTip As clsToolTip

'When rendering the preview, we don't want to always re-request a copy of the main image.  Instead, we
' store one in this layer (at the size of the preview) and simply re-use it when we need to render a preview.
Private origImageCopy As pdLayer
Private previewWidth As Long, previewHeight As Long

'As a further optimizations, we keep a persistent copy of the image in FreeImage format; FreeImage is used to save the
' JP2 in-memory, then render it back out to the picture box.  As JP2 encoding/decoding is an intensive process,
' anything we can do to alleviate its burden is helpful.
Private fi_DIB As Long

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
    
    'Release any remaining FreeImage handles
    If fi_DIB <> 0 Then FreeImage_Unload fi_DIB
    If Not origImageCopy Is Nothing Then Set origImageCopy = Nothing
    
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
        
    'Make a copy of the main image; we'll use this to render the preview image
    Set origImageCopy = New pdLayer
    convertAspectRatio imageBeingExported.Width, imageBeingExported.Height, picPreview.Width, picPreview.Height, previewWidth, previewHeight
    origImageCopy.createFromExistingLayer imageBeingExported.getActiveLayer, previewWidth, previewHeight
    If origImageCopy.getLayerColorDepth = 32 Then origImageCopy.convertTo24bpp
    
    'FreeImage is required to perform the live JPEG-2000 transformation.
    If g_ImageFormats.FreeImageEnabled Then
    
        'Convert our DIB into FreeImage-format; we will maintain this copy to improve JPEG preview performance.
        fi_DIB = FreeImage_CreateFromDC(origImageCopy.getLayerDC)
        
    'If FreeImage is not available, notify the user.  (It should not be possible to trigger this dialog without
    ' FreeImage being present, but it doesn't hurt to provide this fallback!)
    Else
        
        Dim tmpLayer As pdLayer
        Set tmpLayer = New pdLayer
        tmpLayer.createBlank picPreview.ScaleWidth, picPreview.ScaleHeight
    
        Dim notifyFont As pdFont
        Set notifyFont = New pdFont
        notifyFont.setFontFace g_InterfaceFont
        notifyFont.setFontSize 14
        notifyFont.setFontColor 0
        notifyFont.setFontBold True
        notifyFont.setTextAlignment vbCenter
        notifyFont.createFontObject
        notifyFont.attachToDC tmpLayer.getLayerDC
    
        notifyFont.fastRenderText tmpLayer.getLayerWidth \ 2, tmpLayer.getLayerHeight \ 2, g_Language.TranslateMessage("Live previews require the FreeImage plugin.")
        tmpLayer.renderToPictureBox picPreview
        Set tmpLayer = Nothing
        
    End If
        
    'Assign the system hand cursor to all relevant objects
    Set m_ToolTip = New clsToolTip
    makeFormPretty Me, m_ToolTip
    
    'Update the preview
    updatePreview
    
    'Display the dialog
    showPDDialog vbModal, Me

End Sub

'Render a new JPEG-2000 preview
Private Sub updatePreview()

    If cmdBar.previewsAllowed And g_ImageFormats.FreeImageEnabled And sltQuality.IsValid Then
        
        'Perform a live, in-memory conversion to JP2 using FreeImage.  Basically, we ask it to save the image
        ' in JP2 format to a byte array; we then hand that byte array back to it and request a decompression.
        Dim jp2Array() As Byte
        Dim fi_Check As Long
        fi_Check = FreeImage_SaveToMemoryEx(FIF_JP2, fi_DIB, jp2Array, Abs(sltQuality.Value), False)
        
        Dim tmpFI_DIB As Long
        tmpFI_DIB = FreeImage_LoadFromMemoryEx(jp2Array, 0)
        
        'Copy the newly decompressed JPEG-2000 into our original pdLayer object.
        SetDIBitsToDevice origImageCopy.getLayerDC, 0, 0, origImageCopy.getLayerWidth, origImageCopy.getLayerHeight, 0, 0, 0, origImageCopy.getLayerHeight, ByVal FreeImage_GetBits(tmpFI_DIB), ByVal FreeImage_GetInfo(tmpFI_DIB), 0&
        
        'Paint the final image to screen and release all temporary objects
        origImageCopy.renderToPictureBox picPreview
        FreeImage_Unload tmpFI_DIB
        Erase jp2Array
    
    End If

End Sub
