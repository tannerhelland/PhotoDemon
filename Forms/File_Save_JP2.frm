VERSION 5.00
Begin VB.Form dialog_ExportJP2 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   6585
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12630
   DrawStyle       =   5  'Transparent
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HasDC           =   0   'False
   Icon            =   "File_Save_JP2.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   439
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   842
   Begin PhotoDemon.pdCommandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5835
      Width           =   12630
      _ExtentX        =   22278
      _ExtentY        =   1323
      DontAutoUnloadParent=   -1  'True
   End
   Begin PhotoDemon.pdFxPreviewCtl pdFxPreview 
      Height          =   5625
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9922
   End
   Begin PhotoDemon.pdButtonStrip btsCategory 
      Height          =   615
      Left            =   5880
      TabIndex        =   2
      Top             =   120
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   1085
      FontSize        =   11
   End
   Begin PhotoDemon.pdContainer picContainer 
      Height          =   4815
      Index           =   0
      Left            =   5880
      Top             =   840
      Width           =   6615
      _ExtentX        =   0
      _ExtentY        =   0
      Begin PhotoDemon.pdDropDown cboSaveQuality 
         Height          =   735
         Left            =   120
         TabIndex        =   3
         Top             =   1320
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   1296
         Caption         =   "image compression ratio"
      End
      Begin PhotoDemon.pdSlider sltQuality 
         Height          =   405
         Left            =   120
         TabIndex        =   5
         Top             =   2280
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   714
         Min             =   1
         Max             =   256
         Value           =   16
         NotchPosition   =   1
         DefaultValue    =   16
      End
      Begin PhotoDemon.pdLabel lblBefore 
         Height          =   435
         Left            =   240
         Top             =   2760
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   767
         Caption         =   "high quality, large file"
         FontItalic      =   -1  'True
         FontSize        =   8
         ForeColor       =   4210752
         Layout          =   1
      End
      Begin PhotoDemon.pdLabel lblAfter 
         Height          =   435
         Left            =   3240
         Top             =   2760
         Width           =   2190
         _ExtentX        =   3863
         _ExtentY        =   767
         Alignment       =   1
         Caption         =   "low quality, small file"
         FontItalic      =   -1  'True
         FontSize        =   8
         ForeColor       =   4210752
         Layout          =   1
      End
   End
   Begin PhotoDemon.pdContainer picContainer 
      Height          =   4815
      Index           =   1
      Left            =   5880
      Top             =   840
      Width           =   6615
      _ExtentX        =   0
      _ExtentY        =   0
      Begin PhotoDemon.pdMetadataExport mtdManager 
         Height          =   4215
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   7435
      End
   End
End
Attribute VB_Name = "dialog_ExportJP2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'JPEG-2000 (JP2) Export Dialog
'Copyright 2012-2026 by Tanner Helland
'Created: 04/December/12
'Last updated: 11/November/25
'Last update: overhaul to use OpenJPEG instead of FreeImage
'
'This dialog provides the UI for JPEG-2000 exporting.  As of 2025, this dialog requires OpenJPEG for both
' previewing JP2 export settings, and handling the ultimate export to file.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'This form needs to be notified of the image being exported.
' The only time to *not* do this is when invoking this dialog from the batch process tool
' (because a specific image won't be associated with the preview).
Private m_SrcImage As pdImage

'A composite of the current image, 32-bpp, fully composited.
' This only needs to be regenerated when the source image changes.
Private m_CompositedImage As pdDIB

'Preview stream (holds JP2-compressed bytes, cached to improve perf on low-end PCs)
Private m_previewStream As pdStream

'Preview DIB (holds the actual preview image, cached to improve perf)
Private m_PreviewDIB As pdDIB

'OK or CANCEL result.  Must be returned to the caller.
Private m_UserDialogAnswer As VbMsgBoxResult

'Final format-specific XML packet, with all format-specific settings defined as tag+value pairs
Private m_FormatParamString As String

'Final metadata-specific XML packet, with all metadata settings defined as tag+value pairs
Private m_MetadataParamString As String

'For this format, output color depth is auto-calculated based on image contents.
' The caller cannot request specific color modes.
Private m_outputColorDepth As Long

'The user's answer is returned via the following properties
Public Function GetDialogResult() As VbMsgBoxResult
    GetDialogResult = m_UserDialogAnswer
End Function

Public Function GetFormatParams() As String
    GetFormatParams = m_FormatParamString
End Function

Public Function GetMetadataParams() As String
    GetMetadataParams = m_MetadataParamString
End Function

'Switchin between format-specific and metadata-specific settings
Private Sub btsCategory_Click(ByVal buttonIndex As Long)
    UpdatePanelVisibility
End Sub

'QUALITY dropdown must auto-synchronize with the scroll bar
Private Sub cboSaveQuality_Click()
    
    Select Case cboSaveQuality.ListIndex
        
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

'When closing the dialog (via OK or CANCEL), some caches must be manually freed
Private Sub cmdBar_CancelClick()
    m_UserDialogAnswer = vbCancel
    Plugin_OpenJPEG.FreeJp2Caches
    Me.Visible = False
End Sub

Private Sub cmdBar_OKClick()
    
    'Highlight errors and return to the dialog, preventing exit via OK until problems are rectified
    If (Not sltQuality.IsValid) Then Exit Sub
    
    'Serialize all parameters to string (currently just quality at present)
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.AddParam "jp2-quality", Abs(sltQuality)
    
    m_FormatParamString = cParams.GetParamString
    
    'The metadata panel manages its own XML string
    m_MetadataParamString = mtdManager.GetMetadataSettings
    
    'Free JP2-specific resources that are no longer required
    Set m_CompositedImage = Nothing
    Set m_SrcImage = Nothing
    Plugin_OpenJPEG.FreeJp2Caches
    
    'Hide but *DO NOT UNLOAD* the form.  The dialog manager needs to retrieve all setting strings before unloading us.
    m_UserDialogAnswer = vbOK
    Me.Visible = False
    
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub cmdBar_ResetClick()
    mtdManager.Reset
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
    Plugin_OpenJPEG.FreeJp2Caches   'Failsafe cache free (should have already happened via OK/Cancel)
End Sub

Private Sub pdFxPreview_ViewportChanged()
    UpdatePreviewSource
    UpdatePreview
End Sub

Private Sub sltQuality_Change()
    UpdateDropDown
    UpdatePreview
End Sub

'Keep the "compression" text box, scroll bar, and combo box in sync
Private Sub UpdateDropDown()
    
    Select Case sltQuality.Value
        
        Case 1
            If (cboSaveQuality.ListIndex <> 0) Then cboSaveQuality.ListIndex = 0
                
        Case 16
            If (cboSaveQuality.ListIndex <> 1) Then cboSaveQuality.ListIndex = 1
                
        Case 32
            If (cboSaveQuality.ListIndex <> 2) Then cboSaveQuality.ListIndex = 2
                
        Case 64
            If (cboSaveQuality.ListIndex <> 3) Then cboSaveQuality.ListIndex = 3
                
        Case 256
            If (cboSaveQuality.ListIndex <> 4) Then cboSaveQuality.ListIndex = 4
                
        Case Else
            If (cboSaveQuality.ListIndex <> 5) Then cboSaveQuality.ListIndex = 5
                
    End Select
    
End Sub

'The ShowDialog routine presents this form to the user.
Public Sub ShowDialog(Optional ByRef srcImage As pdImage = Nothing)
    
    'Provide a default answer of "cancel" (in the event that the user clicks the "x" button in the top-right)
    m_UserDialogAnswer = vbCancel
    
    'Make sure that the proper cursor is set
    Screen.MousePointer = 0
    Message "Waiting for user to specify export options... "
    
    'Populate the quality drop-down box with JP2-specific presets.
    ' (There's no "good" way to present this, because JPEG-2000 doesn't expose quality the same way
    '  other formats do.  This strategy was the first draft I attempted in PD decades ago, and because
    '  translations are done I've stuck with it ever since.  [Insert shrug emoji])
    cboSaveQuality.Clear
    cboSaveQuality.AddItem g_Language.TranslateMessage("Lossless (%1)", "1:1"), 0
    cboSaveQuality.AddItem g_Language.TranslateMessage("Low compression, good image quality (%1)", "16:1"), 1
    cboSaveQuality.AddItem g_Language.TranslateMessage("Moderate compression, medium image quality (%1)", "32:1"), 2
    cboSaveQuality.AddItem g_Language.TranslateMessage("High compression, poor image quality (%1)", "64:1"), 3
    cboSaveQuality.AddItem g_Language.TranslateMessage("Super compression, very poor image quality (%1)", "256:1"), 4
    cboSaveQuality.AddItem g_Language.TranslateMessage("Custom ratio (%1)", "X:1"), 5
    cboSaveQuality.ListIndex = 0
    
    'Next, prepare various controls on the metadata panel
    Set m_SrcImage = srcImage
    mtdManager.SetParentImage m_SrcImage, PDIF_JP2
    
    'By default, the basic (format-specific) options panel is always shown.
    btsCategory.AddItem "basic", 0
    btsCategory.AddItem "advanced", 1
    btsCategory.ListIndex = 0
    UpdatePanelVisibility
    
    'Make a copy of the composited image; it takes time to composite layers, so we only want to do this once
    If ((m_SrcImage Is Nothing) Or (Not ImageFormats.IsFreeImageEnabled())) Then
        Interface.ShowDisabledPreviewImage pdFxPreview
    Else
        m_SrcImage.GetCompositedImage m_CompositedImage, True
        pdFxPreview.NotifyNonStandardSource m_CompositedImage.GetDIBWidth, m_CompositedImage.GetDIBHeight
    End If
    
    'Draw the initial preview
    UpdatePreviewSource
    UpdatePreview True
    
    'Apply translations and visual themes
    ApplyThemeAndTranslations Me, True, True
    Interface.SetFormCaptionW Me, g_Language.TranslateMessage("%1 options", "JPEG-2000")
    
    'Present the dialog
    ShowPDDialog vbModal, Me, True
    
End Sub

'When a parameter changes that requires a new source image for the preview (e.g. changing the background composite color),
' call this function to generate a new preview image.  Note that you *do not* need to call this function for format-specific
' changes (like quality, subsampling, etc).
Private Sub UpdatePreviewSource()

    If (Not m_CompositedImage Is Nothing) Then
        
        'Because the user can change the preview viewport, we can't guarantee that the preview region hasn't changed
        ' since the last preview.  Prep a new preview now.
        Dim tmpSafeArray As SafeArray2D
        EffectPrep.PreviewNonStandardImage tmpSafeArray, m_CompositedImage, pdFxPreview, False
        
        'The public object "workingDIB" now contains the source preview.  Make a backup copy
        ' of that image to a local DIB object; we'll reuse that on subsequent calls.
        If (m_PreviewDIB Is Nothing) Then Set m_PreviewDIB = New pdDIB
        m_PreviewDIB.CreateFromExistingDIB workingDIB
        
        'Auto-determine output color depth now
        If DIBs.IsDIBTransparent(m_PreviewDIB) Then
            m_outputColorDepth = 32
        Else
            If DIBs.IsDIBGrayscale(m_PreviewDIB) Then
                m_outputColorDepth = 8
            Else
                m_outputColorDepth = 24
            End If
        End If
        
        'Notify the OpenJPEG wrapper to clear any internal caches because the source image has changed
        Plugin_OpenJPEG.FreeJp2Caches
        
    End If
    
End Sub

'Draw a new preview.  Only refreshes JP2-specific settings - image-specific settings must be set prior,
' via UpdatePreviewSource(), above.
Private Sub UpdatePreview(Optional ByVal forceUpdate As Boolean = False)
    
    'Prevent redraws during dialog initialization
    If (Not Me.Visible) Or (m_outputColorDepth = 0) Then Exit Sub
    
    'Previews need to be disabled during batch processes, missing plugins, etc
    If ((cmdBar.PreviewsAllowed Or forceUpdate) And Plugin_OpenJPEG.IsOpenJPEGEnabled() And (Not m_SrcImage Is Nothing)) Then
        
        'Failsafe only
        If (m_PreviewDIB Is Nothing) Then UpdatePreviewSource
        
        'Set quality flags
        Dim saveQuality As Long
        If sltQuality.IsValid Then saveQuality = Abs(sltQuality.Value) Else saveQuality = 0&
        
        'Initialize a pdStream object (memory only for the preview)
        If (m_previewStream Is Nothing) Then
            Set m_previewStream = New pdStream
            m_previewStream.StartStream PD_SM_MemoryBacked, PD_SA_ReadWrite, startingBufferSize:=1048576
        Else
            m_previewStream.StopStream False
            m_previewStream.StartStream PD_SM_MemoryBacked, PD_SA_ReadWrite, reuseExistingBuffer:=True
        End If
        
        'Perform a fast in-memory save to the target stream
        If Plugin_OpenJPEG.SavePdDIBToJp2Stream(m_PreviewDIB, m_previewStream, saveQuality, m_outputColorDepth, forceNewImageObject:=True) Then
            
            'The preview stream now contains the encoded JP2 bytes.
            ' Reset the stream pointer to the start of the stream.
            m_previewStream.SetPosition 0&, FILE_BEGIN
            
            'Now we want to decode the JP2 bytes back into a standard RGBA buffer
            If Plugin_OpenJPEG.FastDecodeFromStreamToDIB(m_previewStream, workingDIB) Then
            
                'Flip the resulting preview to the screen
                workingDIB.SetAlphaPremultiplication True, True
                EffectPrep.FinalizeNonstandardPreview pdFxPreview, True
                
            Else
                Debug.Print "Failed to fast decode jp2 stream"
            End If
            
        Else
            Debug.Print "WARNING: JP2 EXPORT PREVIEW IS HORRIBLY BROKEN!"
        End If
        
    End If

End Sub

'Flip between export settings panels
Private Sub UpdatePanelVisibility()
    Dim i As Long
    For i = 0 To btsCategory.ListCount - 1
        picContainer(i).Visible = (i = btsCategory.ListIndex)
    Next i
End Sub
