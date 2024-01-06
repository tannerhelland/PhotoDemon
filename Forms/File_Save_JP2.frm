VERSION 5.00
Begin VB.Form dialog_ExportJP2 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   6585
   ClientLeft      =   45
   ClientTop       =   390
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   439
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   842
   ShowInTaskbar   =   0   'False
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
'Copyright 2012-2024 by Tanner Helland
'Created: 04/December/12
'Last updated: 07/May/16
'Last update: convert dialog to new export engine
'
'Dialog for presenting the user a number of options related to JPEG-2000 exporting.  Obviously this feature
' relies on FreeImage, and JPEG-2000 support will be disabled if FreeImage cannot be found.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'This form can (and should!) be notified of the image being exported.  The only exception to this rule is invoking
' the dialog from the batch process dialog, as no image is associated with that preview.
Private m_SrcImage As pdImage

'A composite of the current image, 32-bpp, fully composited.  This is only regenerated if the source image changes.
Private m_CompositedImage As pdDIB

'FreeImage-specific copy of the preview window corresponding to m_CompositedImage, above.  We cache this to save time,
' but note that it must be regenerated whenever the preview source is regenerated.
Private m_FIHandle As Long

'OK or CANCEL result
Private m_UserDialogAnswer As VbMsgBoxResult

'Final format-specific XML packet, with all format-specific settings defined as tag+value pairs
Private m_FormatParamString As String

'Final metadata XML packet, with all metadata settings defined as tag+value pairs
Private m_MetadataParamString As String

'The user's answer is returned via this property
Public Function GetDialogResult() As VbMsgBoxResult
    GetDialogResult = m_UserDialogAnswer
End Function

Public Function GetFormatParams() As String
    GetFormatParams = m_FormatParamString
End Function

Public Function GetMetadataParams() As String
    GetMetadataParams = m_MetadataParamString
End Function

Private Sub btsCategory_Click(ByVal buttonIndex As Long)
    UpdatePanelVisibility
End Sub

'QUALITY combo box - when adjusted, change the scroll bar to match
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

Private Sub cmdBar_CancelClick()
    m_UserDialogAnswer = vbCancel
    Me.Visible = False
End Sub

Private Sub cmdBar_OKClick()

    'Determine the compression ratio for the JPEG2000 wavelet transformation
    If (Not sltQuality.IsValid) Then Exit Sub
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.AddParam "jp2-quality", Abs(sltQuality)
    
    m_FormatParamString = cParams.GetParamString
    
    'The metadata panel manages its own XML string
    m_MetadataParamString = mtdManager.GetMetadataSettings
    
    'Free resources that are no longer required
    Set m_CompositedImage = Nothing
    Set m_SrcImage = Nothing
    
    'Hide but *DO NOT UNLOAD* the form.  The dialog manager needs to retrieve the setting strings before unloading us
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
    Plugin_FreeImage.ReleasePreviewCache m_FIHandle
End Sub

Private Sub pdFxPreview_ViewportChanged()
    UpdatePreviewSource
    UpdatePreview
End Sub

Private Sub sltQuality_Change()
    UpdateDropDown
    UpdatePreview
End Sub

'Used to keep the "compression ratio" text box, scroll bar, and combo box in sync
Private Sub UpdateDropDown()
    
    Select Case sltQuality.Value
        
        Case 1
            If cboSaveQuality.ListIndex <> 0 Then cboSaveQuality.ListIndex = 0
                
        Case 16
            If cboSaveQuality.ListIndex <> 1 Then cboSaveQuality.ListIndex = 1
                
        Case 32
            If cboSaveQuality.ListIndex <> 2 Then cboSaveQuality.ListIndex = 2
                
        Case 64
            If cboSaveQuality.ListIndex <> 3 Then cboSaveQuality.ListIndex = 3
                
        Case 256
            If cboSaveQuality.ListIndex <> 4 Then cboSaveQuality.ListIndex = 4
                
        Case Else
            If cboSaveQuality.ListIndex <> 5 Then cboSaveQuality.ListIndex = 5
                
    End Select
    
End Sub

'The ShowDialog routine presents the user with this form.
Public Sub ShowDialog(Optional ByRef srcImage As pdImage = Nothing)
    
    'Provide a default answer of "cancel" (in the event that the user clicks the "x" button in the top-right)
    m_UserDialogAnswer = vbCancel
    
    'Make sure that the proper cursor is set
    Screen.MousePointer = 0
    Message "Waiting for user to specify export options... "
    
    'Populate the quality drop-down box with presets corresponding to the JPEG-2000 file format
    cboSaveQuality.Clear
    cboSaveQuality.AddItem "Lossless (1:1)", 0
    cboSaveQuality.AddItem "Low compression, good image quality (16:1)", 1
    cboSaveQuality.AddItem "Moderate compression, medium image quality (32:1)", 2
    cboSaveQuality.AddItem "High compression, poor image quality (64:1)", 3
    cboSaveQuality.AddItem "Super compression, very poor image quality (256:1)", 4
    cboSaveQuality.AddItem "Custom ratio (X:1)", 5
    cboSaveQuality.ListIndex = 0
    
    'Next, prepare various controls on the metadata panel
    Set m_SrcImage = srcImage
    mtdManager.SetParentImage m_SrcImage, PDIF_JPEG
    
    'By default, the basic options panel is always shown.
    btsCategory.AddItem "basic", 0
    btsCategory.AddItem "advanced", 1
    btsCategory.ListIndex = 0
    UpdatePanelVisibility
    
    'Make a copy of the composited image; it takes time to composite layers, so we don't want to redo this except
    ' when absolutely necessary.
    If ((m_SrcImage Is Nothing) Or (Not ImageFormats.IsFreeImageEnabled())) Then
        Interface.ShowDisabledPreviewImage pdFxPreview
    Else
        m_SrcImage.GetCompositedImage m_CompositedImage, True
        pdFxPreview.NotifyNonStandardSource m_CompositedImage.GetDIBWidth, m_CompositedImage.GetDIBHeight
    End If
    
    'Update the preview
    UpdatePreviewSource
    UpdatePreview True
    
    'Apply translations and visual themes
    ApplyThemeAndTranslations Me
    Interface.SetFormCaptionW Me, g_Language.TranslateMessage("%1 options", "JPEG-2000")
    
    'Display the dialog
    ShowPDDialog vbModal, Me, True
    
End Sub

'When a parameter changes that requires a new source DIB for the preview (e.g. changing the background composite color),
' call this function to generate a new preview DIB.  Note that you *do not* need to call this function for format-specific
' changes (like quality, subsampling, etc).
Private Sub UpdatePreviewSource()

    If (Not (m_CompositedImage Is Nothing)) Then
        
        'Because the user can change the preview viewport, we can't guarantee that the preview region hasn't changed
        ' since the last preview.  Prep a new preview now.
        Dim tmpSafeArray As SafeArray2D
        EffectPrep.PreviewNonStandardImage tmpSafeArray, m_CompositedImage, pdFxPreview, False
        
        'Finally, convert that preview copy to a FreeImage-compatible handle.
        If (m_FIHandle <> 0) Then Plugin_FreeImage.ReleaseFreeImageObject m_FIHandle
        
        'During previews, we can always use 32-bpp mode
        m_FIHandle = Plugin_FreeImage.GetFIDib_SpecificColorMode(workingDIB, 32, PDAS_ComplicatedAlpha)
        
    End If
    
End Sub

Private Sub UpdatePreview(Optional ByVal forceUpdate As Boolean = False)

    If ((cmdBar.PreviewsAllowed Or forceUpdate) And ImageFormats.IsFreeImageEnabled() And (Not m_SrcImage Is Nothing)) Then
        
        'Make sure the preview source is up-to-date
        If (m_FIHandle = 0) Then UpdatePreviewSource
        
        'Prep all relevant FreeImage flags
        Dim fi_Flags As FREE_IMAGE_SAVE_OPTIONS
        If sltQuality.IsValid Then fi_Flags = Abs(sltQuality.Value) Else fi_Flags = 0&
        
        'Retrieve a JPEG-saved version of the current preview image
        workingDIB.ResetDIB
        If Plugin_FreeImage.GetExportPreview(m_FIHandle, workingDIB, PDIF_JP2, fi_Flags) Then
            workingDIB.SetAlphaPremultiplication True, True
            FinalizeNonstandardPreview pdFxPreview, True
        Else
            Debug.Print "WARNING: JP2 EXPORT PREVIEW IS HORRIBLY BROKEN!"
        End If
        
    End If

End Sub

Private Sub UpdatePanelVisibility()
    Dim i As Long
    For i = 0 To btsCategory.ListCount - 1
        picContainer(i).Visible = (i = btsCategory.ListIndex)
    Next i
End Sub
