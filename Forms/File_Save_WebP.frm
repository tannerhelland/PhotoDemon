VERSION 5.00
Begin VB.Form dialog_ExportWebP 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   6585
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12135
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
   ScaleWidth      =   809
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.pdCommandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5835
      Width           =   12135
      _ExtentX        =   21405
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
   Begin PhotoDemon.pdDropDown cboSaveQuality 
      Height          =   735
      Left            =   6120
      TabIndex        =   2
      Top             =   2040
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   1296
      Caption         =   "image compression ratio"
   End
   Begin PhotoDemon.pdSlider sltQuality 
      Height          =   405
      Left            =   6120
      TabIndex        =   3
      Top             =   3000
      Width           =   5775
      _ExtentX        =   15055
      _ExtentY        =   873
      Min             =   1
      Max             =   100
      Value           =   100
      NotchPosition   =   2
      NotchValueCustom=   100
   End
   Begin PhotoDemon.pdLabel lblBefore 
      Height          =   435
      Left            =   6240
      Top             =   3480
      Width           =   2265
      _ExtentX        =   3995
      _ExtentY        =   767
      Caption         =   "low quality, small file"
      FontItalic      =   -1  'True
      FontSize        =   8
      ForeColor       =   4210752
      Layout          =   1
   End
   Begin PhotoDemon.pdLabel lblAfter 
      Height          =   435
      Left            =   8520
      Top             =   3480
      Width           =   2190
      _ExtentX        =   3863
      _ExtentY        =   767
      Alignment       =   1
      Caption         =   "high quality, large file"
      FontItalic      =   -1  'True
      FontSize        =   8
      ForeColor       =   4210752
      Layout          =   1
   End
End
Attribute VB_Name = "dialog_ExportWebP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Google WebP Export Dialog
'Copyright 2014-2021 by Tanner Helland
'Created: 14/February/14
'Last updated: 09/May/16
'Last update: convert dialog to new export engine
'
'Dialog for presenting the user a number of options related to WebP exporting.  Obviously this feature
' relies on FreeImage, and WebP support will be disabled if FreeImage cannot be found.
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

'QUALITY combo box - when adjusted, change the scroll bar to match
Private Sub cboSaveQuality_Click()
    
    Select Case cboSaveQuality.ListIndex
        
        Case 0
            sltQuality = 100
                
        Case 1
            sltQuality = 80
                
        Case 2
            sltQuality = 60
                
        Case 3
            sltQuality = 40
                
        Case 4
            sltQuality = 20
                
    End Select
    
End Sub

Private Sub cmdBar_CancelClick()
    m_UserDialogAnswer = vbCancel
    Me.Visible = False
End Sub

Private Sub cmdBar_OKClick()

    'Determine the compression ratio for the WebP transform
    If (Not sltQuality.IsValid) Then Exit Sub
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.AddParam "webp-quality", sltQuality.Value
    
    m_FormatParamString = cParams.GetParamString
    
    'If ExifTool someday supports WebP metadata embedding, you can add a metadata manager here
    m_MetadataParamString = vbNullString
    
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
    cboSaveQuality.ListIndex = 0
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
        
        Case 100
            If cboSaveQuality.ListIndex <> 0 Then cboSaveQuality.ListIndex = 0
                
        Case 80
            If cboSaveQuality.ListIndex <> 1 Then cboSaveQuality.ListIndex = 1
                
        Case 60
            If cboSaveQuality.ListIndex <> 2 Then cboSaveQuality.ListIndex = 2
                
        Case 40
            If cboSaveQuality.ListIndex <> 3 Then cboSaveQuality.ListIndex = 3
                
        Case 20
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
    
    'Populate the quality drop-down box with presets corresponding to the WebP file format
    cboSaveQuality.Clear
    cboSaveQuality.AddItem "Lossless (100)", 0
    cboSaveQuality.AddItem "Low compression, good image quality (80)", 1
    cboSaveQuality.AddItem "Moderate compression, medium image quality (60)", 2
    cboSaveQuality.AddItem "High compression, poor image quality (40)", 3
    cboSaveQuality.AddItem "Super compression, very poor image quality (20)", 4
    cboSaveQuality.AddItem "Custom ratio (X:1)", 5
    cboSaveQuality.ListIndex = 0
    
    'Make a copy of the composited image; it takes time to composite layers, so we don't want to redo this except
    ' when absolutely necessary.
    Set m_SrcImage = srcImage
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
    Strings.SetFormCaptionW Me, g_Language.TranslateMessage("%1 options", "WebP")
    
    'Display the dialog
    ShowPDDialog vbModal, Me, True
    
End Sub

'When a parameter changes that requires a new source DIB for the preview (e.g. changing the background composite color),
' call this function to generate a new preview DIB.  Note that you *do not* need to call this function for format-specific
' changes (like quality, subsampling, etc).
Private Sub UpdatePreviewSource()
    If (Not m_CompositedImage Is Nothing) Then
        
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
        If sltQuality.IsValid Then fi_Flags = sltQuality.Value Else fi_Flags = 100&
        
        'Retrieve a WebP-saved version of the current preview image
        If Not (workingDIB Is Nothing) Then workingDIB.ResetDIB
        If Plugin_FreeImage.GetExportPreview(m_FIHandle, workingDIB, PDIF_WEBP, fi_Flags) Then
            workingDIB.SetAlphaPremultiplication True, True
            FinalizeNonstandardPreview pdFxPreview, True
        Else
            Debug.Print "WARNING: WEBP EXPORT PREVIEW IS HORRIBLY BROKEN!"
        End If
        
    End If

End Sub
