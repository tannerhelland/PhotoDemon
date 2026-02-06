VERSION 5.00
Begin VB.Form dialog_ExportAVIF 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   6585
   ClientLeft      =   120
   ClientTop       =   465
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
   Icon            =   "File_Save_AVIF.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   439
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   809
   Begin PhotoDemon.pdCheckBox chkLivePreview 
      Height          =   375
      Left            =   6120
      TabIndex        =   3
      Top             =   3600
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   661
      Caption         =   "preview quality changes"
      FontSize        =   11
      Value           =   0   'False
   End
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
   Begin PhotoDemon.pdSlider sldQuality 
      Height          =   765
      Left            =   6120
      TabIndex        =   2
      Top             =   2040
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   1349
      Caption         =   "image quality"
      Max             =   63
      Value           =   63
      NotchPosition   =   2
      NotchValueCustom=   100
      DefaultValue    =   63
   End
   Begin PhotoDemon.pdLabel lblBefore 
      Height          =   435
      Left            =   6240
      Top             =   2880
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
      Top             =   2880
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
Attribute VB_Name = "dialog_ExportAVIF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'AVIF (AV1) Export Dialog
'Copyright 2021-2026 by Tanner Helland
'Created: 28/July/21
'Last updated: 03/August/21
'Last update: wrap up initial build
'
'Dialog for presenting the user a number of options related to AVIF exporting.  Obviously this feature
' relies on a 3rd-party library for operation (currently libavif); this export dialog is not accessible
' if the required 3rd-party library isn't available.
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

'Clean, modified source preview image, in PNG format.  (This only needs to be created when the preview source
' image changes - e.g. if the user zooms or scrolls the preview control.)
Private m_PreviewImagePath As String, m_PreviewImageBackup As pdDIB

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

Private Sub chkLivePreview_Click()
    UpdatePreview
End Sub

Private Sub cmdBar_CancelClick()
    m_UserDialogAnswer = vbCancel
    Me.Visible = False
End Sub

Private Sub cmdBar_OKClick()

    'Ensure valid quality settings
    If (Not sldQuality.IsValid) Then Exit Sub
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.AddParam "avif-quality", 63 - sldQuality.Value
    
    m_FormatParamString = cParams.GetParamString
    
    'If ExifTool someday supports WebP metadata embedding, you can add a metadata manager here
    m_MetadataParamString = vbNullString
    
    'Free resources that are no longer required
    Set m_CompositedImage = Nothing
    Set m_SrcImage = Nothing
    If (LenB(m_PreviewImagePath) > 0) Then Files.FileDeleteIfExists m_PreviewImagePath
    
    'Hide but *DO NOT UNLOAD* the form.  The dialog manager needs to retrieve the setting strings before unloading us
    m_UserDialogAnswer = vbOK
    Me.Visible = False
    
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub Form_Load()
    chkLivePreview.AssignTooltip "This image format is very computationally intensive.  On older or slower PCs, you may want to disable live previews."
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    'Ensure we release any temp files on exit
    Files.FileDeleteIfExists m_PreviewImagePath
    
    'Release subclassing form themer
    ReleaseFormTheming Me
    
End Sub

Private Sub pdFxPreview_ViewportChanged()
    UpdatePreviewSource
    UpdatePreview
End Sub

Private Sub sldQuality_Change()
    UpdatePreview
End Sub

'The ShowDialog routine presents the user with this form.
Public Sub ShowDialog(Optional ByRef srcImage As pdImage = Nothing)

    'Provide a default answer of "cancel" (in the event that the user clicks the "x" button in the top-right)
    m_UserDialogAnswer = vbCancel
    
    'Make sure that the proper cursor is set
    Screen.MousePointer = 0
    Message "Waiting for user to specify export options... "
    
    'Make a copy of the composited image; it takes time to composite layers, so we don't want to redo this except
    ' when absolutely necessary.
    Set m_SrcImage = srcImage
    If ((m_SrcImage Is Nothing) Or (Not Plugin_AVIF.IsAVIFExportAvailable())) Then
        Interface.ShowDisabledPreviewImage pdFxPreview
    Else
        m_SrcImage.GetCompositedImage m_CompositedImage, True
        pdFxPreview.NotifyNonStandardSource m_CompositedImage.GetDIBWidth, m_CompositedImage.GetDIBHeight
    End If
    
    'Update the preview
    UpdatePreviewSource
    UpdatePreview True
    
    'Apply translations and visual themes
    ApplyThemeAndTranslations Me, True, True
    Interface.SetFormCaptionW Me, g_Language.TranslateMessage("%1 options", "AVIF")
    If (Not g_WindowManager Is Nothing) Then g_WindowManager.SetFocusAPI cmdBar.hWnd
    
    'Display the dialog
    ShowPDDialog vbModal, Me, True
    
End Sub

'When a parameter changes that requires a new source DIB for the preview (e.g. changing the background composite color),
' call this function to generate a new preview DIB.  Note that you *do not* need to call this function for format-specific
' changes (like quality, subsampling, etc).
Private Sub UpdatePreviewSource()

    If Not (m_CompositedImage Is Nothing) Then
        
        'Because the user can change the preview viewport, we can't guarantee that the preview region hasn't changed
        ' since the last preview.  Prep a new preview now.
        Dim tmpSafeArray As SafeArray2D
        EffectPrep.PreviewNonStandardImage tmpSafeArray, m_CompositedImage, pdFxPreview, False
        
        If (m_PreviewImageBackup Is Nothing) Then Set m_PreviewImageBackup = New pdDIB
        m_PreviewImageBackup.CreateFromExistingDIB workingDIB
        
        'Save a copy of the source image to file, in PNG format.  (PD's current AVIF encoder
        ' works as a command-line tool; we need to pass it a source PNG file.)
        If (LenB(m_PreviewImagePath) > 0) Then Files.FileDeleteIfExists m_PreviewImagePath
        m_PreviewImagePath = OS.UniqueTempFilename(customExtension:="png")
        If (Not Saving.QuickSaveDIBAsPNG(m_PreviewImagePath, workingDIB, False, True)) Then
            InternalError "UpdatePreviewSource", "couldn't save preview png"
        End If
        
    End If
    
End Sub

Private Sub UpdatePreview(Optional ByVal forceUpdate As Boolean = False)
    
    Const funcName As String = "UpdatePreview"
    
    If ((cmdBar.PreviewsAllowed Or forceUpdate) And Plugin_AVIF.IsAVIFExportAvailable() And (Not m_SrcImage Is Nothing)) Then
        
        'Make sure the preview source is up-to-date
        If (workingDIB Is Nothing) Then UpdatePreviewSource
        If (workingDIB Is Nothing) Then Exit Sub
        
        'Because AVIF previews are so intensive to generate, this dialog provides a toggle so the user
        ' can suspend real-time previews.
        If chkLivePreview.Value Then
            
            'Now perform the (ugly) dance of workingDIB > PNG > AVIF > PNG > workingDIB.
            ' (Note that the first workingDIB > PNG step was performed by UpdatePreviewSource.)
            
            'Start by generating temporary filenames for intermediary files
            Dim tmpFilenameBase As String, tmpFilenameIntermediary As String, tmpFilenameAVIF As String
            tmpFilenameBase = OS.UniqueTempFilename()
            tmpFilenameIntermediary = tmpFilenameBase & ".png"
            tmpFilenameAVIF = tmpFilenameBase & ".avif"
            
            'Shell libavif, and request it to convert the preview PNG to AVIF
            If Plugin_AVIF.ConvertStandardImageToAVIF(m_PreviewImagePath, tmpFilenameAVIF, 63 - sldQuality.Value, 10) Then
            
                'Immediately shell it again, but this time, ask it to convert the AVIF it just made
                ' back into a PNG
                Files.FileDeleteIfExists tmpFilenameIntermediary
                If Plugin_AVIF.ConvertAVIFtoStandardImage(tmpFilenameAVIF, tmpFilenameIntermediary, False) Then
                    
                    'We are done with the AVIF; kill it
                    Files.FileDeleteIfExists tmpFilenameAVIF
                    
                    'Load the finished PNG *back* into a pdDIB object
                    If Loading.QuickLoadImageToDIB(tmpFilenameIntermediary, workingDIB, False, False, True) Then
                        
                        'We are done with the intermediary image; kill it
                        Files.FileDeleteIfExists tmpFilenameIntermediary
                        
                        'Display the final result
                        workingDIB.SetAlphaPremultiplication True, True
                        FinalizeNonstandardPreview pdFxPreview, True
                        
                    Else
                        InternalError funcName, "couldn't load finished PNG to pdDIB"
                    End If
                
                Else
                    InternalError funcName, "couldn't convert AVIF back to PNG"
                End If
            
            Else
                InternalError funcName, "couldn't save AVIF"
            End If
        
        'Live previews are disabled; just mirror the original image to the screen
        Else
            workingDIB.CreateFromExistingDIB m_PreviewImageBackup
            FinalizeNonstandardPreview pdFxPreview, False
        End If
                
    Else
        InternalError funcName, "avif library broken"
    End If

End Sub

Private Sub InternalError(ByRef funcName As String, ByRef errMsg As String)
    PDDebug.LogAction "WARNING! Problem in dialog_ExportAVIF." & funcName & ": " & errMsg
End Sub
