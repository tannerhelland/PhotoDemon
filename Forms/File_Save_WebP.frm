VERSION 5.00
Begin VB.Form dialog_ExportWebP 
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
   Icon            =   "File_Save_WebP.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   439
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   809
   Begin PhotoDemon.pdDropDown ddImageHint 
      Height          =   855
      Left            =   6000
      TabIndex        =   4
      Top             =   960
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1508
      Caption         =   "image type:"
   End
   Begin PhotoDemon.pdButtonStrip btsCompression 
      Height          =   1095
      Left            =   6000
      TabIndex        =   3
      Top             =   3360
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1931
      Caption         =   "compression"
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
      Left            =   6000
      TabIndex        =   2
      Top             =   1920
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1349
      Caption         =   "quality"
      Min             =   1
      Max             =   100
      Value           =   100
      NotchPosition   =   2
      NotchValueCustom=   100
   End
   Begin PhotoDemon.pdLabel lblInfo 
      Height          =   435
      Index           =   0
      Left            =   6240
      Top             =   2760
      Width           =   2265
      _ExtentX        =   3995
      _ExtentY        =   767
      Caption         =   "low quality, small file"
      FontItalic      =   -1  'True
      FontSize        =   8
      ForeColor       =   4210752
      Layout          =   1
   End
   Begin PhotoDemon.pdLabel lblInfo 
      Height          =   435
      Index           =   1
      Left            =   8520
      Top             =   2760
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
'Copyright 2014-2026 by Tanner Helland
'Created: 14/February/14
'Last updated: 25/September/21
'Last update: overhaul UI to match new approach using libwebp directly (instead of FreeImage)
'
'Dialog for presenting the user a number of options related to WebP exporting.
'
'Google's official WebP library (libwebp) handles actual encoding duties, so this dialog will not
' function without that library being present.  See the pdWebP class for encoding details;
' there's a lot of work involved in encoding WebP files, and that class handles it all for us.
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

'A copy of the untouched preview DIB *BEFORE* saving to the target format.
Private m_ImageBeforeSaving As pdDIB

'pdWebP handles all compression duties (through libwebp)
Private m_WebP As pdWebP

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

Private Sub btsCompression_Click(ByVal buttonIndex As Long)
    UpdatePreview
End Sub

Private Sub cmdBar_CancelClick()
    m_UserDialogAnswer = vbCancel
    Me.Visible = False
End Sub

Private Sub cmdBar_OKClick()

    'Determine the compression ratio for the WebP transform
    If (Not sldQuality.IsValid) Then Exit Sub
    
    m_FormatParamString = GetSaveParameters()
    
    'If ExifTool someday supports WebP metadata embedding, you can add a metadata manager here
    m_MetadataParamString = vbNullString
    
    'Free resources that are no longer required
    Set m_CompositedImage = Nothing
    Set m_SrcImage = Nothing
    
    'Hide but *DO NOT UNLOAD* the form.  The dialog manager needs to retrieve the setting strings before unloading us
    m_UserDialogAnswer = vbOK
    Me.Visible = False
    
End Sub

Private Function GetSaveParameters() As String

    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    With cParams
        
        If sldQuality.IsValid Then .AddParam "webp-quality", sldQuality.Value Else .AddParam "webp-quality", 100
        
        Select Case btsCompression.ListIndex
            Case 0
                .AddParam "webp-compression", "fast"
            Case 1
                .AddParam "webp-compression", "default"
            Case 2
                .AddParam "webp-compression", "slow"
        End Select
        
        Select Case ddImageHint.ListIndex
            Case 0
                .AddParam "webp-hint", "generic"
            Case 1
                .AddParam "webp-hint", "photo-indoor"
            Case 2
                .AddParam "webp-hint", "photo-outdoor"
            Case 3
                .AddParam "webp-hint", "chart"
            Case 4
                .AddParam "webp-hint", "art"
            Case 5
                .AddParam "webp-hint", "icon"
            Case 6
                .AddParam "webp-hint", "text"
        End Select
        
    End With
    
    GetSaveParameters = cParams.GetParamString()
    
End Function

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub cmdBar_ResetClick()
    btsCompression.ListIndex = 1
End Sub

Private Sub ddImageHint_Click()
    UpdatePreview
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
    Set workingDIB = Nothing
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
    
    'Populate any UI elements
    ddImageHint.SetAutomaticRedraws False
    ddImageHint.Clear
    ddImageHint.AddItem "generic", 0, True
    ddImageHint.AddItem "indoor photo", 1
    ddImageHint.AddItem "outdoor photo", 2
    ddImageHint.AddItem "chart", 3
    ddImageHint.AddItem "drawing (or other artwork)", 4
    ddImageHint.AddItem "icon", 5
    ddImageHint.AddItem "text", 6
    ddImageHint.SetAutomaticRedraws True, True
    
    btsCompression.AddItem "fast", 0
    btsCompression.AddItem "balanced", 1
    btsCompression.AddItem "best", 2
    btsCompression.ListIndex = 1
    
    'Make a copy of the composited image; it takes time to composite layers, so we don't want to redo this except
    ' when absolutely necessary.
    Set m_SrcImage = srcImage
    If (m_SrcImage Is Nothing) Then
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
    Interface.SetFormCaptionW Me, g_Language.TranslateMessage("%1 options", "WebP")
    
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
        
        'Make a local copy of workingDIB
        If (m_ImageBeforeSaving Is Nothing) Then Set m_ImageBeforeSaving = New pdDIB
        m_ImageBeforeSaving.CreateFromExistingDIB workingDIB
        
    End If
End Sub

Private Sub UpdatePreview(Optional ByVal forceUpdate As Boolean = False)

    If ((cmdBar.PreviewsAllowed Or forceUpdate) And Plugin_WebP.IsWebPEnabled() And (Not m_SrcImage Is Nothing)) Then
        
        'Make sure the preview source is up-to-date
        If (workingDIB Is Nothing) Then UpdatePreviewSource
        
        'Use pdWebP to preview the compression
        If (m_WebP Is Nothing) Then Set m_WebP = New pdWebP
        If m_WebP.SaveWebP_PreviewOnly(m_ImageBeforeSaving, GetSaveParameters, workingDIB) Then
            
            'Ensure premultiplication (the WebP loader will now provide this by default)
            If (Not workingDIB.GetAlphaPremultiplication) Then workingDIB.SetAlphaPremultiplication True
            FinalizeNonstandardPreview pdFxPreview, True
            
        End If
        
    End If

End Sub
