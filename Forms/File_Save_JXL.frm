VERSION 5.00
Begin VB.Form dialog_ExportJXL 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13110
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
   ScaleHeight     =   436
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   874
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.pdButtonStrip btsCategory 
      Height          =   615
      Left            =   6000
      TabIndex        =   1
      Top             =   120
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   1085
      FontSize        =   11
   End
   Begin PhotoDemon.pdCommandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5790
      Width           =   13110
      _ExtentX        =   23125
      _ExtentY        =   1323
      DontAutoUnloadParent=   -1  'True
   End
   Begin PhotoDemon.pdFxPreviewCtl pdFxPreview 
      Height          =   5625
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9922
   End
   Begin PhotoDemon.pdContainer picContainer 
      Height          =   4695
      Index           =   0
      Left            =   5880
      Top             =   1080
      Width           =   7215
      _ExtentX        =   0
      _ExtentY        =   0
      Begin PhotoDemon.pdButtonStrip btsQuality 
         Height          =   1095
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   1931
         Caption         =   "quality"
      End
      Begin PhotoDemon.pdSlider sldEffort 
         Height          =   975
         Left            =   120
         TabIndex        =   6
         Top             =   2280
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   1720
         Caption         =   "compression effort"
         Min             =   1
         Max             =   9
         Value           =   7
         NotchPosition   =   2
         NotchValueCustom=   7
      End
      Begin PhotoDemon.pdSlider sldQuality 
         Height          =   405
         Left            =   120
         TabIndex        =   3
         Top             =   1320
         Width           =   6975
         _ExtentX        =   7223
         _ExtentY        =   873
         Min             =   1
         Max             =   15
         SigDigits       =   2
         Value           =   1
         NotchPosition   =   1
         DefaultValue    =   1
      End
      Begin PhotoDemon.pdLabel lblHint 
         Height          =   255
         Index           =   0
         Left            =   480
         Top             =   3240
         Width           =   2340
         _ExtentX        =   4128
         _ExtentY        =   450
         Caption         =   "fast, larger file"
         FontItalic      =   -1  'True
         FontSize        =   9
      End
      Begin PhotoDemon.pdLabel lblHint 
         Height          =   255
         Index           =   1
         Left            =   2880
         Top             =   3240
         Width           =   2580
         _ExtentX        =   4551
         _ExtentY        =   450
         Alignment       =   1
         Caption         =   "slow, smaller file"
         FontItalic      =   -1  'True
         FontSize        =   9
      End
      Begin PhotoDemon.pdLabel lblHint 
         Height          =   255
         Index           =   3
         Left            =   2880
         Top             =   1800
         Width           =   2580
         _ExtentX        =   4551
         _ExtentY        =   450
         Alignment       =   1
         Caption         =   "low quality, small file"
         FontItalic      =   -1  'True
         FontSize        =   9
         ForeColor       =   4210752
         Layout          =   1
      End
      Begin PhotoDemon.pdLabel lblHint 
         Height          =   255
         Index           =   2
         Left            =   480
         Top             =   1800
         Width           =   2340
         _ExtentX        =   4128
         _ExtentY        =   450
         Caption         =   "high quality, large file"
         FontItalic      =   -1  'True
         FontSize        =   9
         ForeColor       =   4210752
         Layout          =   1
      End
   End
   Begin PhotoDemon.pdContainer picContainer 
      Height          =   4695
      Index           =   1
      Left            =   5880
      Top             =   1080
      Width           =   7215
      _ExtentX        =   0
      _ExtentY        =   0
      Begin PhotoDemon.pdButtonStrip btsDepth 
         Height          =   1095
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   1931
         Caption         =   "depth"
      End
   End
   Begin PhotoDemon.pdContainer picContainer 
      Height          =   4695
      Index           =   2
      Left            =   5880
      Top             =   1080
      Width           =   7215
      _ExtentX        =   0
      _ExtentY        =   0
      Begin PhotoDemon.pdMetadataExport mtdManager 
         Height          =   4215
         Left            =   240
         TabIndex        =   4
         Top             =   120
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   7435
      End
   End
End
Attribute VB_Name = "dialog_ExportJXL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'JPEG XL Export Dialog
'Copyright 2022-2022 by Tanner Helland
'Created: 08/November/22
'Last updated: 08/November/22
'Last update: initial build
'
'Dialog for presenting the user various options related to JPEG XL exporting.  All export options rely on
' libjxl for their actual implementation.
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

'Current original preview DIB, cropped and zoomed as necessary (but otherwise unmodified).
Private m_PreviewDIB As pdDIB

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

Private Sub UpdatePanelVisibility()
    Dim i As Long
    For i = 0 To btsCategory.ListCount - 1
        picContainer(i).Visible = (i = btsCategory.ListIndex)
    Next i
End Sub

Private Sub btsDepth_Click(ByVal buttonIndex As Long)
    UpdatePreviewSource
    UpdatePreview
End Sub

Private Sub btsQuality_Click(ByVal buttonIndex As Long)
    UpdateQualityVisibility
    UpdatePreview
End Sub

Private Sub cmdBar_CancelClick()
    m_UserDialogAnswer = vbCancel
    Me.Hide
End Sub

Private Sub cmdBar_OKClick()
    
    'Store all parameters inside an XML string
    m_FormatParamString = GetParamString_JXL()
    
    'The metadata panel manages its own XML string
    m_MetadataParamString = mtdManager.GetMetadataSettings
    
    'Free resources that are no longer required
    Set m_CompositedImage = Nothing
    Set m_SrcImage = Nothing
    Set m_PreviewDIB = Nothing
    
    'Hide but *DO NOT UNLOAD* the form.  The dialog manager needs to retrieve the setting strings before unloading us
    m_UserDialogAnswer = vbOK
    Me.Visible = False
    
End Sub

Private Function GetParamString_JXL() As String
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.AddParam "jxl-lossless", (btsQuality.ListIndex = 0)
    cParams.AddParam "jxl-lossy-quality", sldQuality.Value
    cParams.AddParam "jxl-effort", sldEffort.Value
    
    Select Case btsDepth.ListIndex
        Case 0
            cParams.AddParam "jxl-color-format", "auto"
        Case 1
            cParams.AddParam "jxl-color-format", "color"
        Case 2
            cParams.AddParam "jxl-color-format", "gray"
    End Select
    
    GetParamString_JXL = cParams.GetParamString
    
End Function

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub cmdBar_ResetClick()
    
    'Default save quality is "Excellent"
    sldQuality.Value = 1    'Visually lossless, but underlying RGB may change due to color space conversion(s)
    sldEffort.Value = 7     'Default per libjxl
    
    'Auto color model detection
    btsDepth.ListIndex = 0
    
    mtdManager.Reset
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'The ShowDialog routine presents the user with this form.
Public Sub ShowDialog(Optional ByRef srcImage As pdImage = Nothing)
    
    'Provide a default answer of "cancel" (in the event that the user clicks the "x" button in the top-right)
    m_UserDialogAnswer = vbCancel
    
    'Make sure that the proper cursor is set
    Screen.MousePointer = 0
    Message "Waiting for user to specify export options... "
    
    'Populate the category button strip
    btsCategory.AddItem "basic", 0
    btsCategory.AddItem "advanced", 1
    btsCategory.AddItem "metadata", 2
    
    'Populate the "basic" options panel
    btsQuality.AddItem "lossless", 0
    btsQuality.AddItem "lossy", 1
    btsQuality.ListIndex = 0
    UpdateQualityVisibility
    
    'Populate the "advanced" options panel
    btsDepth.AddItem "auto", 0
    btsDepth.AddItem "color", 1
    btsDepth.AddItem "grayscale", 2
    btsDepth.ListIndex = 0
    
    'Next, prepare various controls on the metadata panel
    Set m_SrcImage = srcImage
    mtdManager.SetParentImage m_SrcImage, PDIF_JXL
    
    'By default, the basic options panel is always shown.
    btsCategory.ListIndex = 0
    UpdatePanelVisibility
    
    'Make a copy of the composited image; it takes time to composite layers, so we don't want to redo this except
    ' when absolutely necessary.
    If (Not m_SrcImage Is Nothing) Then
        m_SrcImage.GetCompositedImage m_CompositedImage, True
        pdFxPreview.NotifyNonStandardSource m_CompositedImage.GetDIBWidth, m_CompositedImage.GetDIBHeight
    End If
    
    'In batch process mode, we won't have a sample image to preview
    If (m_SrcImage Is Nothing) Then Interface.ShowDisabledPreviewImage pdFxPreview
    
    'Update the preview
    UpdatePreviewSource
    UpdatePreview True
    
    'Apply translations and visual themes
    ApplyThemeAndTranslations Me
    Interface.SetFormCaptionW Me, g_Language.TranslateMessage("%1 options", "JPEG XL")
    If (Not g_WindowManager Is Nothing) Then g_WindowManager.SetFocusAPI cmdBar.hWnd
    
    'Display the dialog
    ShowPDDialog vbModal, Me, True

End Sub

Private Sub pdFxPreview_ViewportChanged()
    UpdatePreviewSource
    UpdatePreview
End Sub

Private Sub sldEffort_Change()
    UpdatePreview
End Sub

Private Sub sldQuality_Change()
    UpdatePreview
End Sub

'The dialog differentiates between lossless and lossy using a hard toggle; lossy settings are hidden
' when lossless mode is requested.
Private Sub UpdateQualityVisibility()
    
    Dim showLossySettings As Boolean
    showLossySettings = (btsQuality.ListIndex > 0)
    
    sldQuality.Visible = showLossySettings
    lblHint(2).Visible = showLossySettings
    lblHint(3).Visible = showLossySettings
    
End Sub

'When a parameter changes that requires a new base image for the preview (e.g. changing the background composite color),
' call this function to generate a new preview DIB.  Note that you *do not* need to call this function for format-specific
' changes (like quality, subsampling, etc).
Private Sub UpdatePreviewSource()

    If (Not m_CompositedImage Is Nothing) Then
        
        'Because the user can change the preview viewport, we can't guarantee that the preview region hasn't changed
        ' since the last preview.  Prep a new preview now.
        Dim tmpSafeArray As SafeArray2D
        EffectPrep.PreviewNonStandardImage tmpSafeArray, m_CompositedImage, pdFxPreview, False
        
        'The public workingDIB object now contains the preview area image.  Clone it locally.
        If (m_PreviewDIB Is Nothing) Then Set m_PreviewDIB = New pdDIB
        m_PreviewDIB.CreateFromExistingDIB workingDIB
        
        'Perform a one-time swizzle here (from BGRA to RGBA order)
        DIBs.SwizzleBR m_PreviewDIB
        
        'TODO: forcible color-depth changes here?
        
    End If
    
End Sub

Private Sub UpdatePreview(Optional ByVal forceUpdate As Boolean = False)

    If (cmdBar.PreviewsAllowed Or forceUpdate) And (Not m_SrcImage Is Nothing) And (Not m_PreviewDIB Is Nothing) Then
        
        'Retrieve a JPEG-XL version of the current preview image.
        If Plugin_jxl.PreviewJXL(m_PreviewDIB, workingDIB, GetParamString_JXL()) Then
            FinalizeNonstandardPreview pdFxPreview, True
        Else
            PDDebug.LogAction "WARNING: JPEG-XL EXPORT PREVIEW PROBLEM!"
        End If
        
    End If
    
End Sub
