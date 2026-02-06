VERSION 5.00
Begin VB.Form dialog_ExportGIF 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   7230
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13095
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
   Icon            =   "File_Save_GIF.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   482
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   873
   Begin PhotoDemon.pdCommandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   6480
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   1323
   End
   Begin PhotoDemon.pdFxPreviewCtl pdFxPreview 
      Height          =   6225
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   10980
      ColorSelection  =   -1  'True
   End
   Begin PhotoDemon.pdButtonStrip btsCategory 
      Height          =   615
      Left            =   5880
      TabIndex        =   10
      Top             =   120
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   1085
      FontSize        =   11
   End
   Begin PhotoDemon.pdContainer picContainer 
      Height          =   5535
      Index           =   0
      Left            =   5880
      Top             =   840
      Width           =   7095
      _ExtentX        =   0
      _ExtentY        =   0
      Begin PhotoDemon.pdSlider sldAlphaCutoff 
         Height          =   855
         Left            =   0
         TabIndex        =   3
         Top             =   4440
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   1508
         Caption         =   "alpha cut-off"
         Max             =   254
         SliderTrackStyle=   1
         Value           =   64
         GradientColorRight=   1703935
         NotchPosition   =   2
         NotchValueCustom=   64
      End
      Begin PhotoDemon.pdCheckBox chkColorCount 
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   1200
         Width           =   6975
         _ExtentX        =   7858
         _ExtentY        =   661
         Caption         =   "restrict palette size"
         Value           =   0   'False
      End
      Begin PhotoDemon.pdColorSelector clsBackground 
         Height          =   975
         Left            =   0
         TabIndex        =   5
         Top             =   2160
         Width           =   7095
         _ExtentX        =   15690
         _ExtentY        =   1720
         Caption         =   "background color"
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   375
         Index           =   0
         Left            =   3480
         Top             =   1740
         Width           =   3615
         _ExtentX        =   9340
         _ExtentY        =   661
         Caption         =   "unique colors"
      End
      Begin PhotoDemon.pdSlider sldColorCount 
         Height          =   375
         Left            =   360
         TabIndex        =   6
         Top             =   1680
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   661
         Min             =   2
         Max             =   256
         Value           =   256
         NotchPosition   =   2
         NotchValueCustom=   256
      End
      Begin PhotoDemon.pdButtonStrip btsAlpha 
         Height          =   1095
         Left            =   0
         TabIndex        =   7
         Top             =   3240
         Width           =   7095
         _ExtentX        =   15690
         _ExtentY        =   1931
         Caption         =   "transparency"
      End
      Begin PhotoDemon.pdButtonStrip btsColorModel 
         Height          =   1095
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   7095
         _ExtentX        =   15690
         _ExtentY        =   1931
         Caption         =   "color model"
      End
      Begin PhotoDemon.pdColorSelector clsAlphaColor 
         Height          =   975
         Left            =   0
         TabIndex        =   9
         Top             =   4440
         Width           =   7095
         _ExtentX        =   15690
         _ExtentY        =   1720
         Caption         =   "transparent color (right-click image to select)"
         curColor        =   16711935
      End
   End
   Begin PhotoDemon.pdContainer picContainer 
      Height          =   5535
      Index           =   1
      Left            =   5880
      Top             =   840
      Width           =   7095
      _ExtentX        =   0
      _ExtentY        =   0
      Begin PhotoDemon.pdMetadataExport mtdManager 
         Height          =   4935
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   8705
      End
   End
End
Attribute VB_Name = "dialog_ExportGIF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'GIF export dialog
'Copyright 2012-2026 by Tanner Helland
'Created: 11/December/12
'Last updated: 20/October/21
'Last update: remove all FreeImage dependencies; everything here is 100% homebrew, baby!
'
'UI for GIF export settings.  Live previews are available for all relevant features.  Note that some
' ancient GIF features (like interlacing) are deliberately hidden to prevent users from producing large,
' slow GIFs for no good reason.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'This form should always be notified of the image being exported.  (One exception to this rule is
' invoking the dialog from the batch process dialog, as no image is associated with that preview.
' This case is automatically covered by this dialog.)
Private m_SrcImage As pdImage

'A composite of the current image, 32-bpp, fully composited.  This only needs to be regenerated
' if the source image changes.
Private m_CompositedImage As pdDIB

'OK or CANCEL result
Private m_UserDialogAnswer As VbMsgBoxResult

'Final format-specific XML packet, with all format-specific settings defined as tag+value pairs
Private m_FormatParamString As String

'Final metadata XML packet, with all metadata settings defined as tag+value pairs.  Currently unused as ExifTool
' cannot write any BMP-specific data.
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

'The ShowDialog routine presents the user with this form.
Public Sub ShowDialog(Optional ByRef srcImage As pdImage = Nothing)

    'Provide a default answer of "cancel" (in the event that the user clicks the "x" button in the top-right)
    m_UserDialogAnswer = vbCancel
    
    Message "Waiting for user to specify export options... "
    
    'Populate the category button strip
    btsCategory.AddItem "basic", 0
    btsCategory.AddItem "metadata", 1
    btsCategory.ListIndex = 0
    
    btsColorModel.AddItem "auto", 0
    btsColorModel.AddItem "color", 1
    btsColorModel.AddItem "grayscale", 2
    
    btsAlpha.AddItem "auto", 0
    btsAlpha.AddItem "none", 1
    btsAlpha.AddItem "by cut-off", 2
    btsAlpha.AddItem "by color", 3
    
    sldAlphaCutoff.NotchValueCustom = PD_DEFAULT_ALPHA_CUTOFF
    
    'Prep a preview (if any)
    Set m_SrcImage = srcImage
    If Not (m_SrcImage Is Nothing) Then
        m_SrcImage.GetCompositedImage m_CompositedImage, True
        pdFxPreview.NotifyNonStandardSource m_CompositedImage.GetDIBWidth, m_CompositedImage.GetDIBHeight
    End If
    If (m_SrcImage Is Nothing) Then Interface.ShowDisabledPreviewImage pdFxPreview
    
    'Next, prepare various controls on the metadata panel
    mtdManager.SetParentImage m_SrcImage, PDIF_GIF
    
    'Update the preview
    UpdatePanelVisibility
    UpdateAllVisibility
    UpdateTransparencyOptions
    UpdatePreview
    
    'Apply translations and visual themes
    ApplyThemeAndTranslations Me, True, True
    Interface.SetFormCaptionW Me, g_Language.TranslateMessage("%1 options", "GIF")
    
    'Display the dialog
    ShowPDDialog vbModal, Me, True

End Sub

Private Sub btsAlpha_Click(ByVal buttonIndex As Long)
    UpdateTransparencyOptions
    UpdatePreview
End Sub

Private Sub UpdateTransparencyOptions()
    
    Select Case btsAlpha.ListIndex
    
        'auto
        Case 0
            sldAlphaCutoff.Visible = False
            clsAlphaColor.Visible = False
            pdFxPreview.AllowColorSelection = False
            
        'no alpha
        Case 1
            sldAlphaCutoff.Visible = False
            clsAlphaColor.Visible = False
            pdFxPreview.AllowColorSelection = False
        
        'alpha by cut-off
        Case 2
            sldAlphaCutoff.Visible = True
            clsAlphaColor.Visible = False
            pdFxPreview.AllowColorSelection = False
        
        'alpha by color
        Case 3
            sldAlphaCutoff.Visible = False
            clsAlphaColor.Visible = True
            pdFxPreview.AllowColorSelection = True
    
    End Select
    
End Sub

Private Sub btsCategory_Click(ByVal buttonIndex As Long)
    UpdatePanelVisibility
End Sub

Private Sub UpdatePanelVisibility()
    Dim i As Long
    For i = 0 To btsCategory.ListCount - 1
        picContainer(i).Visible = (i = btsCategory.ListIndex)
    Next i
End Sub

Private Sub btsColorModel_Click(ByVal buttonIndex As Long)
    UpdateAllVisibility
    UpdatePreview
End Sub

Private Sub UpdateAllVisibility()

    Select Case btsColorModel.ListIndex
    
        'Auto
        Case 0
            UpdateColorCountVisibility False
            
        'Color
        Case 1
            UpdateColorCountVisibility True
            
        'Grayscale
        Case 2
            UpdateColorCountVisibility True
            
    End Select
    
End Sub

Private Sub UpdateColorCountVisibility(ByVal newValue As Boolean)
    chkColorCount.Visible = newValue
    sldColorCount.Visible = newValue
    lblTitle(0).Visible = newValue
End Sub

Private Sub chkColorCount_Click()
    UpdatePreview
End Sub

Private Sub clsAlphaColor_ColorChanged()
    UpdatePreview
End Sub

Private Sub clsBackground_ColorChanged()
    UpdatePreview
End Sub

Private Sub cmdBar_CancelClick()
    m_UserDialogAnswer = vbCancel
    Me.Visible = False
End Sub

Private Sub cmdBar_OKClick()
    m_FormatParamString = GetExportParamString
    m_MetadataParamString = mtdManager.GetMetadataSettings
    m_UserDialogAnswer = vbOK
    Me.Visible = False
End Sub

Private Sub cmdBar_ReadCustomPresetData()
    UpdateAllVisibility
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub cmdBar_ResetClick()
    btsColorModel.ListIndex = 0
    chkColorCount.Value = False
    sldColorCount.Value = 256
    clsBackground.Color = vbWhite
    btsAlpha.ListIndex = 0
    sldAlphaCutoff.Value = PD_DEFAULT_ALPHA_CUTOFF
    clsAlphaColor.Color = RGB(255, 0, 255)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Function GetExportParamString() As String

    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    
    'Convert the color depth option buttons into a usable numeric value
    Dim outputColorMode As String
    
    Select Case btsColorModel.ListIndex
        Case 0
            outputColorMode = "auto"
        Case 1
            outputColorMode = "color"
        Case 2
            outputColorMode = "gray"
    End Select
    
    cParams.AddParam "gif-color-mode", outputColorMode
    
    Dim outputAlphaMode As String
    Select Case btsAlpha.ListIndex
        Case 0
            outputAlphaMode = "auto"
        Case 1
            outputAlphaMode = "none"
        Case 2
            outputAlphaMode = "by-cutoff"
        Case 3
            outputAlphaMode = "by-color"
    End Select
    
    cParams.AddParam "gif-alpha-mode", outputAlphaMode
    
    'If "auto" mode is selected, we currently enforce a hard-coded cut-off value.
    ' There may be a better way to do this (Otsu's, but for alpha?), but I haven't investigated in detail.
    Dim outputAlphaCutoff As Long
    If (btsAlpha.ListIndex = 0) Or (Not sldAlphaCutoff.IsValid) Then outputAlphaCutoff = PD_DEFAULT_ALPHA_CUTOFF Else outputAlphaCutoff = sldAlphaCutoff.Value
    cParams.AddParam "gif-alpha-cutoff", outputAlphaCutoff
    
    Dim colorCount As Long
    If (btsColorModel.ListIndex <> 0) Then
        If chkColorCount.Value And sldColorCount.IsValid Then colorCount = sldColorCount.Value Else colorCount = 256
    Else
        colorCount = 256
    End If
    cParams.AddParam "gif-color-count", colorCount
    cParams.AddParam "gif-backcolor", clsBackground.Color
    cParams.AddParam "gif-alpha-color", clsAlphaColor.Color
    
    GetExportParamString = cParams.GetParamString
    
End Function

Private Sub pdFxPreview_ColorSelected()
    clsAlphaColor.Color = pdFxPreview.SelectedColor
End Sub

Private Sub pdFxPreview_ViewportChanged()
    UpdateAllVisibility
    UpdatePreview
End Sub

Private Sub UpdatePreview()

    If (cmdBar.PreviewsAllowed And sldColorCount.IsValid And (Not m_CompositedImage Is Nothing)) Then
        
        'Because the user can change the preview viewport, we can't guarantee that the preview region
        ' hasn't changed since the last preview.  Prep a new preview now.
        Dim tmpSafeArray As SafeArray2D
        EffectPrep.PreviewNonStandardImage tmpSafeArray, m_CompositedImage, pdFxPreview
        
        'This export dialog is a little weird because we need to preview a lot of possible settings,
        ' but none of them require the GIF encoder.  Instead, they are just convenience options
        ' for converting a 32-bpp image to an 8-bpp GIF with (or without) transparency.
        '
        'As such, we won't actually involve the GIF encoder at all here.  Just create a pdGIF class
        ' and request the pre-processing step by itself.
        Dim cGIF As pdGIF
        Set cGIF = New pdGIF
        cGIF.GetGifReadyImage workingDIB, GetExportParamString(), True
        FinalizeNonstandardPreview pdFxPreview, False
        
    End If
    
End Sub

Private Sub sldAlphaCutoff_Change()
    UpdatePreview
End Sub

Private Sub sldColorCount_Change()
    If (Not chkColorCount.Value) Then chkColorCount.Value = True
    UpdatePreview
End Sub
