VERSION 5.00
Begin VB.Form dialog_ExportPNG 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " PNG export options"
   ClientHeight    =   7635
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   13110
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
   ScaleHeight     =   509
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   874
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.pdButtonStrip btsMasterType 
      Height          =   735
      Left            =   5880
      TabIndex        =   2
      Top             =   120
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   1296
      FontSize        =   12
   End
   Begin PhotoDemon.pdCommandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   6885
      Width           =   13110
      _ExtentX        =   23125
      _ExtentY        =   1323
   End
   Begin PhotoDemon.pdFxPreviewCtl pdFxPreview 
      Height          =   6615
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   11668
      ColorSelection  =   -1  'True
   End
   Begin VB.PictureBox picCategory 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      HasDC           =   0   'False
      Height          =   5895
      Index           =   1
      Left            =   5880
      ScaleHeight     =   393
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   473
      TabIndex        =   19
      Top             =   960
      Width           =   7095
      Begin PhotoDemon.pdButton cmdUpdateLossyPreview 
         Height          =   615
         Left            =   360
         TabIndex        =   29
         Top             =   3480
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   1085
         Caption         =   "click to generate a new preview image"
      End
      Begin PhotoDemon.pdTitle ttlWebOptimize 
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   27
         Top             =   120
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   661
         Caption         =   "lossy optimization options"
         FontSize        =   12
      End
      Begin PhotoDemon.pdCheckBox chkOptimizeDither 
         Height          =   375
         Left            =   360
         TabIndex        =   25
         Top             =   1080
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   661
         Caption         =   "use dithering to improve quality"
      End
      Begin PhotoDemon.pdSlider sltTargetQuality 
         Height          =   735
         Left            =   360
         TabIndex        =   23
         Top             =   1560
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   1296
         Caption         =   "target quality"
         FontSizeCaption =   10
         Max             =   100
         Value           =   80
         NotchPosition   =   2
         NotchValueCustom=   80
      End
      Begin PhotoDemon.pdCheckBox chkOptimizeLossy 
         Height          =   375
         Left            =   360
         TabIndex        =   22
         Top             =   630
         Width           =   6735
         _ExtentX        =   12515
         _ExtentY        =   661
         Caption         =   "apply lossy optimizations"
      End
      Begin PhotoDemon.pdSlider sltLossyPerformance 
         Height          =   735
         Left            =   360
         TabIndex        =   24
         Top             =   2310
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   1296
         Caption         =   "optimization level"
         FontSizeCaption =   10
         Value           =   8
         NotchPosition   =   2
         NotchValueCustom=   8
      End
      Begin PhotoDemon.pdSlider sltLosslessPerformance 
         Height          =   735
         Left            =   360
         TabIndex        =   26
         Top             =   4800
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   1296
         Caption         =   "optimization level"
         FontSizeCaption =   10
         Max             =   7
         Value           =   2
         NotchPosition   =   2
         NotchValueCustom=   2
      End
      Begin PhotoDemon.pdLabel lblHint 
         Height          =   255
         Index           =   2
         Left            =   510
         Top             =   3090
         Width           =   2580
         _ExtentX        =   4551
         _ExtentY        =   450
         Caption         =   "fast, larger file"
         FontItalic      =   -1  'True
         FontSize        =   9
      End
      Begin PhotoDemon.pdLabel lblHint 
         Height          =   255
         Index           =   3
         Left            =   3180
         Top             =   3090
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   450
         Alignment       =   1
         Caption         =   "slow, smaller file"
         FontItalic      =   -1  'True
         FontSize        =   9
      End
      Begin PhotoDemon.pdLabel lblHint 
         Height          =   255
         Index           =   4
         Left            =   525
         Top             =   5580
         Width           =   2580
         _ExtentX        =   4551
         _ExtentY        =   450
         Caption         =   "fast, larger file"
         FontItalic      =   -1  'True
         FontSize        =   9
      End
      Begin PhotoDemon.pdLabel lblHint 
         Height          =   255
         Index           =   5
         Left            =   3180
         Top             =   5580
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   450
         Alignment       =   1
         Caption         =   "slow, smaller file"
         FontItalic      =   -1  'True
         FontSize        =   9
      End
      Begin PhotoDemon.pdTitle ttlWebOptimize 
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   28
         Top             =   4320
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   661
         Caption         =   "lossless optimization options"
         FontSize        =   12
      End
   End
   Begin VB.PictureBox picCategory 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      HasDC           =   0   'False
      Height          =   5775
      Index           =   0
      Left            =   5880
      ScaleHeight     =   385
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   473
      TabIndex        =   3
      Top             =   960
      Width           =   7095
      Begin PhotoDemon.pdButtonStrip btsCategory 
         Height          =   615
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   1085
      End
      Begin VB.PictureBox picContainer 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         HasDC           =   0   'False
         Height          =   5175
         Index           =   1
         Left            =   0
         ScaleHeight     =   345
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   473
         TabIndex        =   4
         Top             =   720
         Width           =   7095
         Begin PhotoDemon.pdSlider sldAlphaCutoff 
            Height          =   855
            Left            =   0
            TabIndex        =   5
            Top             =   4080
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
         Begin PhotoDemon.pdLabel lblColorCount 
            Height          =   375
            Left            =   4920
            Top             =   2460
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   661
            Caption         =   "palette size"
         End
         Begin PhotoDemon.pdSlider sldColorCount 
            Height          =   375
            Left            =   240
            TabIndex        =   6
            Top             =   2400
            Width           =   4575
            _ExtentX        =   8070
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
            Top             =   2880
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
         Begin PhotoDemon.pdButtonStrip btsDepthColor 
            Height          =   1095
            Left            =   0
            TabIndex        =   16
            Top             =   1200
            Width           =   7095
            _ExtentX        =   15690
            _ExtentY        =   1931
            Caption         =   "depth"
         End
         Begin PhotoDemon.pdColorSelector clsAlphaColor 
            Height          =   975
            Left            =   0
            TabIndex        =   9
            Top             =   4080
            Width           =   7095
            _ExtentX        =   15690
            _ExtentY        =   1720
            Caption         =   "transparent color (right-click image to select)"
            curColor        =   16711935
         End
         Begin PhotoDemon.pdButtonStrip btsDepthGrayscale 
            Height          =   1095
            Left            =   0
            TabIndex        =   20
            Top             =   1200
            Width           =   7095
            _ExtentX        =   15690
            _ExtentY        =   1931
            Caption         =   "depth"
         End
      End
      Begin VB.PictureBox picContainer 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         HasDC           =   0   'False
         Height          =   5175
         Index           =   0
         Left            =   0
         ScaleHeight     =   345
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   473
         TabIndex        =   13
         Top             =   720
         Width           =   7095
         Begin PhotoDemon.pdLabel lblHint 
            Height          =   255
            Index           =   0
            Left            =   180
            Top             =   960
            Width           =   2700
            _ExtentX        =   4763
            _ExtentY        =   450
            Caption         =   "fast, larger file"
            FontItalic      =   -1  'True
            FontSize        =   9
         End
         Begin PhotoDemon.pdCheckBox chkInterlace 
            Height          =   375
            Left            =   120
            TabIndex        =   15
            Top             =   1440
            Width           =   6975
            _ExtentX        =   12303
            _ExtentY        =   661
            Caption         =   "use interlacing"
            Value           =   0
         End
         Begin PhotoDemon.pdSlider sldCompression 
            Height          =   735
            Left            =   0
            TabIndex        =   14
            Top             =   240
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   1720
            Caption         =   "compression level"
            Max             =   9
            Value           =   3
            GradientColorRight=   1703935
            NotchPosition   =   2
            NotchValueCustom=   3
         End
         Begin PhotoDemon.pdColorSelector clsBackground 
            Height          =   975
            Left            =   0
            TabIndex        =   17
            Top             =   2160
            Width           =   7095
            _ExtentX        =   15690
            _ExtentY        =   1720
            Caption         =   "background color"
         End
         Begin PhotoDemon.pdCheckBox chkEmbedBackground 
            Height          =   375
            Left            =   120
            TabIndex        =   18
            Top             =   3240
            Width           =   6975
            _ExtentX        =   12303
            _ExtentY        =   661
            Caption         =   "embed background color in file"
            Value           =   0
         End
         Begin PhotoDemon.pdLabel lblHint 
            Height          =   255
            Index           =   1
            Left            =   3075
            Top             =   960
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   450
            Alignment       =   1
            Caption         =   "slow, smaller file"
            FontItalic      =   -1  'True
            FontSize        =   9
         End
         Begin PhotoDemon.pdButtonStrip btsStandardOptimize 
            Height          =   1095
            Left            =   0
            TabIndex        =   21
            Top             =   3960
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   1931
            Caption         =   "file size optimization"
         End
      End
      Begin VB.PictureBox picContainer 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         HasDC           =   0   'False
         Height          =   5175
         Index           =   2
         Left            =   0
         ScaleHeight     =   345
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   473
         TabIndex        =   11
         Top             =   720
         Width           =   7095
         Begin PhotoDemon.pdMetadataExport mtdManager 
            Height          =   3255
            Left            =   0
            TabIndex        =   12
            Top             =   0
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   5741
         End
      End
   End
End
Attribute VB_Name = "dialog_ExportPNG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'PNG export dialog
'Copyright 2012-2016 by Tanner Helland
'Created: 11/December/12
'Last updated: 21/April/16
'Last update: repurpose old color-depth dialog into a PNG-specific one
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
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

'Final metadata XML packet, with all metadata settings defined as tag+value pairs.  Currently unused as ExifTool
' cannot write any BMP-specific data.
Private m_MetadataParamString As String

'Default alpha cut-off when "auto" is selected
Private Const DEFAULT_ALPHA_CUTOFF As Long = 64

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

Private Sub btsAlpha_Click(ByVal buttonIndex As Long)
    UpdateTransparencyOptions
    UpdatePreviewSource
    UpdatePreview
End Sub

Private Sub UpdateTransparencyOptions()
    
    Select Case btsAlpha.ListIndex
    
        'auto, full alpha
        Case 0, 1
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
            
        'no alpha
        Case 4
            sldAlphaCutoff.Visible = False
            clsAlphaColor.Visible = False
            pdFxPreview.AllowColorSelection = False
    
    End Select
    
    ReflowColorPanel
    
End Sub

Private Sub btsCategory_Click(ByVal buttonIndex As Long)
    UpdatePanelVisibility
End Sub

Private Sub UpdatePanelVisibility()
    Dim i As Long
    For i = 0 To btsCategory.ListCount - 1
        picContainer(i).Visible = CBool(i = btsCategory.ListIndex)
    Next i
End Sub

Private Sub btsColorModel_Click(ByVal buttonIndex As Long)
    UpdateColorDepthVisibility
    UpdatePreviewSource
    UpdatePreview
End Sub

Private Sub UpdateColorDepthVisibility()

    Select Case btsColorModel.ListIndex
    
        'Auto
        Case 0
            btsDepthColor.Visible = False
            btsDepthGrayscale.Visible = False
        
        'Color
        Case 1
            btsDepthColor.Visible = True
            btsDepthGrayscale.Visible = False
        
        'Grayscale
        Case 2
            btsDepthColor.Visible = False
            btsDepthGrayscale.Visible = True
    
    End Select

    UpdateColorDepthOptions

End Sub

Private Sub UpdateColorDepthOptions()
    
    'Indexed color modes allow for variable palette sizes
    If (btsDepthColor.Visible) Then
        sldColorCount.Visible = CBool(btsDepthColor.ListIndex = 2)
        lblColorCount.Visible = sldColorCount.Visible
    
    'Indexed grayscale mode also allows for variable palette sizes
    ElseIf (btsDepthGrayscale.Visible) Then
        sldColorCount.Visible = CBool(btsDepthGrayscale.ListIndex = 1)
        lblColorCount.Visible = sldColorCount.Visible
    
    'Other modes do not expose palette settings
    Else
        sldColorCount.Visible = False
        lblColorCount.Visible = False
    End If
    
    ReflowColorPanel
    
End Sub

Private Sub UpdateMasterPanelVisibility()
    Dim i As Long
    For i = picCategory.lBound To picCategory.UBound
        picCategory(i).Visible = CBool(btsMasterType.ListIndex = i)
    Next i
End Sub

Private Sub ReflowColorPanel()

    Dim yOffset As Long, yPadding As Long
    yOffset = btsColorModel.GetTop + btsColorModel.GetHeight
    yPadding = FixDPI(8)
    yOffset = yOffset + yPadding
    
    If btsDepthColor.Visible Then
        btsDepthColor.SetTop yOffset
        yOffset = yOffset + btsDepthColor.GetHeight + yPadding
    ElseIf btsDepthGrayscale.Visible Then
        btsDepthGrayscale.SetTop yOffset
        yOffset = yOffset + btsDepthGrayscale.GetHeight + yPadding
    End If
    
    If sldColorCount.Visible Then
        sldColorCount.SetTop yOffset
        lblColorCount.SetTop (sldColorCount.GetTop + sldColorCount.GetHeight) - lblColorCount.GetHeight
        yOffset = yOffset + sldColorCount.GetHeight + yPadding
    End If
    
    btsAlpha.SetTop yOffset
    yOffset = yOffset + btsAlpha.GetHeight + yPadding
    
    If sldAlphaCutoff.Visible Then
        sldAlphaCutoff.SetTop yOffset
    ElseIf clsAlphaColor.Visible Then
        clsAlphaColor.SetTop yOffset
    End If
    
End Sub

Private Sub btsDepthColor_Click(ByVal buttonIndex As Long)
    UpdateColorDepthOptions
    UpdatePreviewSource
    UpdatePreview
End Sub

Private Sub btsDepthGrayscale_Click(ByVal buttonIndex As Long)
    UpdateColorDepthOptions
    UpdatePreviewSource
    UpdatePreview
End Sub

Private Sub btsMasterType_Click(ByVal buttonIndex As Long)
    UpdateMasterPanelVisibility
End Sub

Private Sub chkOptimizeDither_Click()
    UpdatePreviewButtonText
End Sub

Private Sub chkOptimizeLossy_Click()
    EnableLossyOptimizationOptions
End Sub

Private Sub EnableLossyOptimizationOptions()
    
    Dim enabledState As Boolean
    enabledState = CBool(chkOptimizeLossy.Value)
    
    chkOptimizeDither.Enabled = enabledState
    sltTargetQuality.Enabled = enabledState
    sltLossyPerformance.Enabled = enabledState
    lblHint(2).Enabled = enabledState
    lblHint(3).Enabled = enabledState
    cmdUpdateLossyPreview.Enabled = enabledState
    
End Sub

Private Sub UpdatePreviewButtonText()
    If (StrComp(cmdUpdateLossyPreview.Caption, g_Language.TranslateMessage("click to generate a new preview image"), vbBinaryCompare) <> 0) Then
        cmdUpdateLossyPreview.Caption = g_Language.TranslateMessage("click to generate a new preview image")
    End If
End Sub

Private Sub clsAlphaColor_ColorChanged()
    UpdatePreviewSource
    UpdatePreview
End Sub

Private Sub clsBackground_ColorChanged()
    UpdatePreviewSource
    UpdatePreview
End Sub

Private Sub cmdBar_CancelClick()
    m_UserDialogAnswer = vbCancel
    Me.Hide
End Sub

Private Sub cmdBar_OKClick()
    
    m_FormatParamString = GetExportParamString
    
    If (btsMasterType.ListIndex = 0) Then
        m_MetadataParamString = mtdManager.GetMetadataSettings
    
    'While in web optimization mode, we forcibly request no metadata writing
    Else
        m_MetadataParamString = mtdManager.GetNullMetadataSettings
    End If
    
    m_UserDialogAnswer = vbOK
    Me.Hide
    
End Sub

Private Sub cmdBar_ReadCustomPresetData()
    ReflowWebOptimizePanel
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub cmdBar_ResetClick()
    
    cmdBar.MarkPreviewStatus False
    
    'General panel settings
    sldCompression.Value = sldCompression.NotchValueCustom
    chkInterlace.Value = vbUnchecked
    
    If (Not m_SrcImage Is Nothing) Then
        If m_SrcImage.imgStorage.DoesKeyExist("pngBackgroundColor") Then
            clsBackground.Color = m_SrcImage.imgStorage.GetEntry_Long("pngBackgroundColor")
            chkEmbedBackground.Value = vbChecked
        Else
            clsBackground.Color = vbWhite
            chkEmbedBackground.Value = vbUnchecked
        End If
    Else
        clsBackground.Color = vbWhite
        chkEmbedBackground.Value = vbUnchecked
    End If
    
    'Color and transparency settings
    btsColorModel.ListIndex = 0
    btsDepthColor.ListIndex = 1
    btsDepthGrayscale.ListIndex = 1
    btsAlpha.ListIndex = 0
    
    sldColorCount.Value = 256
    sldAlphaCutoff.Value = DEFAULT_ALPHA_CUTOFF
    clsAlphaColor.Color = RGB(255, 0, 255)
    
    'Web-optimized settings
    chkOptimizeLossy.Value = vbChecked
    sltTargetQuality.Value = sltTargetQuality.NotchValueCustom
    sltLossyPerformance.Value = sltLossyPerformance.NotchValueCustom
    chkOptimizeDither.Value = vbChecked
    sltLosslessPerformance.Value = sltLosslessPerformance.NotchValueCustom
    
    'Metadata settings
    mtdManager.Reset
    
    cmdBar.MarkPreviewStatus True
    UpdatePreviewSource
    UpdatePreview
    
End Sub

'The ShowDialog routine presents the user with this form.
Public Sub ShowDialog(Optional ByRef srcImage As pdImage = Nothing)

    'Provide a default answer of "cancel" (in the event that the user clicks the "x" button in the top-right)
    m_UserDialogAnswer = vbCancel
    
    Message "Waiting for user to specify export options... "
    
    'PNG has two master panels: standard PNGs, and web-optimized PNGs
    btsMasterType.AddItem "standard PNG", 0
    btsMasterType.AddItem "web-optimized PNG", 1
    btsMasterType.ListIndex = 0
    
    'Populate the category button strip
    btsCategory.AddItem "basic", 0
    btsCategory.AddItem "advanced", 1
    btsCategory.AddItem "metadata", 2
    btsCategory.ListIndex = 0
    
    'Populate standard model options
    btsStandardOptimize.AddItem "none", 0
    btsStandardOptimize.AddItem "basic (default)", 1
    btsStandardOptimize.AddItem "moderate", 2
    btsStandardOptimize.AddItem "maximum", 3
    btsStandardOptimize.ListIndex = 1
    
    'Color model and color depth are closely related; populate all button strips, then show/hide the relevant pairings
    btsColorModel.AddItem "auto", 0
    btsColorModel.AddItem "color", 1
    btsColorModel.AddItem "grayscale", 2
    btsColorModel.ListIndex = 0
    
    btsDepthColor.AddItem "HDR", 0
    btsDepthColor.AddItem "standard", 1
    btsDepthColor.AddItem "indexed", 2
    btsDepthColor.ListIndex = 1
    
    btsDepthGrayscale.AddItem "HDR", 0
    btsDepthGrayscale.AddItem "standard", 1
    btsDepthGrayscale.AddItem "monochrome", 2
    btsDepthGrayscale.ListIndex = 1
    
    UpdateColorDepthVisibility
    
    'PNGs also support a (ridiculous) amount of alpha settings
    btsAlpha.AddItem "auto", 0
    btsAlpha.AddItem "full", 1
    btsAlpha.AddItem "binary (by cut-off)", 2
    btsAlpha.AddItem "binary (by color)", 3
    btsAlpha.AddItem "none", 4
    
    sldAlphaCutoff.NotchValueCustom = DEFAULT_ALPHA_CUTOFF
    
    'Populate web-optimized options
    EnableLossyOptimizationOptions
    
    'Prep a preview (if any)
    Set m_SrcImage = srcImage
    If Not (m_SrcImage Is Nothing) Then
        m_SrcImage.GetCompositedImage m_CompositedImage, True
        pdFxPreview.NotifyNonStandardSource m_CompositedImage.GetDIBWidth, m_CompositedImage.GetDIBHeight
    End If
    If (Not g_ImageFormats.FreeImageEnabled) Or (m_SrcImage Is Nothing) Then Interface.ShowDisabledPreviewImage pdFxPreview
    
    'Next, prepare various controls on the metadata panel
    mtdManager.SetParentImage m_SrcImage, PDIF_PNG
    
    'If the source image was a PNG, and it also contained a background color, retrieve and set the matching color now
    If (Not m_SrcImage Is Nothing) Then
        If m_SrcImage.imgStorage.DoesKeyExist("pngBackgroundColor") Then
            clsBackground.Color = m_SrcImage.imgStorage.GetEntry_Long("pngBackgroundColor")
            chkEmbedBackground.Value = vbChecked
        End If
    End If
    
    'Update the preview
    UpdatePreviewSource
    UpdatePreview
    
    'Apply translations and visual themes
    ApplyThemeAndTranslations Me
    
    'Many UI options are dynamically shown/hidden depending on other settings; make sure their initial state is correct
    UpdateMasterPanelVisibility
    UpdatePanelVisibility
    UpdateColorDepthVisibility
    UpdateTransparencyOptions
    ReflowWebOptimizePanel
    
    'Display the dialog
    ShowPDDialog vbModal, Me, True

End Sub

'Lossy previews are time-consuming to generate, so we cannot provide them "on-demand".  Instead, the user must
' manually compress them via button click.
Private Sub cmdUpdateLossyPreview_Click()

    cmdUpdateLossyPreview.Caption = g_Language.TranslateMessage("please wait while a new preview image is created...")
    
    Dim updateSuccess As Boolean
    updateSuccess = False
    
    'Make sure a composite image was created successfully
    If Not (m_CompositedImage Is Nothing) Then
        
        'Because the user can change the preview viewport, we can't guarantee that the preview region hasn't changed
        ' since the last preview.  Prep a new preview now.
        Dim tmpSafeArray As SAFEARRAY2D
        FastDrawing.PreviewNonStandardImage tmpSafeArray, m_CompositedImage, pdFxPreview, False
        
        'Create a FreeImage copy of the current preview image
        If (m_FIHandle <> 0) Then Plugin_FreeImage.ReleaseFreeImageObject m_FIHandle
        m_FIHandle = Plugin_FreeImage.GetFIDib_SpecificColorMode(workingDIB, 32, PDAS_ComplicatedAlpha, PDAS_ComplicatedAlpha)
        
        'Write that image out to a temporary file
        Dim tmpFilename As String
        tmpFilename = FileSystem.RequestTempFile()
        If FreeImage_Save(FIF_PNG, m_FIHandle, tmpFilename, FISO_PNG_Z_BEST_SPEED) Then
            
            'Retrieve the size of the base PNG file
            Dim cFile As pdFSO
            Set cFile = New pdFSO
            
            Dim oldFileSize As Long
            oldFileSize = cFile.FileLenW(tmpFilename)
            
            'Next, request optimization from pngquant
            If Plugin_PNGQuant.ApplyPNGQuantToFile_Synchronous(tmpFilename, sltTargetQuality.Value, 11 - sltLossyPerformance.Value, CBool(chkOptimizeDither.Value), False) Then
                
                Dim newFileSize As Long
                newFileSize = cFile.FileLenW(tmpFilename)
                
                'If successful, pngquant will overwrite the original file with its optimized copy.  Retrieve it now.
                If Loading.QuickLoadImageToDIB(tmpFilename, workingDIB, False) Then
                    FastDrawing.FinalizeNonstandardPreview Me.pdFxPreview, False
                    updateSuccess = True
                End If
                
            Else
                #If DEBUGMODE = 1 Then
                    pdDebug.LogAction "WARNING!  The pngquant preview step failed for reasons unknown!"
                #End If
            End If
            
            If cFile.FileExist(tmpFilename) Then cFile.KillFile tmpFilename
            
        End If
        
    End If
    
    If updateSuccess Then
        cmdUpdateLossyPreview.Caption = g_Language.TranslateMessage("These lossy optimization settings reduced file size by %1.", Format$((1 - (newFileSize / oldFileSize)) * 100, "0.0") & "%")
    Else
        UpdatePreviewButtonText
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
    Plugin_FreeImage.ReleasePreviewCache m_FIHandle
End Sub

Private Function GetExportParamString() As String

    Dim cParams As pdParamXML
    Set cParams = New pdParamXML
    
    'The parameters this function returns vary based on the current PNG mode (standard vs web-optimized).
    cParams.AddParam "PNGCreateWebOptimized", CBool(btsMasterType.ListIndex = 1)
    
    'Standard parameters are the more complicated ones, if you can believe it
    If (btsMasterType.ListIndex = 0) Then
    
        'Start with the standard PNG settings, which are consistent across all standard PNG types
        If sldCompression.IsValid Then cParams.AddParam "PNGCompressionLevel", sldCompression.Value Else cParams.AddParam "PNGCompressionLevel", sldCompression.NotchValueCustom
        cParams.AddParam "PNGInterlacing", CBool(chkInterlace.Value)
        cParams.AddParam "PNGBackgroundColor", clsBackground.Color
        cParams.AddParam "PNGCreateBkgdChunk", CBool(chkEmbedBackground.Value)
        cParams.AddParam "PNGStandardOptimization", btsStandardOptimize.ListIndex
        
        'Next come all the messy color-depth possibilities
        Dim outputColorModel As String
        Select Case btsColorModel.ListIndex
            Case 0
                outputColorModel = "Auto"
            Case 1
                outputColorModel = "Color"
            Case 2
                outputColorModel = "Gray"
        End Select
        cParams.AddParam "PNGColorModel", outputColorModel
        
        'Which color depth we write is contingent on the color model, as color and gray use different button strips.
        ' (Gray supports some depths that color does not, e.g. 1-bit and 4-bit.)
        Dim outputColorDepth As String, outputPaletteSize As String
        
        'Color modes
        If (btsColorModel.ListIndex = 1) Then
            
            Select Case btsDepthColor.ListIndex
                Case 0
                    outputColorDepth = "48"
                Case 1
                    outputColorDepth = "24"
                Case 2
                    outputColorDepth = "8"
                    If sldColorCount.IsValid Then outputPaletteSize = CStr(sldColorCount.Value) Else outputPaletteSize = "256"
            End Select
            
        'Gray modes
        ElseIf (btsColorModel.ListIndex = 2) Then
            
            Select Case btsDepthGrayscale.ListIndex
                Case 0
                    outputColorDepth = "16"
                Case 1
                    outputColorDepth = "8"
                    If sldColorCount.IsValid Then outputPaletteSize = CStr(sldColorCount.Value) Else outputPaletteSize = "256"
                Case 2
                    outputColorDepth = "1"
            End Select
        
        End If
        
        If (Len(outputColorDepth) <> 0) Then cParams.AddParam "PNGBitDepth", outputColorDepth
        If (Len(outputPaletteSize) <> 0) Then cParams.AddParam "PNGPaletteSize", outputPaletteSize
        
        'Next, we've got a bunch of possible alpha modes to deal with (uuuuuugh)
        Dim outputAlphaModel As String
        Select Case btsAlpha.ListIndex
            Case 0
                outputAlphaModel = "Auto"
            Case 1
                outputAlphaModel = "Full"
            Case 2
                outputAlphaModel = "ByCutoff"
            Case 3
                outputAlphaModel = "ByColor"
            Case 4
                outputAlphaModel = "None"
        End Select
        
        cParams.AddParam "PNGAlphaModel", outputAlphaModel
        If sldAlphaCutoff.IsValid Then cParams.AddParam "PNGAlphaCutoff", sldAlphaCutoff.Value Else cParams.AddParam "PNGAlphaCutoff", DEFAULT_ALPHA_CUTOFF
        cParams.AddParam "PNGAlphaColor", clsAlphaColor.Color
        
    'Remember: web-optimized parameters must not use any UI elements not visible from the web-optimization panel!
    Else
    
        cParams.AddParam "PNGOptimizeLossy", CBool(chkOptimizeLossy.Value)
        cParams.AddParam "PNGOptimizeLossyQuality", sltTargetQuality.Value
        
        'pngquant accepts this value on a 1-11 scale, with 1 being slowest and 11 being fastest.  We show the user a
        ' [0, 10] scale where [10] is slowest (like the other settings on the form); reset to the proper range now.
        cParams.AddParam "PNGOptimizeLossyPerformance", 11 - sltLossyPerformance.Value
        cParams.AddParam "PNGOptimizeLossyDithering", CBool(chkOptimizeDither.Value)
        
        cParams.AddParam "PNGOptimizeLosslessPerformance", sltLosslessPerformance.Value
        
    End If
    
    GetExportParamString = cParams.GetParamString
    
End Function

Private Sub pdFxPreview_ColorSelected()
    clsAlphaColor.Color = pdFxPreview.SelectedColor
End Sub

Private Sub pdFxPreview_ViewportChanged()
    UpdatePreviewSource
    UpdatePreview
End Sub

'When a parameter changes that requires a new source DIB for the preview (e.g. changing the background composite color,
' changing the output color depth), you must call this function to generate a new preview DIB.  Note that you *do not*
' need to call this function for format-specific changes (e.g. compression settings).
Private Sub UpdatePreviewSource()
    If Not (m_CompositedImage Is Nothing) Then
        
        'Because the user can change the preview viewport, we can't guarantee that the preview region hasn't changed
        ' since the last preview.  Prep a new preview now.
        Dim tmpSafeArray As SAFEARRAY2D
        FastDrawing.PreviewNonStandardImage tmpSafeArray, m_CompositedImage, pdFxPreview, True
        
        'To reduce the chance of bugs, we use the same parameter parsing technique as the core PNG encoder
        Dim cParams As pdParamXML
        Set cParams = New pdParamXML
        cParams.SetParamString GetExportParamString()
        
        'Color and grayscale modes require different processing, so start there
        Dim forceGrayscale As Boolean
        forceGrayscale = ParamsEqual(cParams.GetString("PNGColorModel", "Auto"), "Gray")
        
        'For 8-bit modes, grab a palette size.  (This parameter will be ignored in other color modes.)
        Dim newPaletteSize As Long
        newPaletteSize = cParams.GetLong("PNGPaletteSize", 256)
        
        Dim newColorDepth As Long
        
        If ParamsEqual(cParams.GetString("PNGColorModel", "Auto"), "Auto") Then
            newColorDepth = 32
        Else
            
            'HDR modes do not need to be previewed, so we forcibly downsample them here
            If forceGrayscale Then
                newColorDepth = cParams.GetLong("PNGBitDepth", 8)
                If newColorDepth > 8 Then newColorDepth = 8
                If newColorDepth = 1 Then
                    newPaletteSize = 2
                    newColorDepth = 8
                End If
            Else
                newColorDepth = cParams.GetLong("PNGBitDepth", 24)
                If newColorDepth = 48 Then newColorDepth = 24
                If newColorDepth = 64 Then newColorDepth = 32
            End If
        
        End If
        
        'Next comes transparency, which is somewhat messy because PNG alpha behavior deviates significantly from normal alpha behavior.
        Dim desiredAlphaMode As PD_ALPHA_STATUS, desiredAlphaCutoff As Long
        
        If ParamsEqual(cParams.GetString("PNGAlphaModel", "Auto"), "Auto") Or ParamsEqual(cParams.GetString("PNGAlphaModel", "Auto"), "Full") Then
            desiredAlphaMode = PDAS_ComplicatedAlpha
            If newColorDepth = 24 Then newColorDepth = 32
        ElseIf ParamsEqual(cParams.GetString("PNGAlphaModel", "Auto"), "None") Then
            desiredAlphaMode = PDAS_NoAlpha
            If newColorDepth = 32 Then newColorDepth = 24
            desiredAlphaCutoff = 0
        ElseIf ParamsEqual(cParams.GetString("PNGAlphaModel", "Auto"), "ByCutoff") Then
            desiredAlphaMode = PDAS_BinaryAlpha
            desiredAlphaCutoff = cParams.GetLong("PNGAlphaCutoff", DEFAULT_ALPHA_CUTOFF)
            If newColorDepth = 24 Then newColorDepth = 32
        ElseIf ParamsEqual(cParams.GetString("PNGAlphaModel", "Auto"), "ByColor") Then
            desiredAlphaMode = PDAS_NewAlphaFromColor
            desiredAlphaCutoff = cParams.GetLong("PNGAlphaColor", vbWhite)
            If newColorDepth = 24 Then newColorDepth = 32
        End If
        
        If (m_FIHandle <> 0) Then Plugin_FreeImage.ReleaseFreeImageObject m_FIHandle
        m_FIHandle = Plugin_FreeImage.GetFIDib_SpecificColorMode(workingDIB, newColorDepth, desiredAlphaMode, PDAS_ComplicatedAlpha, desiredAlphaCutoff, cParams.GetLong("PNGBackgroundColor", vbWhite), forceGrayscale, newPaletteSize, , True)
        
    End If
    
End Sub

Private Function ParamsEqual(ByVal param1 As String, ByVal param2 As String) As Boolean
    ParamsEqual = CBool(StrComp(param1, param2, vbTextCompare) = 0)
End Function

Private Sub UpdatePreview()

    If cmdBar.PreviewsAllowed And g_ImageFormats.FreeImageEnabled And sldColorCount.IsValid Then
        
        'Make sure the preview source is up-to-date
        If (m_FIHandle = 0) Then UpdatePreviewSource
        
        'Retrieve a PNG-saved version of the current preview image
        workingDIB.ResetDIB
        If Plugin_FreeImage.GetExportPreview(m_FIHandle, workingDIB, PDIF_PNG) Then
            FinalizeNonstandardPreview pdFxPreview, True
        End If
        
    End If
    
End Sub

Private Sub sldAlphaCutoff_Change()
    UpdatePreviewSource
    UpdatePreview
End Sub

Private Sub sldColorCount_Change()
    UpdatePreviewSource
    UpdatePreview
End Sub

Private Sub sltLossyPerformance_Change()
    UpdatePreviewButtonText
End Sub

Private Sub sltTargetQuality_Change()
    UpdatePreviewButtonText
End Sub

Private Sub ttlWebOptimize_Click(Index As Integer, ByVal newState As Boolean)
    ReflowWebOptimizePanel
End Sub

'The web optimization panel supports a couple different collapsible sections
Private Sub ReflowWebOptimizePanel()
    
    Dim offsetY As Long
    Dim isVisible As Boolean
    
    'Show/hide the lossy compression options
    isVisible = ttlWebOptimize(0).Value
    
    chkOptimizeLossy.Visible = isVisible
    sltTargetQuality.Visible = isVisible
    sltLossyPerformance.Visible = isVisible
    lblHint(2).Visible = isVisible
    lblHint(3).Visible = isVisible
    chkOptimizeDither.Visible = isVisible
    cmdUpdateLossyPreview.Visible = isVisible
    
    'Determine a vertical offset for the bottom part of the panel, contingent on the top panel being open or shut
    If isVisible Then
        offsetY = cmdUpdateLossyPreview.GetTop + cmdUpdateLossyPreview.GetHeight + FixDPI(16)
    Else
        offsetY = ttlWebOptimize(0).GetTop + ttlWebOptimize(0).GetHeight + FixDPI(16)
    End If
    
    'Show/hide the lossless compression options
    ttlWebOptimize(1).SetTop offsetY
    isVisible = ttlWebOptimize(1).Value
    
    If isVisible Then
        offsetY = ttlWebOptimize(1).GetTop + ttlWebOptimize(1).GetHeight + FixDPI(6)
        sltLosslessPerformance.SetTop offsetY
        offsetY = sltLosslessPerformance.GetTop + sltLosslessPerformance.GetHeight + FixDPI(3)
        lblHint(4).SetTop offsetY
        lblHint(5).SetTop offsetY
    End If
    
    sltLosslessPerformance.Visible = isVisible
    lblHint(4).Visible = isVisible
    lblHint(5).Visible = isVisible
    
End Sub
