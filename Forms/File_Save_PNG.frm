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
      Height          =   6630
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   11695
      ColorSelection  =   -1  'True
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
         Height          =   5535
         Index           =   0
         Left            =   0
         ScaleHeight     =   369
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   473
         TabIndex        =   13
         Top             =   720
         Width           =   7095
         Begin PhotoDemon.pdCheckBox chkInterlace 
            Height          =   375
            Left            =   120
            TabIndex        =   15
            Top             =   1200
            Width           =   6975
            _ExtentX        =   12303
            _ExtentY        =   661
            Caption         =   "use interlacing"
            Value           =   0
         End
         Begin PhotoDemon.pdSlider sldCompression 
            Height          =   975
            Left            =   0
            TabIndex        =   14
            Top             =   240
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   1720
            Caption         =   "compression level"
            Max             =   9
            Value           =   9
            GradientColorRight=   1703935
            NotchPosition   =   2
            NotchValueCustom=   9
         End
         Begin PhotoDemon.pdColorSelector clsBackground 
            Height          =   975
            Left            =   0
            TabIndex        =   17
            Top             =   1800
            Width           =   7095
            _ExtentX        =   15690
            _ExtentY        =   1720
            Caption         =   "background color"
         End
         Begin PhotoDemon.pdCheckBox chkEmbedBackground 
            Height          =   375
            Left            =   120
            TabIndex        =   18
            Top             =   2880
            Width           =   6975
            _ExtentX        =   12303
            _ExtentY        =   661
            Caption         =   "embed background color in file"
            Value           =   0
         End
      End
      Begin VB.PictureBox picContainer 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         HasDC           =   0   'False
         Height          =   5535
         Index           =   2
         Left            =   0
         ScaleHeight     =   369
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
End Sub

Private Sub btsDepthGrayscale_Click(ByVal buttonIndex As Long)
    UpdateColorDepthOptions
End Sub

Private Sub btsMasterType_Click(ByVal buttonIndex As Long)
    UpdateMasterPanelVisibility
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
    m_MetadataParamString = mtdManager.GetMetadataSettings
    m_UserDialogAnswer = vbOK
    Me.Hide
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub cmdBar_ResetClick()
    
    'General panel settings
    sldCompression.Value = sldCompression.Max
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
    
    'Metadata settings
    mtdManager.Reset
    
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
    
    'Color model and color depth are closely related; populate all button strips, then show/hide the relevant pairings
    btsColorModel.AddItem "auto", 0
    btsColorModel.AddItem "color", 1
    btsColorModel.AddItem "grayscale", 2
    btsColorModel.ListIndex = 0
    
    btsDepthColor.AddItem "48-bpp (HDR)", 0
    btsDepthColor.AddItem "24-bpp (standard)", 1
    btsDepthColor.AddItem "8-bpp (indexed)", 2
    btsDepthColor.ListIndex = 1
    
    btsDepthGrayscale.AddItem "16-bpp (HDR)", 0
    btsDepthGrayscale.AddItem "8-bpp (standard)", 1
    btsDepthGrayscale.AddItem "4-bpp", 2
    btsDepthGrayscale.AddItem "1-bpp (monochrome)", 3
    btsDepthGrayscale.ListIndex = 1
    
    UpdateColorDepthVisibility
    
    'PNGs also support a (ridiculous) amount of alpha settings
    btsAlpha.AddItem "auto", 0
    btsAlpha.AddItem "full", 1
    btsAlpha.AddItem "binary (by cut-off)", 2
    btsAlpha.AddItem "binary (by color)", 3
    btsAlpha.AddItem "none", 4
    
    sldAlphaCutoff.NotchValueCustom = DEFAULT_ALPHA_CUTOFF
    
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
    
    'Display the dialog
    ShowPDDialog vbModal, Me, True

End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
    Plugin_FreeImage.ReleasePreviewCache
End Sub

Private Function GetExportParamString() As String

    Dim cParams As pdParamXML
    Set cParams = New pdParamXML
    
    'Convert the color depth option buttons into a usable numeric value
    Dim outputColorMode As String
    
    Select Case btsColorModel.ListIndex
        Case 0
            outputColorMode = "Auto"
        Case 1
            outputColorMode = "Color"
        Case 2
            outputColorMode = "Gray"
    End Select
    
    cParams.AddParam "GIFColorMode", outputColorMode
    
    Dim outputAlphaMode As String
    Select Case btsAlpha.ListIndex
        Case 0
            outputAlphaMode = "Auto"
        Case 1
            outputAlphaMode = "None"
        Case 2
            outputAlphaMode = "ByCutoff"
        Case 3
            outputAlphaMode = "ByColor"
    End Select
    
    cParams.AddParam "GIFAlphaMode", outputAlphaMode
    
    'If "auto" mode is selected, we currently enforce a hard-coded cut-off value.  There may be a better way to do this,
    ' but I'm not currently aware of it!
    Dim outputAlphaCutoff As Long
    If (btsAlpha.ListIndex = 0) Or (Not sldAlphaCutoff.IsValid) Then outputAlphaCutoff = DEFAULT_ALPHA_CUTOFF Else outputAlphaCutoff = sldAlphaCutoff.Value
    cParams.AddParam "GIFAlphaCutoff", outputAlphaCutoff
    
    Dim colorCount As Long
    If (btsColorModel.ListIndex <> 0) Then
        If sldColorCount.IsValid Then colorCount = sldColorCount.Value Else colorCount = 256
    Else
        colorCount = 256
    End If
    cParams.AddParam "GIFColorCount", colorCount
    cParams.AddParam "GIFBackgroundColor", clsBackground.Color
    cParams.AddParam "GIFAlphaColor", clsAlphaColor.Color
    
    GetExportParamString = cParams.GetParamString
    
End Function

Private Sub pdFxPreview_ColorSelected()
    clsAlphaColor.Color = pdFxPreview.SelectedColor
End Sub

Private Sub pdFxPreview_ViewportChanged()
    UpdatePreviewSource
    UpdatePreview
End Sub

'When a parameter changes that requires a new source DIB for the preview (e.g. changing the background composite color),
' call this function to generate a new preview DIB.  Note that you *do not* need to call this function for format-specific
' changes (like quality, subsampling, etc).
Private Sub UpdatePreviewSource()
    If Not (m_CompositedImage Is Nothing) Then
        
        'Because the user can change the preview viewport, we can't guarantee that the preview region hasn't changed
        ' since the last preview.  Prep a new preview now.
        Dim tmpSafeArray As SAFEARRAY2D
        FastDrawing.PreviewNonStandardImage tmpSafeArray, m_CompositedImage, pdFxPreview, True
        
        'Convert the DIB to a FreeImage-compatible handle, at a color-depth that matches the current settings.
        ' (Note that one way or another, we'll always be converting the image to an 8-bpp mode.)
        Dim forceGrayscale As Boolean
        forceGrayscale = CBool(btsColorModel.ListIndex = 2)
        
        Dim paletteCount As Long
        If (btsColorModel.ListIndex = 0) Then
            paletteCount = 256
        Else
            If sldColorCount.IsValid Then paletteCount = sldColorCount.Value Else paletteCount = 256
        End If
        
        Dim desiredAlphaMode As PD_ALPHA_STATUS, desiredAlphaCutoff As Long
        If btsAlpha.ListIndex = 0 Then
            desiredAlphaMode = PDAS_BinaryAlpha       'Auto
            desiredAlphaCutoff = DEFAULT_ALPHA_CUTOFF
        ElseIf btsAlpha.ListIndex = 1 Then
            desiredAlphaMode = PDAS_NoAlpha           'None
            desiredAlphaCutoff = 0
        ElseIf btsAlpha.ListIndex = 2 Then
            desiredAlphaMode = PDAS_BinaryAlpha       'By cut-off
            If sldAlphaCutoff.IsValid Then desiredAlphaCutoff = sldAlphaCutoff.Value Else desiredAlphaCutoff = 96
        Else
            desiredAlphaMode = PDAS_NewAlphaFromColor 'By color
            desiredAlphaCutoff = clsAlphaColor.Color
        End If
        
        m_FIHandle = Plugin_FreeImage.GetFIDib_SpecificColorMode(workingDIB, 8, desiredAlphaMode, PDAS_ComplicatedAlpha, desiredAlphaCutoff, clsBackground.Color, forceGrayscale, paletteCount)
        
    End If
    
End Sub

Private Sub UpdatePreview()

    If cmdBar.PreviewsAllowed And g_ImageFormats.FreeImageEnabled And sldColorCount.IsValid Then
        
        'Make sure the preview source is up-to-date
        If (m_FIHandle = 0) Then UpdatePreviewSource
        
        'Retrieve a BMP-saved version of the current preview image
        workingDIB.ResetDIB
        If Plugin_FreeImage.GetExportPreview(m_FIHandle, workingDIB, PDIF_GIF) Then
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
