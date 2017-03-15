VERSION 5.00
Begin VB.Form dialog_ExportTIFF 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " TIFF export options"
   ClientHeight    =   6900
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   13095
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
   ScaleHeight     =   460
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   873
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.pdContainer picContainer 
      Height          =   5175
      Index           =   0
      Left            =   5880
      TabIndex        =   5
      Top             =   840
      Width           =   7095
      _ExtentX        =   0
      _ExtentY        =   0
      Begin PhotoDemon.pdButtonStrip btsCompressionColor 
         Height          =   1095
         Left            =   0
         TabIndex        =   15
         Top             =   120
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   1931
         Caption         =   "color compression"
      End
      Begin PhotoDemon.pdColorSelector clsBackground 
         Height          =   975
         Left            =   0
         TabIndex        =   6
         Top             =   2520
         Width           =   7095
         _ExtentX        =   15690
         _ExtentY        =   1720
         Caption         =   "background color"
      End
      Begin PhotoDemon.pdButtonStrip btsCompressionMono 
         Height          =   1095
         Left            =   0
         TabIndex        =   16
         Top             =   1320
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   1931
         Caption         =   "monochrome compression"
      End
      Begin PhotoDemon.pdButtonStrip btsMultipage 
         Height          =   1095
         Left            =   0
         TabIndex        =   17
         Top             =   3600
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   1931
         Caption         =   "page format"
      End
   End
   Begin PhotoDemon.pdContainer picContainer 
      Height          =   5175
      Index           =   2
      Left            =   5880
      TabIndex        =   3
      Top             =   840
      Width           =   7095
      _ExtentX        =   0
      _ExtentY        =   0
      Begin PhotoDemon.pdMetadataExport mtdManager 
         Height          =   3255
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   5741
      End
   End
   Begin PhotoDemon.pdCommandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   6150
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   1323
   End
   Begin PhotoDemon.pdFxPreviewCtl pdFxPreview 
      Height          =   5895
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   10398
      ColorSelection  =   -1  'True
   End
   Begin PhotoDemon.pdButtonStrip btsCategory 
      Height          =   615
      Left            =   5880
      TabIndex        =   2
      Top             =   120
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   1085
   End
   Begin PhotoDemon.pdContainer picContainer 
      Height          =   5175
      Index           =   1
      Left            =   5880
      TabIndex        =   7
      Top             =   840
      Width           =   7095
      _ExtentX        =   0
      _ExtentY        =   0
      Begin PhotoDemon.pdSlider sldAlphaCutoff 
         Height          =   855
         Left            =   0
         TabIndex        =   8
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
         TabIndex        =   9
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
         TabIndex        =   10
         Top             =   2880
         Width           =   7095
         _ExtentX        =   15690
         _ExtentY        =   1931
         Caption         =   "transparency"
      End
      Begin PhotoDemon.pdButtonStrip btsColorModel 
         Height          =   1095
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Width           =   7095
         _ExtentX        =   15690
         _ExtentY        =   1931
         Caption         =   "color model"
      End
      Begin PhotoDemon.pdButtonStrip btsDepthColor 
         Height          =   1095
         Left            =   0
         TabIndex        =   12
         Top             =   1200
         Width           =   7095
         _ExtentX        =   15690
         _ExtentY        =   1931
         Caption         =   "depth"
      End
      Begin PhotoDemon.pdColorSelector clsAlphaColor 
         Height          =   975
         Left            =   0
         TabIndex        =   13
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
         TabIndex        =   14
         Top             =   1200
         Width           =   7095
         _ExtentX        =   15690
         _ExtentY        =   1931
         Caption         =   "depth"
      End
   End
End
Attribute VB_Name = "dialog_ExportTIFF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'TIFF export dialog
'Copyright 2012-2017 by Tanner Helland
'Created: 11/December/12
'Last updated: 29/April/16
'Last update: repurpose old color-depth dialog into a TIFF-specific one
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
    Me.Visible = False
End Sub

Private Sub cmdBar_OKClick()
    m_FormatParamString = GetExportParamString
    m_MetadataParamString = mtdManager.GetMetadataSettings
    m_UserDialogAnswer = vbOK
    Me.Visible = False
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub cmdBar_ResetClick()
    
    cmdBar.MarkPreviewStatus False
    
    'General panel settings
    btsCompressionColor.ListIndex = 0
    btsCompressionMono.ListIndex = 0
    btsMultipage.ListIndex = 0
    
    'Color and transparency settings
    btsColorModel.ListIndex = 0
    btsDepthColor.ListIndex = 1
    btsDepthGrayscale.ListIndex = 1
    btsAlpha.ListIndex = 0
    
    sldColorCount.Value = 256
    sldAlphaCutoff.Value = PD_DEFAULT_ALPHA_CUTOFF
    clsAlphaColor.Color = RGB(255, 0, 255)
    
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
    
    'Populate the category button strip
    btsCategory.AddItem "basic", 0
    btsCategory.AddItem "advanced", 1
    btsCategory.AddItem "metadata", 2
    btsCategory.ListIndex = 0
    
    'Basic options
    btsMultipage.AddItem "single page (composited image)", 0
    btsMultipage.AddItem "multipage (one page per layer)", 1
    btsMultipage.ListIndex = 0
    
    If Not (srcImage Is Nothing) Then
        btsMultipage.Visible = CBool(srcImage.GetNumOfLayers > 0)
    End If
    
    btsCompressionColor.AddItem "auto", 0
    btsCompressionColor.AddItem "LZW", 1
    btsCompressionColor.AddItem "ZIP", 2
    btsCompressionColor.AddItem "none", 3
    btsCompressionColor.ListIndex = 0
    
    btsCompressionMono.AddItem "auto", 0
    btsCompressionMono.AddItem "CCITT Fax 4", 1
    btsCompressionMono.AddItem "CCITT Fax 3", 2
    btsCompressionMono.AddItem "LZW", 3
    btsCompressionMono.AddItem "none", 4
    btsCompressionMono.ListIndex = 0
    
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
    
    'TIFFs also support a (ridiculous) amount of alpha settings
    btsAlpha.AddItem "auto", 0
    btsAlpha.AddItem "full", 1
    btsAlpha.AddItem "binary (by cut-off)", 2
    btsAlpha.AddItem "binary (by color)", 3
    btsAlpha.AddItem "none", 4
    
    sldAlphaCutoff.NotchValueCustom = PD_DEFAULT_ALPHA_CUTOFF
    
    'Prep a preview (if any)
    Set m_SrcImage = srcImage
    If Not (m_SrcImage Is Nothing) Then
        m_SrcImage.GetCompositedImage m_CompositedImage, True
        pdFxPreview.NotifyNonStandardSource m_CompositedImage.GetDIBWidth, m_CompositedImage.GetDIBHeight
    End If
    If (Not g_ImageFormats.FreeImageEnabled) Or (m_SrcImage Is Nothing) Then Interface.ShowDisabledPreviewImage pdFxPreview
    
    'Next, prepare various controls on the metadata panel
    mtdManager.SetParentImage m_SrcImage, PDIF_TIFF
    
    'Update the preview
    UpdatePreviewSource
    UpdatePreview
    
    'Apply translations and visual themes
    ApplyThemeAndTranslations Me
    
    'Many UI options are dynamically shown/hidden depending on other settings; make sure their initial state is correct
    UpdatePanelVisibility
    UpdateColorDepthVisibility
    UpdateTransparencyOptions
    
    'Display the dialog
    ShowPDDialog vbModal, Me, True

End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
    Plugin_FreeImage.ReleasePreviewCache m_FIHandle
End Sub

Private Function GetExportParamString() As String

    Dim cParams As pdParamXML
    Set cParams = New pdParamXML
    
    'Start with the standard TIFF settings
    Dim compressName As String
    
    Select Case btsCompressionColor.ListIndex
        'Auto
        Case 0
            compressName = "LZW"
        'LZW
        Case 1
            compressName = "LZW"
        'ZIP
        Case 2
            compressName = "ZIP"
        'NONE
        Case 3
            compressName = "none"
        Case Else
            compressName = "LZW"
    End Select
    
    cParams.AddParam "TIFFCompressionColor", compressName
    
    Select Case btsCompressionColor.ListIndex
        'Auto
        Case 0
            compressName = "Fax4"
        'CCITT Fax 4
        Case 1
            compressName = "Fax4"
        'CCITT Fax 3
        Case 2
            compressName = "Fax3"
        'LZW
        Case 3
            compressName = "LZW"
        'NONE
        Case 4
            compressName = "none"
        Case Else
            compressName = "Fax4"
    End Select
    
    cParams.AddParam "TIFFCompressionMono", compressName
    
    cParams.AddParam "TIFFBackgroundColor", clsBackground.Color
    If (btsMultipage.ListIndex <> 0) Then
        cParams.AddParam "TIFFMultipage", True
    Else
        cParams.AddParam "TIFFMultipage", False
    End If
        
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
    cParams.AddParam "TIFFColorModel", outputColorModel
        
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
    
    If (Len(outputColorDepth) <> 0) Then cParams.AddParam "TIFFBitDepth", outputColorDepth
    If (Len(outputPaletteSize) <> 0) Then cParams.AddParam "TIFFPaletteSize", outputPaletteSize
    
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
    
    cParams.AddParam "TIFFAlphaModel", outputAlphaModel
    If sldAlphaCutoff.IsValid Then cParams.AddParam "TIFFAlphaCutoff", sldAlphaCutoff.Value Else cParams.AddParam "TIFFAlphaCutoff", PD_DEFAULT_ALPHA_CUTOFF
    cParams.AddParam "TIFFAlphaColor", clsAlphaColor.Color
    
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
    
    If (Not (m_CompositedImage Is Nothing)) Then
        
        'Because the user can change the preview viewport, we can't guarantee that the preview region hasn't changed
        ' since the last preview.  Prep a new preview now.
        Dim tmpSafeArray As SAFEARRAY2D
        EffectPrep.PreviewNonStandardImage tmpSafeArray, m_CompositedImage, pdFxPreview, True
        
        'To reduce the chance of bugs, we use the same parameter parsing technique as the core TIFF encoder
        Dim cParams As pdParamXML
        Set cParams = New pdParamXML
        cParams.SetParamString GetExportParamString()
        
        'Color and grayscale modes require different processing, so start there
        Dim forceGrayscale As Boolean
        forceGrayscale = ParamsEqual(cParams.GetString("TIFFColorModel", "Auto"), "Gray")
        
        'For 8-bit modes, grab a palette size.  (This parameter will be ignored in other color modes.)
        Dim newPaletteSize As Long
        newPaletteSize = cParams.GetLong("TIFFPaletteSize", 256)
        
        Dim newColorDepth As Long
        
        If ParamsEqual(cParams.GetString("TIFFColorModel", "Auto"), "Auto") Then
            newColorDepth = 32
        Else
            
            'HDR modes do not need to be previewed, so we forcibly downsample them here
            If forceGrayscale Then
                newColorDepth = cParams.GetLong("TIFFBitDepth", 8)
                If newColorDepth > 8 Then newColorDepth = 8
                If newColorDepth = 1 Then
                    newPaletteSize = 2
                    newColorDepth = 8
                End If
            Else
                newColorDepth = cParams.GetLong("TIFFBitDepth", 24)
                If newColorDepth = 48 Then newColorDepth = 24
                If newColorDepth = 64 Then newColorDepth = 32
            End If
        
        End If
        
        'Next comes transparency, which is somewhat messy because we offer alpha behavior identical to the PNG plugin
        Dim desiredAlphaMode As PD_ALPHA_STATUS, desiredAlphaCutoff As Long
        
        If ParamsEqual(cParams.GetString("TIFFAlphaModel", "Auto"), "Auto") Or ParamsEqual(cParams.GetString("TIFFAlphaModel", "Auto"), "Full") Then
            desiredAlphaMode = PDAS_ComplicatedAlpha
            If newColorDepth = 24 Then newColorDepth = 32
        ElseIf ParamsEqual(cParams.GetString("TIFFAlphaModel", "Auto"), "None") Then
            desiredAlphaMode = PDAS_NoAlpha
            If newColorDepth = 32 Then newColorDepth = 24
            desiredAlphaCutoff = 0
        ElseIf ParamsEqual(cParams.GetString("TIFFAlphaModel", "Auto"), "ByCutoff") Then
            desiredAlphaMode = PDAS_BinaryAlpha
            desiredAlphaCutoff = cParams.GetLong("TIFFAlphaCutoff", PD_DEFAULT_ALPHA_CUTOFF)
            If newColorDepth = 24 Then newColorDepth = 32
        ElseIf ParamsEqual(cParams.GetString("TIFFAlphaModel", "Auto"), "ByColor") Then
            desiredAlphaMode = PDAS_NewAlphaFromColor
            desiredAlphaCutoff = cParams.GetLong("TIFFAlphaColor", vbWhite)
            If newColorDepth = 24 Then newColorDepth = 32
        End If
        
        If (m_FIHandle <> 0) Then Plugin_FreeImage.ReleaseFreeImageObject m_FIHandle
        m_FIHandle = Plugin_FreeImage.GetFIDib_SpecificColorMode(workingDIB, newColorDepth, desiredAlphaMode, PDAS_ComplicatedAlpha, desiredAlphaCutoff, cParams.GetLong("TIFFBackgroundColor", vbWhite), forceGrayscale, newPaletteSize, , True)
        
    End If
    
End Sub

Private Function ParamsEqual(ByVal param1 As String, ByVal param2 As String) As Boolean
    ParamsEqual = CBool(StrComp(param1, param2, vbTextCompare) = 0)
End Function

Private Sub UpdatePreview()

    If (cmdBar.PreviewsAllowed And g_ImageFormats.FreeImageEnabled And sldColorCount.IsValid And (Not m_SrcImage Is Nothing)) Then
        
        'Make sure the preview source is up-to-date
        If (m_FIHandle = 0) Then UpdatePreviewSource
        
        'Retrieve a TIFF-saved version of the current preview image
        workingDIB.ResetDIB
        If Plugin_FreeImage.GetExportPreview(m_FIHandle, workingDIB, PDIF_TIFF, FISO_TIFF_NONE) Then
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
