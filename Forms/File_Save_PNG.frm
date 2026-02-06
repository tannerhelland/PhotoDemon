VERSION 5.00
Begin VB.Form dialog_ExportPNG 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   7500
   ClientLeft      =   120
   ClientTop       =   465
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
   Icon            =   "File_Save_PNG.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   500
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   874
   Begin PhotoDemon.pdCommandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   6750
      Width           =   13110
      _ExtentX        =   23125
      _ExtentY        =   1323
   End
   Begin PhotoDemon.pdFxPreviewCtl pdFxPreview 
      Height          =   6495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   11456
   End
   Begin PhotoDemon.pdContainer picCategory 
      Height          =   6495
      Index           =   0
      Left            =   5880
      Top             =   120
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   11245
      Begin PhotoDemon.pdTitle ttlStandard 
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   0
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   661
         Caption         =   "basic settings"
         FontBold        =   -1  'True
         FontSize        =   12
      End
      Begin PhotoDemon.pdTitle ttlStandard 
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   661
         Caption         =   "advanced settings"
         FontBold        =   -1  'True
         FontSize        =   12
         Value           =   0   'False
      End
      Begin PhotoDemon.pdTitle ttlStandard 
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   661
         Caption         =   "metadata settings"
         FontBold        =   -1  'True
         FontSize        =   12
         Value           =   0   'False
      End
      Begin PhotoDemon.pdContainer picContainer 
         Height          =   2535
         Index           =   0
         Left            =   120
         Top             =   1200
         Visible         =   0   'False
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   4471
         Begin PhotoDemon.pdDropDown cboOptimize 
            Height          =   855
            Left            =   360
            TabIndex        =   6
            Top             =   1080
            Width           =   6615
            _ExtentX        =   11668
            _ExtentY        =   1508
            Caption         =   "compression optimization"
         End
         Begin PhotoDemon.pdLabel lblHint 
            Height          =   255
            Index           =   0
            Left            =   480
            Top             =   720
            Width           =   2340
            _ExtentX        =   4128
            _ExtentY        =   450
            Caption         =   "fast, larger file"
            FontItalic      =   -1  'True
            FontSize        =   9
         End
         Begin PhotoDemon.pdSlider sldCompression 
            Height          =   735
            Left            =   360
            TabIndex        =   7
            Top             =   0
            Width           =   6615
            _ExtentX        =   11668
            _ExtentY        =   1296
            Caption         =   "compression level"
            Max             =   12
            Value           =   9
            GradientColorRight=   1703935
            NotchPosition   =   2
            NotchValueCustom=   9
         End
         Begin PhotoDemon.pdColorSelector clsBackground 
            Height          =   375
            Left            =   5400
            TabIndex        =   8
            Top             =   2040
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            FontSize        =   10
            ShowMainWindowColor=   0   'False
         End
         Begin PhotoDemon.pdCheckBox chkEmbedBackground 
            Height          =   375
            Left            =   360
            TabIndex        =   9
            Top             =   2070
            Width           =   4935
            _ExtentX        =   8705
            _ExtentY        =   661
            Caption         =   "embed background color (bKGD chunk)"
            Value           =   0   'False
         End
         Begin PhotoDemon.pdLabel lblHint 
            Height          =   255
            Index           =   1
            Left            =   2880
            Top             =   720
            Width           =   2580
            _ExtentX        =   4551
            _ExtentY        =   450
            Alignment       =   1
            Caption         =   "slow, smaller file"
            FontItalic      =   -1  'True
            FontSize        =   9
         End
      End
      Begin PhotoDemon.pdContainer picContainer 
         Height          =   5175
         Index           =   1
         Left            =   120
         Top             =   1200
         Visible         =   0   'False
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   9128
         Begin PhotoDemon.pdColorDepth clrDepth 
            Height          =   5055
            Left            =   360
            TabIndex        =   10
            Top             =   0
            Width           =   6615
            _ExtentX        =   11668
            _ExtentY        =   8916
         End
      End
      Begin PhotoDemon.pdContainer picContainer 
         Height          =   3255
         Index           =   2
         Left            =   120
         Top             =   1200
         Visible         =   0   'False
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   5741
         Begin PhotoDemon.pdMetadataExport mtdManager 
            Height          =   3255
            Left            =   360
            TabIndex        =   5
            Top             =   0
            Width           =   6495
            _ExtentX        =   11456
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
'Copyright 2012-2026 by Tanner Helland
'Created: 11/December/12
'Last updated: 29/October/21
'Last update: remove "web-optimized PNG" panel; instead, we're gonna do a full-blown Save for Web tool
'
'PhotoDemon ships with a custom-built PNG encoder capable of better performance and compression than
' standard libraries like libPNG.  This means that we can expose a lot of extra options for pro users,
' without having to hack up an external library to support all those options.
'
'Old versions of PD divided this dialog into two panels: "standard" PNGs and "web-optimized" PNGs.
' This has since been retired in favor of a full Save for Web tool.
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

'OK or CANCEL result
Private m_UserDialogAnswer As VbMsgBoxResult

'Final format-specific XML packet, with all format-specific settings defined as tag+value pairs
Private m_FormatParamString As String

'Final metadata XML packet, with all metadata settings defined as tag+value pairs.  Currently unused as ExifTool
' cannot write any BMP-specific data.
Private m_MetadataParamString As String

'Used to avoid recursive setting changes
Private m_ActiveTitleBar As Long, m_PanelChangesActive As Boolean

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

Private Sub chkEmbedBackground_Click()
    UpdateBkgdColorVisibility
End Sub

Private Sub UpdateBkgdColorVisibility()
    clsBackground.Visible = chkEmbedBackground.Value
End Sub

Private Sub clrDepth_Change()
    UpdatePreview
End Sub

Private Sub clrDepth_ColorSelectionRequired(ByVal selectState As Boolean)
    pdFxPreview.AllowColorSelection = selectState
End Sub

Private Sub clrDepth_SizeChanged()
    clrDepth.SyncToIdealSize
    picContainer(1).SetHeight clrDepth.GetIdealSize
    ttlStandard(2).SetTop picContainer(1).GetTop + picContainer(1).GetHeight + FixDPI(8)
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

Private Sub cmdBar_RequestPreviewUpdate()
    
    UpdatePreview
    
    'Reflow titlebars too, to ensure panel layout is correct
    UpdateStandardPanelVisibility
    
End Sub

Private Sub cmdBar_ResetClick()
    
    cmdBar.SetPreviewStatus False
    
    'General panel settings
    sldCompression.Value = sldCompression.NotchValueCustom
    
    If (Not m_SrcImage Is Nothing) Then
        If m_SrcImage.ImgStorage.DoesKeyExist("png-background-color") Then
            clsBackground.Color = m_SrcImage.ImgStorage.GetEntry_Long("png-background-color")
            chkEmbedBackground.Value = True
        Else
            clsBackground.Color = vbWhite
            chkEmbedBackground.Value = False
        End If
    Else
        clsBackground.Color = vbWhite
        chkEmbedBackground.Value = False
    End If
    
    'Metadata settings
    mtdManager.Reset
    
    cmdBar.SetPreviewStatus True
    UpdatePreview
    
End Sub

'The ShowDialog routine presents the user with this form.
Public Sub ShowDialog(Optional ByRef srcImage As pdImage = Nothing)

    'Provide a default answer of "cancel" (in the event that the user clicks the "x" button in the top-right)
    m_UserDialogAnswer = vbCancel
    
    Message "Waiting for user to specify export options... "
    
    'Standard settings are accessed via pdTitle controls.  Because the panels are so large, only one panel
    ' is allowed open at a time.
    Dim i As Long
    For i = picContainer.lBound To picContainer.UBound
        picContainer(i).SetLeft 0
    Next i
    
    'If the file being saved was originally a PNG, notify the color-depth handler that we want
    ' to expose a "use original file settings" option
    If (Not srcImage Is Nothing) Then clrDepth.SetOriginalSettingsAvailable (srcImage.GetOriginalFileFormat = PDIF_PNG)
    
    clrDepth.SyncToIdealSize
    ttlStandard(0).Value = True
    m_ActiveTitleBar = 0
    UpdateStandardTitlebars
    
    'Populate filter strategy options
    cboOptimize.Clear
    cboOptimize.SetAutomaticRedraws False
    cboOptimize.AddItem "automatic", 0
    cboOptimize.AddItem "optimize: fast filters", 1
    cboOptimize.AddItem "optimize: all filters", 2
    cboOptimize.AddItem "single filter: none", 3
    cboOptimize.AddItem "single filter: sub", 4
    cboOptimize.AddItem "single filter: up", 5
    cboOptimize.AddItem "single filter: average", 6
    cboOptimize.AddItem "single filter: paeth", 7
    cboOptimize.AssignTooltip "PNG files support different compression strategies (called ""filters"").  Smart filter selection produces better compression.  Use the automatic setting to have PhotoDemon test multiple strategies, and automatically select the one that produces the best compression."
    cboOptimize.ListIndex = 0
    cboOptimize.SetAutomaticRedraws True, True
    
    'Prep a preview (if any)
    Set m_SrcImage = srcImage
    If (Not m_SrcImage Is Nothing) Then
        m_SrcImage.GetCompositedImage m_CompositedImage, True
        pdFxPreview.NotifyNonStandardSource m_CompositedImage.GetDIBWidth, m_CompositedImage.GetDIBHeight
    End If
    If (m_SrcImage Is Nothing) Then Interface.ShowDisabledPreviewImage pdFxPreview
    
    'Next, prepare various controls on the metadata panel
    mtdManager.SetParentImage m_SrcImage, PDIF_PNG
    
    'If the source image was a PNG, and it also contained a background color, retrieve and set the matching color now
    If (Not m_SrcImage Is Nothing) Then
        If m_SrcImage.ImgStorage.DoesKeyExist("png-background-color") Then
            clsBackground.Color = m_SrcImage.ImgStorage.GetEntry_Long("png-background-color")
            chkEmbedBackground.Value = True
        End If
    End If
    
    UpdateBkgdColorVisibility
    
    'Update the preview
    UpdatePreview
    
    'Apply translations and visual themes
    ApplyThemeAndTranslations Me, True, True
    Interface.SetFormCaptionW Me, g_Language.TranslateMessage("%1 options", "PNG")
    
    'Many UI options are dynamically shown/hidden depending on other settings; make sure their initial state is correct
    UpdateStandardPanelVisibility
    
    'Display the dialog
    ShowPDDialog vbModal, Me, True

End Sub

Private Sub Form_Activate()
    clrDepth.SyncToIdealSize
    picContainer(1).SetHeight clrDepth.GetIdealSize
    ttlStandard(2).SetTop picContainer(1).GetTop + picContainer(1).GetHeight + FixDPI(8)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Function GetExportParamString() As String

    Dim cParams As pdSerialize
    Set cParams = New pdSerialize

    'Start with the standard PNG settings, which are consistent across all standard PNG types
    If sldCompression.IsValid Then cParams.AddParam "png-compression-level", sldCompression.Value Else cParams.AddParam "png-compression-level", sldCompression.NotchValueCustom
    cParams.AddParam "png-background-color", clsBackground.Color
    cParams.AddParam "png-create-bkgd", chkEmbedBackground.Value
    cParams.AddParam "png-filter-strategy", cboOptimize.ListIndex
    
    'Next come all the messy color-depth possibilities
    cParams.AddParam "png-color-depth", clrDepth.GetAllSettings
    
    GetExportParamString = cParams.GetParamString()
    
End Function

Private Sub pdFxPreview_ColorSelected()
    clrDepth.NotifyNewAlphaColor pdFxPreview.SelectedColor
End Sub

Private Sub pdFxPreview_ViewportChanged()
    UpdatePreview
End Sub

Private Function ParamsEqual(ByRef param1 As String, ByRef param2 As String) As Boolean
    ParamsEqual = Strings.StringsEqual(param1, param2, True)
End Function

Private Sub UpdatePreview()

    If (cmdBar.PreviewsAllowed And clrDepth.IsValid) And (Not m_SrcImage Is Nothing) And (Not m_CompositedImage Is Nothing) Then
        
        'Because the user can change the preview viewport, we can't guarantee that the preview region
        ' hasn't changed since the last preview.  Prep a new preview base image now.
        Dim tmpSafeArray As SafeArray2D
        EffectPrep.PreviewNonStandardImage tmpSafeArray, m_CompositedImage, pdFxPreview, True
        
        'To reduce the chance of bugs, we use the same parameter parsing technique as the core PNG encoder
        Dim cParams As pdSerialize
        Set cParams = New pdSerialize
        cParams.SetParamString GetExportParamString()
        
        Dim bkgdColor As Long
        bkgdColor = cParams.GetLong("png-background-color", vbWhite)
        
        'The color-depth-specific options are embedded as a single option, so extract them into their
        ' own parser.
        Dim cParamsDepth As pdSerialize
        Set cParamsDepth = New pdSerialize
        cParamsDepth.SetParamString cParams.GetString("png-color-depth", vbNullString)
        
        'Retrieve color and alpha model for this preview; everything else extends from these
        Dim outputColorModel As String, outputAlphaModel As String
        outputColorModel = cParamsDepth.GetString("cd-color-model", "auto", True)
        outputAlphaModel = cParamsDepth.GetString("cd-alpha-model", "auto", True)
    
        'Before doing anything else, figure out how to handle alpha.
        Dim previewAlphaMode As PD_ALPHA_STATUS
        If Strings.StringsEqual(outputAlphaModel, "full", True) Then
            previewAlphaMode = PDAS_ComplicatedAlpha
        ElseIf Strings.StringsEqual(outputAlphaModel, "none", True) Then
            previewAlphaMode = PDAS_NoAlpha
        ElseIf Strings.StringsEqual(outputAlphaModel, "by-cutoff", True) Then
            previewAlphaMode = PDAS_BinaryAlpha
        ElseIf Strings.StringsEqual(outputAlphaModel, "by-color", True) Then
            previewAlphaMode = PDAS_NewAlphaFromColor
        Else
            previewAlphaMode = PDAS_ComplicatedAlpha
        End If
        
        Dim trnsTable() As Byte
        
        Dim outputAlphaCutoff As Long, outputAlphaColor As Long
        outputAlphaCutoff = cParamsDepth.GetLong("cd-alpha-cutoff", PD_DEFAULT_ALPHA_CUTOFF)
        outputAlphaColor = cParamsDepth.GetLong("cd-alpha-color", vbMagenta)
        
        'If the caller specified "original file settings", override any settings we have
        ' calculated with the file's *original* settings.
        Dim useOrigMode As Boolean, origColorType As PD_PNGColorType
        useOrigMode = Strings.StringsEqual(cParamsDepth.GetString("cd-color-model", "auto", True), "original", True)
        
        If useOrigMode Then
            
            'First we want to determine alpha status; colors will be dealt with later
            origColorType = m_SrcImage.ImgStorage.GetEntry_Long("png-color-type", png_AutoColorType)
            
            'For PNGs with full alpha channels, we want to enable full alpha channel output
            If (origColorType = png_GreyscaleAlpha) Or (origColorType = png_TruecolorAlpha) Then
                previewAlphaMode = PDAS_ComplicatedAlpha
            Else
                
                'Override the background color (if any) with the one from the original file
                If m_SrcImage.ImgStorage.DoesKeyExist("png-background-color") Then bkgdColor = m_SrcImage.ImgStorage.GetEntry_Long("png-background-color")
                
                'If the file used some other form of transparency, assume binary transparency here
                If m_SrcImage.GetOriginalAlpha Then
                    previewAlphaMode = PDAS_BinaryAlpha
                    outputAlphaCutoff = 127
                Else
                    previewAlphaMode = PDAS_NoAlpha
                End If
                
            End If
            
        Else
            If (previewAlphaMode <> PDAS_ComplicatedAlpha) Then bkgdColor = cParamsDepth.GetLong("cd-matte-color", bkgdColor)
        End If
        
        'If the caller wants alpha removed, do so now.
        If (previewAlphaMode = PDAS_NoAlpha) Then
            workingDIB.CompositeBackgroundColor Colors.ExtractRed(bkgdColor), Colors.ExtractGreen(bkgdColor), Colors.ExtractBlue(bkgdColor)
            
        ElseIf (previewAlphaMode = PDAS_BinaryAlpha) Then
            DIBs.ApplyAlphaCutoff_Ex workingDIB, trnsTable, outputAlphaCutoff
            DIBs.ApplyBinaryTransparencyTable workingDIB, trnsTable, bkgdColor
        
        ElseIf (previewAlphaMode = PDAS_NewAlphaFromColor) Then
            DIBs.MakeColorTransparent_Ex workingDIB, trnsTable, outputAlphaColor
            DIBs.ApplyBinaryTransparencyTable workingDIB, trnsTable, bkgdColor
        
        'Other alpha modes require no changes on our part
        Else
        
        End If
        
        'With alpha handled successfully, we now need to handle grayscale and/or palette requirements
        If Strings.StringsNotEqual(outputColorModel, "auto", True) Then
            
            Dim forceGrayscale As Boolean, forceIndexed As Boolean, newPaletteSize As Long
            Dim newPalette() As RGBQuad
            
            If useOrigMode Then
            
                forceGrayscale = (origColorType = png_Greyscale) Or (origColorType = png_GreyscaleAlpha)
                forceIndexed = (origColorType = png_Indexed)
                
                Dim tmpPalette As pdPalette
                If m_SrcImage.HasOriginalPalette Then
                    m_SrcImage.GetOriginalPalette tmpPalette
                    If (Not tmpPalette Is Nothing) Then tmpPalette.CopyPaletteToArray newPalette
                End If
                
            Else
                forceGrayscale = ParamsEqual(cParamsDepth.GetString("cd-color-model", "auto"), "gray")
                forceIndexed = ParamsEqual(cParamsDepth.GetString("cd-color-depth", "color-standard"), "color-indexed")
                newPaletteSize = cParamsDepth.GetLong("cd-palette-size", 256)
                If ParamsEqual(cParamsDepth.GetString("cd-gray-depth", "auto"), "gray-monochrome") Then newPaletteSize = 2
            End If
            
            If forceGrayscale Then
                DIBs.MakeDIBGrayscale workingDIB, newPaletteSize
                
            ElseIf forceIndexed Then
                If (Not useOrigMode) Then Palettes.GetNeuquantPalette_RGBA workingDIB, newPalette, newPaletteSize
                Palettes.ApplyPaletteToImage_IncAlpha_KDTree workingDIB, newPalette, True
            End If
            
        End If
        
        EffectPrep.FinalizeNonstandardPreview pdFxPreview, True
        
    End If
    
End Sub

Private Sub ttlStandard_Click(Index As Integer, ByVal newState As Boolean)
    
    If newState Then m_ActiveTitleBar = Index
    picContainer(Index).Visible = newState
    
    If (Not m_PanelChangesActive) Then
        If newState Then UpdateStandardTitlebars Else UpdateStandardPanelVisibility
    End If
    
End Sub

Private Sub UpdateStandardTitlebars()
    
    m_PanelChangesActive = True
    
    '"Turn off" all titlebars except the selected one, and hide all panels except the selected one
    Dim i As Long
    For i = ttlStandard.lBound To ttlStandard.UBound
        ttlStandard(i).Value = (i = m_ActiveTitleBar)
        picContainer(i).Visible = ttlStandard(i).Value
    Next i
    
    'Because window visibility changes involve a number of window messages, let the message pump catch up.
    ' (We need window visibility finalized, because we need to query things like window size in order to
    '  reflow the current dialog layout.)
    DoEvents
    UpdateStandardPanelVisibility
    
    m_PanelChangesActive = False
    
End Sub

Private Sub UpdateStandardPanelVisibility()
    
    'Reflow the interface to match
    Dim yPos As Long, yPadding As Long
    yPos = 0
    yPadding = FixDPI(8)
    
    Dim i As Long
    For i = ttlStandard.lBound To ttlStandard.UBound
    
        ttlStandard(i).SetTop yPos
        yPos = yPos + ttlStandard(i).GetHeight + yPadding
        
        'The "advanced settings" panel uses a specialized custom control whose height may vary at run-time
        If (i = 1) Then
            clrDepth.SyncToIdealSize
            picContainer(i).SetHeight clrDepth.GetIdealSize
        End If
        
        If ttlStandard(i).Value Then
            picContainer(i).SetTop yPos
            yPos = yPos + picContainer(i).GetHeight + yPadding
        End If
        
    Next i
    
End Sub
