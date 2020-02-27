VERSION 5.00
Begin VB.Form dialog_ExportPNG 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " PNG export options"
   ClientHeight    =   8595
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
   ScaleHeight     =   573
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   874
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.pdButtonStrip btsMasterType 
      Height          =   1095
      Left            =   5880
      TabIndex        =   2
      Top             =   120
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   1931
      Caption         =   "PNG type"
      FontSize        =   12
   End
   Begin PhotoDemon.pdCommandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   7845
      Width           =   13110
      _ExtentX        =   23125
      _ExtentY        =   1323
   End
   Begin PhotoDemon.pdFxPreviewCtl pdFxPreview 
      Height          =   7575
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   13361
   End
   Begin PhotoDemon.pdContainer picCategory 
      Height          =   6495
      Index           =   0
      Left            =   5880
      TabIndex        =   3
      Top             =   1320
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   11245
      Begin PhotoDemon.pdTitle ttlStandard 
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   20
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
         TabIndex        =   21
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
         TabIndex        =   22
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
         TabIndex        =   7
         Top             =   1200
         Visible         =   0   'False
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   4471
         Begin PhotoDemon.pdDropDown cboOptimize 
            Height          =   855
            Left            =   360
            TabIndex        =   24
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
            TabIndex        =   8
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
            TabIndex        =   9
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
            TabIndex        =   10
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
         TabIndex        =   4
         Top             =   1200
         Visible         =   0   'False
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   9128
         Begin PhotoDemon.pdColorDepth clrDepth 
            Height          =   5055
            Left            =   360
            TabIndex        =   23
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
         TabIndex        =   5
         Top             =   1200
         Visible         =   0   'False
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   5741
         Begin PhotoDemon.pdMetadataExport mtdManager 
            Height          =   3255
            Left            =   360
            TabIndex        =   6
            Top             =   0
            Width           =   6495
            _ExtentX        =   11456
            _ExtentY        =   5741
         End
      End
   End
   Begin PhotoDemon.pdContainer picCategory 
      Height          =   6495
      Index           =   1
      Left            =   5880
      TabIndex        =   11
      Top             =   1320
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   11245
      Begin PhotoDemon.pdHyperlink hypWebOptimized 
         Height          =   255
         Index           =   0
         Left            =   0
         Top             =   4110
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   450
         Alignment       =   1
         Caption         =   "lossy optimization is provided by pngquant"
         FontSize        =   9
         URL             =   "https://pngquant.org/"
      End
      Begin PhotoDemon.pdButton cmdUpdateLossyPreview 
         Height          =   615
         Left            =   360
         TabIndex        =   19
         Top             =   3360
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   1085
         Caption         =   "click to generate a new preview image"
      End
      Begin PhotoDemon.pdTitle ttlWebOptimize 
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   17
         Top             =   120
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   661
         Caption         =   "lossy optimization options"
         FontBold        =   -1  'True
         FontSize        =   12
      End
      Begin PhotoDemon.pdCheckBox chkOptimizeDither 
         Height          =   375
         Left            =   360
         TabIndex        =   15
         Top             =   1005
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   661
         Caption         =   "use dithering to improve quality"
      End
      Begin PhotoDemon.pdSlider sltTargetQuality 
         Height          =   735
         Left            =   360
         TabIndex        =   13
         Top             =   1440
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
         TabIndex        =   12
         Top             =   630
         Width           =   6735
         _ExtentX        =   12515
         _ExtentY        =   661
         Caption         =   "apply lossy optimizations"
      End
      Begin PhotoDemon.pdSlider sltLossyPerformance 
         Height          =   735
         Left            =   360
         TabIndex        =   14
         Top             =   2190
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
         TabIndex        =   16
         Top             =   5160
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
         Top             =   2970
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
         Top             =   2970
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
         Top             =   5910
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
         Top             =   5910
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
         TabIndex        =   18
         Top             =   4680
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   661
         Caption         =   "lossless optimization options"
         FontBold        =   -1  'True
         FontSize        =   12
      End
      Begin PhotoDemon.pdHyperlink hypWebOptimized 
         Height          =   255
         Index           =   1
         Left            =   0
         Top             =   6240
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   450
         Alignment       =   1
         Caption         =   "lossless optimization is provided by OptiPNG"
         FontSize        =   9
         URL             =   "http://optipng.sourceforge.net/"
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
'Copyright 2012-2020 by Tanner Helland
'Created: 11/December/12
'Last updated: 24/April/19
'Last update: remove any remaining dependence on FreeImage; all PNG duties are now handled by our homebrew PNG engine
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit https://photodemon.org/license/
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

Private Sub UpdateMasterPanelVisibility()
    Dim i As Long
    For i = picCategory.lBound To picCategory.UBound
        picCategory(i).Visible = (btsMasterType.ListIndex = i)
    Next i
End Sub

Private Sub btsMasterType_Click(ByVal buttonIndex As Long)
    UpdateMasterPanelVisibility
End Sub

Private Sub chkEmbedBackground_Click()
    UpdateBkgdColorVisibility
End Sub

Private Sub UpdateBkgdColorVisibility()
    clsBackground.Visible = chkEmbedBackground.Value
End Sub

Private Sub chkOptimizeDither_Click()
    UpdatePreviewButtonText
End Sub

Private Sub chkOptimizeLossy_Click()
    EnableLossyOptimizationOptions
End Sub

Private Sub EnableLossyOptimizationOptions()
    
    Dim enabledState As Boolean
    enabledState = chkOptimizeLossy.Value
    
    chkOptimizeDither.Enabled = enabledState
    sltTargetQuality.Enabled = enabledState
    sltLossyPerformance.Enabled = enabledState
    lblHint(2).Enabled = enabledState
    lblHint(3).Enabled = enabledState
    cmdUpdateLossyPreview.Enabled = enabledState
    
End Sub

Private Sub UpdatePreviewButtonText()
    If Strings.StringsNotEqual(cmdUpdateLossyPreview.Caption, g_Language.TranslateMessage("click to generate a new preview image"), False) Then
        cmdUpdateLossyPreview.Caption = g_Language.TranslateMessage("click to generate a new preview image")
    End If
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
    
    If (btsMasterType.ListIndex = 0) Then
        m_MetadataParamString = mtdManager.GetMetadataSettings
    
    'While in web optimization mode, we forcibly request no metadata writing
    Else
        m_MetadataParamString = mtdManager.GetNullMetadataSettings
    End If
    
    m_UserDialogAnswer = vbOK
    Me.Visible = False
    
End Sub

Private Sub cmdBar_ReadCustomPresetData()
    ReflowWebOptimizePanel
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub cmdBar_ResetClick()
    
    cmdBar.SetPreviewStatus False
    
    'General panel settings
    sldCompression.Value = sldCompression.NotchValueCustom
    
    If (Not m_SrcImage Is Nothing) Then
        If m_SrcImage.ImgStorage.DoesKeyExist("pngBackgroundColor") Then
            clsBackground.Color = m_SrcImage.ImgStorage.GetEntry_Long("pngBackgroundColor")
            chkEmbedBackground.Value = True
        Else
            clsBackground.Color = vbWhite
            chkEmbedBackground.Value = False
        End If
    Else
        clsBackground.Color = vbWhite
        chkEmbedBackground.Value = False
    End If
    
    'Web-optimized settings
    chkOptimizeLossy.Value = True
    sltTargetQuality.Value = sltTargetQuality.NotchValueCustom
    sltLossyPerformance.Value = sltLossyPerformance.NotchValueCustom
    chkOptimizeDither.Value = True
    sltLosslessPerformance.Value = sltLosslessPerformance.NotchValueCustom
    
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
    
    'PNG has two master panels: standard PNGs, and web-optimized PNGs
    btsMasterType.AddItem "standard PNG", 0
    btsMasterType.AddItem "web-optimized PNG", 1
    btsMasterType.ListIndex = 0
    
    'Standard settings are accessed via pdTitle controls.  Because the panels are so large, only one panel
    ' is allowed open at a time.
    Dim i As Long
    For i = picContainer.lBound To picContainer.UBound
        picContainer(i).SetLeft 0
    Next i
    
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
    
    'Populate web-optimized options
    EnableLossyOptimizationOptions
    
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
        If m_SrcImage.ImgStorage.DoesKeyExist("pngBackgroundColor") Then
            clsBackground.Color = m_SrcImage.ImgStorage.GetEntry_Long("pngBackgroundColor")
            chkEmbedBackground.Value = True
        End If
    End If
    
    UpdateBkgdColorVisibility
    
    'Update the preview
    UpdatePreview
    
    'Apply translations and visual themes
    ApplyThemeAndTranslations Me
    
    'Many UI options are dynamically shown/hidden depending on other settings; make sure their initial state is correct
    UpdateMasterPanelVisibility
    UpdateStandardPanelVisibility
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
    If (Not m_CompositedImage Is Nothing) Then
        
        'Because the user can change the preview viewport, we can't guarantee that the preview region hasn't changed
        ' since the last preview.  Prep a new preview now.
        Dim tmpSafeArray As SafeArray2D
        EffectPrep.PreviewNonStandardImage tmpSafeArray, m_CompositedImage, pdFxPreview, False
        
        'Write the image out to a temporary file
        Dim tmpFilename As String
        tmpFilename = Files.RequestTempFile() & ".png"
        
        Dim cPNG As pdPNG
        Set cPNG = New pdPNG
        If (cPNG.SavePNG_ToFile(tmpFilename, workingDIB, Nothing, png_TruecolorAlpha, 8, 0) < png_Failure) Then
            
            'Retrieve the size of the base PNG file
            Dim oldFileSize As Long
            oldFileSize = Files.FileLenW(tmpFilename)
            
            'Next, request optimization from pngquant
            If Plugin_PNGQuant.ApplyPNGQuantToFile_Synchronous(tmpFilename, sltTargetQuality.Value, 11 - sltLossyPerformance.Value, chkOptimizeDither.Value, False) Then
                
                Dim newFileSize As Long
                newFileSize = Files.FileLenW(tmpFilename)
                
                'If successful, pngquant will overwrite the original file with its optimized copy.  Retrieve it now.
                If Loading.QuickLoadImageToDIB(tmpFilename, workingDIB, False) Then
                    EffectPrep.FinalizeNonstandardPreview Me.pdFxPreview, False
                    updateSuccess = True
                End If
                
            Else
                PDDebug.LogAction "WARNING!  The pngquant preview step failed for reasons unknown!"
            End If
            
            Files.FileDeleteIfExists tmpFilename
            
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
End Sub

Private Function GetExportParamString() As String

    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    
    'The parameters this function returns vary based on the current PNG mode (standard vs web-optimized).
    cParams.AddParam "PNGCreateWebOptimized", (btsMasterType.ListIndex = 1)
    
    'Standard parameters are the more complicated ones, if you can believe it
    If (btsMasterType.ListIndex = 0) Then
    
        'Start with the standard PNG settings, which are consistent across all standard PNG types
        If sldCompression.IsValid Then cParams.AddParam "PNGCompressionLevel", sldCompression.Value Else cParams.AddParam "PNGCompressionLevel", sldCompression.NotchValueCustom
        cParams.AddParam "PNGBackgroundColor", clsBackground.Color
        cParams.AddParam "PNGCreateBkgdChunk", chkEmbedBackground.Value
        cParams.AddParam "PNGFilterStrategy", cboOptimize.ListIndex
        
        'Next come all the messy color-depth possibilities
        cParams.AddParam "PNGColorDepth", clrDepth.GetAllSettings
        
    'Remember: web-optimized parameters must not use any UI elements from the "normal" settings panel!
    Else
    
        cParams.AddParam "PNGOptimizeLossy", chkOptimizeLossy.Value
        cParams.AddParam "PNGOptimizeLossyQuality", sltTargetQuality.Value
        
        'pngquant accepts this value on a 1-11 scale, with 1 being slowest and 11 being fastest.  We show the user a
        ' [0, 10] scale where [10] is slowest (like the other settings on the form); reset to the proper range now.
        cParams.AddParam "PNGOptimizeLossyPerformance", 11 - sltLossyPerformance.Value
        cParams.AddParam "PNGOptimizeLossyDithering", chkOptimizeDither.Value
        
        cParams.AddParam "PNGOptimizeLosslessPerformance", sltLosslessPerformance.Value
        
    End If
    
    GetExportParamString = cParams.GetParamString
    
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
        
        'The color-depth-specific options are embedded as a single option, so extract them into their
        ' own parser.
        Dim cParamsDepth As pdSerialize
        Set cParamsDepth = New pdSerialize
        cParamsDepth.SetParamString cParams.GetString("PNGColorDepth", vbNullString)
        
        'Retrieve color and alpha model for this preview; everything else extends from these
        Dim outputColorModel As String, outputAlphaModel As String
        outputColorModel = cParamsDepth.GetString("ColorDepth_ColorModel", "Auto", True)
        outputAlphaModel = cParamsDepth.GetString("ColorDepth_AlphaModel", "Auto", True)
    
        'Before doing anything else, figure out how to handle alpha.
        Dim previewAlphaMode As PD_ALPHA_STATUS
        If Strings.StringsEqual(outputAlphaModel, "full", True) Then
            previewAlphaMode = PDAS_ComplicatedAlpha
        ElseIf Strings.StringsEqual(outputAlphaModel, "none", True) Then
            previewAlphaMode = PDAS_NoAlpha
        ElseIf Strings.StringsEqual(outputAlphaModel, "bycutoff", True) Then
            previewAlphaMode = PDAS_BinaryAlpha
        ElseIf Strings.StringsEqual(outputAlphaModel, "bycolor", True) Then
            previewAlphaMode = PDAS_NewAlphaFromColor
        Else
            previewAlphaMode = PDAS_ComplicatedAlpha
        End If
        
        Dim bkgdColor As Long, trnsTable() As Byte
        bkgdColor = cParams.GetLong("PNGBackgroundColor", vbWhite)
        
        Dim outputAlphaCutoff As Long, outputAlphaColor As Long
        outputAlphaCutoff = cParamsDepth.GetLong("ColorDepth_AlphaCutoff", PD_DEFAULT_ALPHA_CUTOFF)
        outputAlphaColor = cParamsDepth.GetLong("ColorDepth_AlphaColor", vbMagenta)
    
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
            forceGrayscale = ParamsEqual(cParamsDepth.GetString("ColorDepth_ColorModel", "Auto"), "Gray")
            forceIndexed = ParamsEqual(cParamsDepth.GetString("ColorDepth_ColorDepth", "Color_Standard"), "Color_Indexed")
            newPaletteSize = cParamsDepth.GetLong("ColorDepth_PaletteSize", 256)
            If ParamsEqual(cParamsDepth.GetString("ColorDepth_GrayDepth", "Auto"), "Gray_Monochrome") Then newPaletteSize = 2
            
            If forceGrayscale Then
                DIBs.MakeDIBGrayscale workingDIB, newPaletteSize
                
            ElseIf forceIndexed Then
                Dim newPalette() As RGBQuad
                Palettes.GetOptimizedPaletteIncAlpha workingDIB, newPalette, newPaletteSize
                Palettes.ApplyPaletteToImage_IncAlpha_KDTree workingDIB, newPalette, True
            End If
            
        End If
        
        'If the image is in "use original settings" mode, we will need to forcibly overwrite various
        ' settings to match the original file's settings.)
        ' (TODO!)
        'Dim useOrigMode As Boolean
        'useOrigMode = ParamsEqual(cParamsDepth.GetString("ColorDepth_ColorModel", "Original"), "Original")
        
        'In "use original file settings" mode, we need to steal a palette copy from the source image
        ' (TODO!)
        'Dim tmpPalette As pdPalette
        'If (useOrigMode And m_SrcImage.HasOriginalPalette) Then m_SrcImage.GetOriginalPalette tmpPalette
        
        EffectPrep.FinalizeNonstandardPreview pdFxPreview, True
        
    End If
    
End Sub

Private Sub sltLossyPerformance_Change()
    UpdatePreviewButtonText
End Sub

Private Sub sltTargetQuality_Change()
    UpdatePreviewButtonText
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
    hypWebOptimized(0).Visible = isVisible
    
    'Determine a vertical offset for the bottom part of the panel, contingent on the top panel being open or shut
    If isVisible Then
        offsetY = hypWebOptimized(0).GetTop + hypWebOptimized(0).GetHeight + Interface.FixDPI(16)
    Else
        offsetY = ttlWebOptimize(0).GetTop + ttlWebOptimize(0).GetHeight + Interface.FixDPI(16)
    End If
    
    'Show/hide the lossless compression options
    ttlWebOptimize(1).SetTop offsetY
    isVisible = ttlWebOptimize(1).Value
    
    If isVisible Then
        offsetY = ttlWebOptimize(1).GetTop + ttlWebOptimize(1).GetHeight + Interface.FixDPI(6)
        sltLosslessPerformance.SetTop offsetY
        offsetY = sltLosslessPerformance.GetTop + sltLosslessPerformance.GetHeight + Interface.FixDPI(3)
        lblHint(4).SetTop offsetY
        lblHint(5).SetTop offsetY
        offsetY = offsetY + lblHint(5).GetHeight + Interface.FixDPI(8)
        hypWebOptimized(1).SetTop offsetY
    End If
    
    sltLosslessPerformance.Visible = isVisible
    lblHint(4).Visible = isVisible
    lblHint(5).Visible = isVisible
    hypWebOptimized(1).Visible = isVisible
    
End Sub
