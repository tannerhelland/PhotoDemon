VERSION 5.00
Begin VB.Form FormPalettize 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Palettize"
   ClientHeight    =   7020
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12315
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   468
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   821
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.pdButtonStrip btsOptions 
      Height          =   615
      Left            =   5880
      TabIndex        =   4
      Top             =   120
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   1085
   End
   Begin PhotoDemon.pdCommandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   6270
      Width           =   12315
      _ExtentX        =   21722
      _ExtentY        =   1323
   End
   Begin PhotoDemon.pdFxPreviewCtl pdFxPreview 
      Height          =   6105
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9922
   End
   Begin PhotoDemon.pdContainer pnlQuantize 
      Height          =   5175
      Index           =   1
      Left            =   5880
      TabIndex        =   2
      Top             =   960
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   9128
      Begin PhotoDemon.pdListBox lstPalettes 
         Height          =   2775
         Left            =   120
         TabIndex        =   21
         Top             =   840
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   4895
         Caption         =   "palettes in this file:"
      End
      Begin PhotoDemon.pdButton cmdLoadPalette 
         Height          =   495
         Left            =   5400
         TabIndex        =   18
         Top             =   345
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   873
         Caption         =   "..."
      End
      Begin PhotoDemon.pdTextBox txtPalette 
         Height          =   375
         Left            =   360
         TabIndex        =   17
         Top             =   390
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   661
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   255
         Left            =   120
         Top             =   0
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   450
         Caption         =   "choose a palette file:"
         FontSize        =   12
      End
      Begin PhotoDemon.pdDropDown cboDither 
         Height          =   855
         Index           =   1
         Left            =   120
         TabIndex        =   19
         Top             =   3720
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   1508
         Caption         =   "dithering"
      End
      Begin PhotoDemon.pdCheckBox chkReduceBleed 
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   20
         Top             =   4680
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   661
         Caption         =   "reduce color bleed"
         Value           =   0
      End
   End
   Begin PhotoDemon.pdContainer pnlQuantize 
      Height          =   4800
      Index           =   0
      Left            =   5880
      TabIndex        =   3
      Top             =   960
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   8467
      Begin PhotoDemon.pdTitle ttlStandard 
         Height          =   375
         Index           =   0
         Left            =   0
         TabIndex        =   15
         Top             =   0
         Width           =   6315
         _ExtentX        =   11139
         _ExtentY        =   661
         Caption         =   "basic settings"
         FontSize        =   12
      End
      Begin PhotoDemon.pdTitle ttlStandard 
         Height          =   375
         Index           =   1
         Left            =   0
         TabIndex        =   16
         Top             =   360
         Width           =   6315
         _ExtentX        =   11139
         _ExtentY        =   661
         Caption         =   "advanced settings"
         FontSize        =   12
         Value           =   0   'False
      End
      Begin PhotoDemon.pdContainer pnlBasic 
         Height          =   3735
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   6588
         Begin PhotoDemon.pdDropDown cboDither 
            Height          =   855
            Index           =   0
            Left            =   120
            TabIndex        =   6
            Top             =   2400
            Width           =   6015
            _ExtentX        =   10610
            _ExtentY        =   1508
            Caption         =   "dithering"
         End
         Begin PhotoDemon.pdCheckBox chkReduceBleed 
            Height          =   375
            Index           =   0
            Left            =   240
            TabIndex        =   7
            Top             =   3300
            Width           =   5895
            _ExtentX        =   10398
            _ExtentY        =   661
            Caption         =   "reduce color bleed"
            Value           =   0
         End
         Begin PhotoDemon.pdSlider sldPalette 
            Height          =   735
            Left            =   120
            TabIndex        =   8
            Top             =   1200
            Width           =   6015
            _ExtentX        =   10610
            _ExtentY        =   1296
            Caption         =   "palette size"
            Min             =   2
            Max             =   256
            Value           =   256
            GradientColorRight=   1703935
            NotchPosition   =   2
            NotchValueCustom=   256
         End
         Begin PhotoDemon.pdButtonStrip btsMethod 
            Height          =   1095
            Left            =   120
            TabIndex        =   9
            Top             =   0
            Width           =   6015
            _ExtentX        =   10610
            _ExtentY        =   1931
            Caption         =   "quantization method"
         End
         Begin PhotoDemon.pdCheckBox chkPreserveWB 
            Height          =   375
            Left            =   240
            TabIndex        =   10
            Top             =   1965
            Width           =   5895
            _ExtentX        =   10398
            _ExtentY        =   661
            Caption         =   "preserve white and black"
         End
      End
      Begin PhotoDemon.pdContainer pnlBasic 
         Height          =   3375
         Index           =   1
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   5953
         Begin PhotoDemon.pdButtonStrip btsAlpha 
            Height          =   1095
            Left            =   120
            TabIndex        =   12
            Top             =   1200
            Width           =   6015
            _ExtentX        =   10610
            _ExtentY        =   1931
            Caption         =   "transparency"
         End
         Begin PhotoDemon.pdSlider sldAlphaCutoff 
            Height          =   855
            Left            =   120
            TabIndex        =   13
            Top             =   2400
            Width           =   6015
            _ExtentX        =   10610
            _ExtentY        =   1508
            Caption         =   "alpha cut-off"
            Max             =   254
            SliderTrackStyle=   1
            Value           =   64
            GradientColorRight=   1703935
            NotchPosition   =   2
            NotchValueCustom=   64
         End
         Begin PhotoDemon.pdColorSelector clsBackground 
            Height          =   1095
            Left            =   120
            TabIndex        =   14
            Top             =   0
            Width           =   6015
            _ExtentX        =   10610
            _ExtentY        =   1931
            Caption         =   "background color"
         End
      End
   End
End
Attribute VB_Name = "FormPalettize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'"Palettize" (e.g. reduce image color count) Dialog
'Copyright 2000-2018 by Tanner Helland
'Created: 4/October/00
'Last updated: 01/February/18
'Last update: improve progress bar reporting
'
'This dialog allows the user to reduce the number of colors in the current image.  In the future, it would be nice
' to allow palettes loaded from file or selected from an internal swatch manager (and in fact, the code is already
' set up to handle this) but at present, only optimized palettes are supported.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Used to avoid recursive setting changes
Private m_ActiveTitleBar As Long, m_PanelChangesActive As Boolean

'When loading a palette from file, the pdPalette class handles all the actual parsing
Private m_Palette As pdPalette

'To reduce redraws when interacting with the UI, we manually track changes to the palette filename
Private m_PalettePath As String, m_PaletteFileSize As Long

Private Sub btsAlpha_Click(ByVal buttonIndex As Long)
    UpdateTransparencyOptions
    UpdatePreview
End Sub

Private Sub UpdateTransparencyOptions()
    sldAlphaCutoff.Visible = (btsAlpha.ListIndex = 1)
End Sub

Private Sub btsMethod_Click(ByVal buttonIndex As Long)
    UpdatePreview
End Sub

Private Sub btsOptions_Click(ByVal buttonIndex As Long)
    UpdateVisiblePanel
    UpdatePreview
End Sub

Private Sub UpdateVisiblePanel()
    Dim i As Long
    For i = pnlQuantize.lBound To pnlQuantize.UBound
        pnlQuantize(i).Visible = (i = btsOptions.ListIndex)
    Next i
End Sub

Private Sub cboDither_Click(Index As Integer)
    UpdatePreview
    UpdateColorBleedVisibility
End Sub

Private Sub chkPreserveWB_Click()
    UpdatePreview
End Sub

Private Sub chkReduceBleed_Click(Index As Integer)
    UpdatePreview
End Sub

Private Sub clsBackground_ColorChanged()
    UpdatePreview
End Sub

'OK button
Private Sub cmdBar_OKClick()
    Process "Palettize", , GetToolParamString, UNDO_Layer
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub cmdBar_ResetClick()
    btsMethod.ListIndex = 0
    clsBackground.Color = vbWhite
    btsAlpha.ListIndex = 0
    sldAlphaCutoff.Value = sldAlphaCutoff.NotchValueCustom
    UpdatePreview
End Sub

Private Sub cmdLoadPalette_Click()
    Dim srcPaletteFile As String
    If Palettes.DisplayPaletteLoadDialog(vbNullString, srcPaletteFile) Then txtPalette.Text = srcPaletteFile
End Sub

Private Sub Form_Load()
    
    'Suspend previews until the dialog has been fully initialized
    cmdBar.MarkPreviewStatus False
    
    btsOptions.AddItem "optimal", 0
    btsOptions.AddItem "from file", 1
    btsOptions.ListIndex = 0
    UpdateVisiblePanel
    
    btsMethod.AddItem "median cut", 0
    btsMethod.AddItem "Xiaolin Wu", 1
    btsMethod.AddItem "NeuQuant", 2
    btsMethod.ListIndex = 0
    
    btsAlpha.AddItem "full", 0
    btsAlpha.AddItem "binary", 1
    btsAlpha.AddItem "none", 2
    btsAlpha.ListIndex = 0
    UpdateTransparencyOptions
    
    Dim i As Long
    For i = cboDither.lBound To cboDither.UBound
        cboDither(i).Clear
        cboDither(i).AddItem "None", 0
        cboDither(i).AddItem "Ordered (Bayer 4x4)", 1
        cboDither(i).AddItem "Ordered (Bayer 8x8)", 2
        cboDither(i).AddItem "Single neighbor", 3
        cboDither(i).AddItem "Floyd-Steinberg", 4
        cboDither(i).AddItem "Jarvis, Judice, and Ninke", 5
        cboDither(i).AddItem "Stucki", 6
        cboDither(i).AddItem "Burkes", 7
        cboDither(i).AddItem "Sierra-3", 8
        cboDither(i).AddItem "Two-Row Sierra", 9
        cboDither(i).AddItem "Sierra Lite", 10
        cboDither(i).AddItem "Atkinson / Classic Macintosh", 11
        cboDither(i).ListIndex = 6
    Next i
    
    UpdateColorBleedVisibility
    
    'Many UI options are dynamically shown/hidden depending on other settings; make sure their initial state is correct
    ttlStandard(0).Value = True
    m_ActiveTitleBar = 0
    UpdateStandardTitlebars
    
    'UpdateMasterPanelVisibility
    UpdateStandardPanelVisibility
    
    'Apply translations and visual themes
    ApplyThemeAndTranslations Me
    cmdBar.MarkPreviewStatus True
    UpdatePreview
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'Automatic 8-bit color reduction.  Some option combinations require the FreeImage plugin.
Private Sub ApplyRuntimePalettizeEffect(ByVal toolParams As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)
    
    'Parse the parameter string and determine concrete values for our color conversion
    Dim cParams As pdParamXML
    Set cParams = New pdParamXML
    cParams.SetParamString toolParams
    
    Dim quantMethod As PD_COLOR_QUANTIZE
    If Strings.StringsEqual(cParams.GetString("method", "mediancut"), "neuquant", True) Then
        quantMethod = PDCQ_Neuquant
    ElseIf Strings.StringsEqual(cParams.GetString("method", "mediancut"), "wu", True) Then
        quantMethod = PDCQ_Wu
    Else
        quantMethod = PDCQ_MedianCut
    End If
    
    Dim paletteSize As Long
    paletteSize = cParams.GetLong("palettesize", 256)
    
    Dim preserveWhiteBlack As Boolean
    preserveWhiteBlack = cParams.GetBool("preservewhiteblack", False)
    
    Dim ditherMethod As PD_DITHER_METHOD
    ditherMethod = cParams.GetLong("dithering", 0)
    
    Dim reduceBleed As Boolean
    reduceBleed = cParams.GetBool("reducebleed", False)
    
    Dim finalBackColor As Long
    finalBackColor = cParams.GetLong("backgroundcolor", vbWhite)
    
    Dim outputAlphaMode As PD_ALPHA_STATUS
    If Strings.StringsEqual(cParams.GetString("alphamode", "full"), "full", True) Then
        outputAlphaMode = PDAS_ComplicatedAlpha
    ElseIf Strings.StringsEqual(cParams.GetString("alphamode", "full"), "binary", True) Then
        outputAlphaMode = PDAS_BinaryAlpha
    Else
        outputAlphaMode = PDAS_NoAlpha
    End If
    
    Dim alphaCutoff As Long
    alphaCutoff = cParams.GetLong("alphacutoff", 64)
    
    Dim tmpSA As SafeArray2D
    EffectPrep.PrepImageData tmpSA, toPreview, pdFxPreview
    
    If (Not toPreview) Then
        SetProgBarMax workingDIB.GetDIBHeight * 2
        SetProgBarVal 0
        Message "Generating optimal palette..."
    End If
    
    'Some quantization methods require FreeImage.  If FreeImage doesn't exist, fall back to internal PD methods.
    If (quantMethod <> PDCQ_MedianCut) Then
        If (Not g_ImageFormats.FreeImageEnabled) Then quantMethod = PDCQ_MedianCut
    End If
    
    'If the caller doesn't want transparency, composite the image against the specified backcolor *in advance*.
    Dim currentAlphaState As PD_ALPHA_STATUS
    currentAlphaState = PDAS_ComplicatedAlpha
    
    If (outputAlphaMode = PDAS_NoAlpha) Then
        workingDIB.CompositeBackgroundColor Colors.ExtractRed(finalBackColor), Colors.ExtractGreen(finalBackColor), Colors.ExtractBlue(finalBackColor)
        currentAlphaState = PDAS_NoAlpha
        
    'Similarly, if they want binary alpha treatment, apply that now as well.
    ElseIf (outputAlphaMode = PDAS_BinaryAlpha) Then
        
        Dim transTable() As Byte
        ReDim transTable(0 To 255) As Byte
        DIBs.ApplyAlphaCutoff_Ex workingDIB, transTable, alphaCutoff
        DIBs.ApplyBinaryTransparencyTable workingDIB, transTable, finalBackColor
        
        currentAlphaState = PDAS_BinaryAlpha
        
    End If
    
    'Branch according to internal or plugin-based quantization methods.  Note that if the user does *NOT* want
    ' dithering, we can use the plugin to apply the palette as well, trimming processing time a bit.
    Dim finalPalette() As RGBQuad, finalPaletteCount As Long
    
    If (Not toPreview) Then
        SetProgBarVal workingDIB.GetDIBHeight
        Message "Applying new palette to image..."
    End If
    
    If (quantMethod = PDCQ_MedianCut) Then
    
        'Resize the target DIB to a smaller size
        Dim smallDIB As pdDIB
        DIBs.ResizeDIBByPixelCount workingDIB, smallDIB, 500000
        
        'Generate an optimal palette
        Palettes.GetOptimizedPalette smallDIB, finalPalette, paletteSize
        
        'Preserve black and white, as necessary
        If preserveWhiteBlack Then Palettes.EnsureBlackAndWhiteInPalette finalPalette, smallDIB
        
        'Apply said palette to the image
        If (ditherMethod = PDDM_None) Then
            Palettes.ApplyPaletteToImage_KDTree workingDIB, finalPalette, toPreview, workingDIB.GetDIBHeight * 2, workingDIB.GetDIBHeight
        Else
            Palettes.ApplyPaletteToImage_Dithered workingDIB, finalPalette, ditherMethod, reduceBleed, toPreview, workingDIB.GetDIBHeight * 2, workingDIB.GetDIBHeight
        End If
    
    Else
        
        'Apply all color and transparency changes simultaneously
        Dim fiQuantMode As FREE_IMAGE_QUANTIZE
        If (quantMethod = PDCQ_Wu) Then fiQuantMode = FIQ_WUQUANT Else fiQuantMode = FIQ_NNQUANT
        
        Dim fi_DIB8 As Long
        fi_DIB8 = Plugin_FreeImage.GetFIDib_SpecificColorMode(workingDIB, 8, outputAlphaMode, currentAlphaState, alphaCutoff, finalBackColor, , paletteSize, , , fiQuantMode)
        FreeImage_FlipVertically fi_DIB8
        
        'If the caller does *not* want dithering, copy the (already palettized) FreeImage DIB over our
        ' original DIB.
        If (ditherMethod = PDDM_None) And (Not preserveWhiteBlack) Then
        
            'Convert that DIB to 32-bpp
            Dim fi_DIB As Long
            fi_DIB = FreeImage_ConvertTo32Bits(fi_DIB8)
            FreeImage_Unload fi_DIB8
            
            'Paint the result to workingDIB
            workingDIB.ResetDIB 0
            Plugin_FreeImage.PaintFIDibToPDDib workingDIB, fi_DIB, 0, 0, workingDIB.GetDIBWidth, workingDIB.GetDIBHeight
            FreeImage_Unload fi_DIB
        
        'If the caller wants dithering, we must apply the palette manually
        Else
        
            'Retrieve the generated palette, then free the FreeImage source
            finalPaletteCount = Plugin_FreeImage.GetFreeImagePalette(fi_DIB8, finalPalette)
            ReDim Preserve finalPalette(0 To paletteSize - 1) As RGBQuad
            FreeImage_Unload fi_DIB8
            
            'Preserve black and white, as necessary
            If preserveWhiteBlack Then Palettes.EnsureBlackAndWhiteInPalette finalPalette, smallDIB
            
            'Apply the generated palette to our target image, using the method requested
            If (finalPaletteCount <> 0) Then
                If (ditherMethod = PDDM_None) Then
                    Palettes.ApplyPaletteToImage_KDTree workingDIB, finalPalette, toPreview, workingDIB.GetDIBHeight * 2, workingDIB.GetDIBHeight
                Else
                    Palettes.ApplyPaletteToImage_Dithered workingDIB, finalPalette, ditherMethod, reduceBleed, toPreview, workingDIB.GetDIBHeight * 2, workingDIB.GetDIBHeight
                End If
            End If
            
        End If
        
    End If
    
    EffectPrep.FinalizeImageData toPreview, pdFxPreview
    
End Sub

'Automatic 8-bit color reduction.  Some option combinations require the FreeImage plugin.
Private Sub ApplyPaletteFromFile(ByVal toolParams As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)
    
    'Parse the parameter string and determine concrete values for our color conversion
    Dim cParams As pdParamXML
    Set cParams = New pdParamXML
    cParams.SetParamString toolParams
    
    Dim srcPaletteFile As String
    srcPaletteFile = cParams.GetString("palettefile")
    
    Dim mustLoadPalette As Boolean: mustLoadPalette = True
    If (Not m_Palette Is Nothing) Then mustLoadPalette = Strings.StringsNotEqual(srcPaletteFile, m_Palette.GetPaletteFilename())
    If mustLoadPalette Then
        Set m_Palette = New pdPalette
        m_Palette.LoadPaletteFromFile srcPaletteFile, True, False
    End If
    
    'Make sure the passed palette group ID is valid.  (Some palette file formats support multiple palettes
    ' within a single file; as such, filename alone may not be enough to identify the palette we need.)
    Dim srcPaletteIndex As Long
    srcPaletteIndex = cParams.GetLong("palettefileindex", 0)
    If (srcPaletteIndex < 0) Then srcPaletteIndex = 0
    If (srcPaletteIndex > m_Palette.GetPaletteGroupCount - 1) Then srcPaletteIndex = m_Palette.GetPaletteGroupCount - 1
    
    Dim ditherMethod As PD_DITHER_METHOD
    ditherMethod = cParams.GetLong("dithering", 0)
    
    Dim reduceBleed As Boolean
    reduceBleed = cParams.GetBool("reducebleed", False)
    
    Dim tmpSA As SafeArray2D
    EffectPrep.PrepImageData tmpSA, toPreview, pdFxPreview
    
    If (Not toPreview) Then
        SetProgBarMax workingDIB.GetDIBHeight
        SetProgBarVal 0
    End If
    
    'Branch according to internal or plugin-based quantization methods.  Note that if the user does *NOT* want
    ' dithering, we can use the plugin to apply the palette as well, trimming processing time a bit.
    Dim finalPalette() As RGBQuad
    If (m_Palette.GetPaletteColorCount(srcPaletteIndex) > 0) Then
        
        m_Palette.CopyPaletteToArray finalPalette, srcPaletteIndex
        
        If (Not toPreview) Then Message "Applying new palette to image..."
        
        'Apply said palette to the image
        If (ditherMethod = PDDM_None) Then
            Palettes.ApplyPaletteToImage_KDTree workingDIB, finalPalette, toPreview, workingDIB.GetDIBHeight
        Else
            Palettes.ApplyPaletteToImage_Dithered workingDIB, finalPalette, ditherMethod, reduceBleed, toPreview, workingDIB.GetDIBHeight
        End If
        
    End If
    
    EffectPrep.FinalizeImageData toPreview, pdFxPreview
    
End Sub

Private Sub UpdateColorBleedVisibility()
    Dim i As Long
    For i = cboDither.lBound To cboDither.UBound
        chkReduceBleed(i).Visible = (cboDither(i).ListIndex <> 0)
    Next i
End Sub

Private Sub lstPalettes_Click()
    UpdatePreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub pdFxPreview_ViewportChanged()
    UpdatePreview
End Sub

Private Sub sldAlphaCutoff_Change()
    UpdatePreview
End Sub

Private Sub sldPalette_Change()
    UpdatePreview
End Sub

Private Sub ttlStandard_Click(Index As Integer, ByVal newState As Boolean)

    If newState Then m_ActiveTitleBar = Index
    pnlBasic(Index).Visible = newState
    
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
        pnlBasic(i).Visible = ttlStandard(i).Value
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
    yPadding = Interface.FixDPI(8)
    
    Dim i As Long
    For i = ttlStandard.lBound To ttlStandard.UBound
    
        ttlStandard(i).SetTop yPos
        yPos = yPos + ttlStandard(i).GetHeight + yPadding
        
        If ttlStandard(i).Value Then
            pnlBasic(i).SetTop yPos
            yPos = yPos + pnlBasic(i).GetHeight + yPadding
        End If
        
    Next i
    
End Sub

'This function simply sorts incoming palettize requests by type, then calls the appropriate sub-function to actually
' palettize the image.  (This was added in Jan '18 as part of supporting "apply palette from file" behavior.)
Public Sub ApplyPalettizeEffect(ByVal toolParams As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)
    
    Dim cParams As pdParamXML
    Set cParams = New pdParamXML
    cParams.SetParamString toolParams
    
    Dim paletteMode As Long
    paletteMode = cParams.GetLong("mode", 0)
    
    If (paletteMode = 0) Then
        ApplyRuntimePalettizeEffect toolParams, toPreview, dstPic
    Else
        ApplyPaletteFromFile toolParams, toPreview, dstPic
    End If

End Sub

Private Sub txtPalette_Change()
    UpdatePaletteFileInfo
End Sub

Private Sub UpdatePaletteFileInfo()
    
    'Try to load the palette into a dedicated class
    If (m_Palette Is Nothing) Then Set m_Palette = New pdPalette
    If Files.FileExists(txtPalette.Text) Then
        
        'See if the palette file has changed since our last attempt at loading.
        If Strings.StringsNotEqual(m_PalettePath, txtPalette.Text) Or (m_PaletteFileSize <> Files.FileLenW(txtPalette.Text)) Then
            
            m_PalettePath = txtPalette.Text
            m_PaletteFileSize = Files.FileLenW(m_PalettePath)
            
            lstPalettes.Clear
            lstPalettes.SetAutomaticRedraws False, False
            
            If m_Palette.LoadPaletteFromFile(txtPalette.Text, True, False) Then
                
                Dim i As Long, palText As String
                For i = 0 To m_Palette.GetPaletteGroupCount - 1
                    palText = g_Language.TranslateMessage("%1 (%2 colors)", m_Palette.GetPaletteName(i), m_Palette.GetPaletteColorCount(i))
                    lstPalettes.AddItem palText, i
                Next i
                
                If (lstPalettes.ListCount > 0) Then lstPalettes.ListIndex = 0
                
            Else
                lstPalettes.AddItem g_Language.TranslateMessage("WARNING!  Palette file invalid.")
            End If
            
            lstPalettes.SetAutomaticRedraws True, True
            
            UpdatePreview
            
        End If
        
    End If
        
End Sub

Private Function GetToolParamString() As String

    Dim cParams As pdParamXML
    Set cParams = New pdParamXML
    
    With cParams
        
        .AddParam "mode", btsOptions.ListIndex
        
        Select Case btsMethod.ListIndex
            Case 0
                .AddParam "method", "MedianCut"
            Case 1
                .AddParam "method", "Wu"
            Case 2
                .AddParam "method", "NeuQuant"
        End Select
        
        .AddParam "palettesize", sldPalette.Value
        .AddParam "preservewhiteblack", CBool(chkPreserveWB.Value)
        .AddParam "backgroundcolor", clsBackground.Color
        
        Select Case btsAlpha.ListIndex
            Case 0
                .AddParam "alphamode", "full"
            Case 1
                .AddParam "alphamode", "binary"
            Case 2
                .AddParam "alphamode", "none"
        End Select
        
        .AddParam "alphacutoff", sldAlphaCutoff.Value
        
        '"From file" data comes next
        .AddParam "palettefile", txtPalette.Text
        .AddParam "palettefileindex", lstPalettes.ListIndex
        
        'Some options are shared between the two methods
        .AddParam "dithering", cboDither(btsOptions.ListIndex).ListIndex
        .AddParam "reducebleed", CBool(chkReduceBleed(btsOptions.ListIndex).Value)
        
    End With
    
    GetToolParamString = cParams.GetParamString

End Function

'Use this sub to update the on-screen preview
Private Sub UpdatePreview()
    If cmdBar.PreviewsAllowed Then Me.ApplyPalettizeEffect GetToolParamString, True, pdFxPreview
End Sub

