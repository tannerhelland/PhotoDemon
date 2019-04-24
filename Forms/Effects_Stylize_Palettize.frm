VERSION 5.00
Begin VB.Form FormPalettize 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Palettize"
   ClientHeight    =   7710
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
   ScaleHeight     =   514
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
      Top             =   6960
      Width           =   12315
      _ExtentX        =   21722
      _ExtentY        =   1323
   End
   Begin PhotoDemon.pdFxPreviewCtl pdFxPreview 
      Height          =   6705
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   11827
   End
   Begin PhotoDemon.pdContainer pnlQuantize 
      Height          =   6000
      Index           =   0
      Left            =   5880
      TabIndex        =   3
      Top             =   960
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   10583
      Begin PhotoDemon.pdColorSelector clsBackground 
         Height          =   855
         Left            =   120
         TabIndex        =   11
         Top             =   2160
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   1508
         Caption         =   "background color"
         FontSize        =   11
      End
      Begin PhotoDemon.pdSlider sldDitherAmount 
         Height          =   735
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   5160
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   1296
         Caption         =   "dithering amount"
         FontSizeCaption =   11
         Max             =   100
         Value           =   100
         GradientColorRight=   1703935
         DefaultValue    =   100
      End
      Begin PhotoDemon.pdDropDown cboDither 
         Height          =   735
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   4320
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   1296
         Caption         =   "dithering"
         FontSizeCaption =   11
      End
      Begin PhotoDemon.pdSlider sldPalette 
         Height          =   735
         Left            =   120
         TabIndex        =   14
         Top             =   3120
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   1296
         Caption         =   "palette size"
         FontSizeCaption =   11
         Min             =   2
         Max             =   256
         Value           =   256
         GradientColorRight=   1703935
         NotchPosition   =   2
         NotchValueCustom=   256
      End
      Begin PhotoDemon.pdButtonStrip btsMethod 
         Height          =   975
         Left            =   120
         TabIndex        =   15
         Top             =   1080
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   1720
         Caption         =   "quantization method"
         FontSizeCaption =   11
      End
      Begin PhotoDemon.pdCheckBox chkPreserveWB 
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   3885
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   661
         Caption         =   "preserve white and black"
      End
      Begin PhotoDemon.pdButtonStrip btsAlpha 
         Height          =   975
         Left            =   120
         TabIndex        =   17
         Top             =   0
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   1720
         Caption         =   "palette type"
         FontSizeCaption =   11
      End
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
      Begin PhotoDemon.pdCheckBox chkMatchAlpha 
         Height          =   375
         Left            =   210
         TabIndex        =   10
         Top             =   3090
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   661
         Caption         =   "use palette's alpha values"
      End
      Begin PhotoDemon.pdListBox lstPalettes 
         Height          =   2175
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   3836
         Caption         =   "palettes in this file:"
         FontSizeCaption =   11
      End
      Begin PhotoDemon.pdButton cmdLoadPalette 
         Height          =   495
         Left            =   5400
         TabIndex        =   6
         Top             =   345
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   873
         Caption         =   "..."
      End
      Begin PhotoDemon.pdTextBox txtPalette 
         Height          =   375
         Left            =   360
         TabIndex        =   5
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
         FontSize        =   11
      End
      Begin PhotoDemon.pdDropDown cboDither 
         Height          =   700
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   3480
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   1244
         Caption         =   "dithering"
         FontSizeCaption =   11
      End
      Begin PhotoDemon.pdSlider sldDitherAmount 
         Height          =   700
         Index           =   1
         Left            =   120
         TabIndex        =   9
         Top             =   4380
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   1244
         Caption         =   "dithering amount"
         FontSizeCaption =   11
         Max             =   100
         Value           =   100
         GradientColorRight=   1703935
         DefaultValue    =   100
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
'Copyright 2000-2019 by Tanner Helland
'Created: 4/October/00
'Last updated: 01/February/18
'Last update: improve progress bar reporting
'
'This dialog allows the user to reduce the number of colors in the current image.  In the future, it would be nice
' to allow palettes loaded from file or selected from an internal swatch manager (and in fact, the code is already
' set up to handle this) but at present, only optimized palettes are supported.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'When loading a palette from file, the pdPalette class handles all the actual parsing
Private m_Palette As pdPalette

'To reduce redraws when interacting with the UI, we manually track changes to the palette filename
Private m_PalettePath As String, m_PaletteFileSize As Long

Private Sub btsAlpha_Click(ByVal buttonIndex As Long)
    ReflowFirstPanel
    UpdatePreview
End Sub

Private Sub ReflowFirstPanel()
    
    'PD currently supports two palette types: RGB and RGBA.  RGBA doesn't require some settings
    ' (like quantizer - as only one is supported - or background color), so we need to reflow the
    ' interface depending on the current palette type.
    Dim rgbPaletteMode As Boolean
    rgbPaletteMode = (btsAlpha.ListIndex = 0)
    
    Dim yOffset As Long, yPadding As Long
    yPadding = Interface.FixDPI(6)
    yOffset = btsAlpha.GetTop + btsAlpha.GetHeight + yPadding
    
    'Only expose quantization method in RGB mode
    btsMethod.Visible = rgbPaletteMode
    If rgbPaletteMode Then
        btsMethod.SetTop yOffset
        yOffset = yOffset + btsMethod.GetHeight + yPadding
    End If
    
    'Palette size and "preserve black and white" are always available
    sldPalette.SetTop yOffset
    yOffset = yOffset + sldPalette.GetHeight + yPadding
    chkPreserveWB.SetTop yOffset
    yOffset = yOffset + chkPreserveWB.GetHeight + yPadding
    
    'Dithering mode and strength are always available
    cboDither(0).SetTop yOffset
    yOffset = yOffset + cboDither(0).GetHeight + yPadding
    
    If sldDitherAmount(0).Visible Then
        sldDitherAmount(0).SetTop yOffset
        yOffset = yOffset + sldDitherAmount(0).GetHeight + yPadding
    End If
    
    'Finally, only expose background color in RGB mode
    clsBackground.Visible = rgbPaletteMode
    If rgbPaletteMode Then
        clsBackground.SetTop yOffset
        yOffset = yOffset + clsBackground.GetHeight + yPadding
    End If
    
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
    SetDitherVisibility Index
    UpdatePreview
End Sub

Private Sub SetDitherVisibility(ByVal srcIndex As Long)
    sldDitherAmount(srcIndex).Visible = (cboDither(srcIndex).ListIndex <> 0)
    If (srcIndex = 0) Then ReflowFirstPanel
End Sub

Private Sub chkMatchAlpha_Click()
    UpdatePreview
End Sub

Private Sub chkPreserveWB_Click()
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
    
    Dim i As Long
    For i = cboDither.lBound To cboDither.UBound
        Palettes.PopulateDitheringDropdown cboDither(i)
        cboDither(i).ListIndex = 6
    Next i
    SetDitherVisibility 0
    
    btsAlpha.AddItem "color only (RGB)", 0
    btsAlpha.AddItem "color and opacity (RGBA)", 1
    btsAlpha.ListIndex = 0
    ReflowFirstPanel
    
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
    
    Dim useRGBAQuantizer As Boolean
    useRGBAQuantizer = cParams.GetBool("useAlpha", False)
    
    'The only quantizer that (currently) supports RGBA output is our homebrew median cut quantizer
    Dim quantMethod As PD_COLOR_QUANTIZE
    If useRGBAQuantizer Then
        quantMethod = PDCQ_MedianCut
    Else
        If Strings.StringsEqual(cParams.GetString("quantizer", "mediancut"), "neuquant", True) Then
            quantMethod = PDCQ_Neuquant
        ElseIf Strings.StringsEqual(cParams.GetString("quantizer", "mediancut"), "wu", True) Then
            quantMethod = PDCQ_Wu
        Else
            quantMethod = PDCQ_MedianCut
        End If
    End If
    
    Dim paletteSize As Long
    paletteSize = cParams.GetLong("palettesize", 256)
    
    Dim preserveWhiteBlack As Boolean
    preserveWhiteBlack = cParams.GetBool("preservewhiteblack", False)
    
    Dim ditherMethod As PD_DITHER_METHOD
    ditherMethod = cParams.GetLong("dithering", 0)
    
    Dim ditherAmount As Single
    ditherAmount = cParams.GetDouble("ditheramount", 100#) * 0.01
    
    Dim finalBackColor As Long
    finalBackColor = cParams.GetLong("backgroundcolor", vbWhite)
    
    Dim tmpSA As SafeArray2D
    EffectPrep.PrepImageData tmpSA, toPreview, dstPic, , , useRGBAQuantizer
    
    If (Not toPreview) Then
        SetProgBarMax workingDIB.GetDIBHeight * 2
        SetProgBarVal 0
        Message "Generating optimal palette..."
    End If
    
    'Some quantization methods require FreeImage.  If FreeImage doesn't exist, fall back to internal PD methods.
    If (quantMethod <> PDCQ_MedianCut) Then
        If (Not ImageFormats.IsFreeImageEnabled()) Then quantMethod = PDCQ_MedianCut
    End If
    
    'If the caller doesn't want transparency, composite the image against the specified backcolor *in advance*.
    If (Not useRGBAQuantizer) Then workingDIB.CompositeBackgroundColor Colors.ExtractRed(finalBackColor), Colors.ExtractGreen(finalBackColor), Colors.ExtractBlue(finalBackColor)
    
    'Branch according to internal or plugin-based quantization methods.  Note that if the user does *NOT* want
    ' dithering, we can use the plugin to apply the palette as well, trimming processing time a bit.
    Dim finalPalette() As RGBQuad, finalPaletteCount As Long
    
    If (quantMethod = PDCQ_MedianCut) Then
    
        'Generate an optimal palette, and if alpha is involved, use it as part of the calculation.
        If useRGBAQuantizer Then
            Palettes.GetOptimizedPaletteIncAlpha workingDIB, finalPalette, paletteSize, , toPreview, workingDIB.GetDIBHeight * 2, 0
        Else
            Palettes.GetOptimizedPalette workingDIB, finalPalette, paletteSize, , toPreview, workingDIB.GetDIBHeight * 2, 0
        End If
        
        'Preserve black and white, as necessary
        If preserveWhiteBlack Then Palettes.EnsureBlackAndWhiteInPalette finalPalette, workingDIB
        
        If (Not toPreview) Then
            SetProgBarVal workingDIB.GetDIBHeight
            Message "Applying new palette to image..."
        End If
        
        'Apply said palette to the image
        If (ditherMethod = PDDM_None) Then
            If useRGBAQuantizer Then
                Palettes.ApplyPaletteToImage_IncAlpha_KDTree workingDIB, finalPalette, toPreview, workingDIB.GetDIBHeight * 2, workingDIB.GetDIBHeight
            Else
                Palettes.ApplyPaletteToImage_KDTree workingDIB, finalPalette, toPreview, workingDIB.GetDIBHeight * 2, workingDIB.GetDIBHeight
            End If
        Else
            If useRGBAQuantizer Then
                Palettes.ApplyPaletteToImage_Dithered_IncAlpha workingDIB, finalPalette, ditherMethod, ditherAmount, toPreview, workingDIB.GetDIBHeight * 2, workingDIB.GetDIBHeight
            Else
                Palettes.ApplyPaletteToImage_Dithered workingDIB, finalPalette, ditherMethod, ditherAmount, toPreview, workingDIB.GetDIBHeight * 2, workingDIB.GetDIBHeight
            End If
        End If
    
    Else
        
        If (Not toPreview) Then
            SetProgBarVal workingDIB.GetDIBHeight
            Message "Applying new palette to image..."
        End If
        
        'Apply all color and transparency changes simultaneously
        Dim fiQuantMode As FREE_IMAGE_QUANTIZE
        If (quantMethod = PDCQ_Wu) Then fiQuantMode = FIQ_WUQUANT Else fiQuantMode = FIQ_NNQUANT
        
        Dim fi_DIB8 As Long
        fi_DIB8 = Plugin_FreeImage.GetFIDib_SpecificColorMode(workingDIB, 8, PDAS_NoAlpha, PDAS_ComplicatedAlpha, , finalBackColor, , paletteSize, , , fiQuantMode)
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
            If preserveWhiteBlack Then Palettes.EnsureBlackAndWhiteInPalette finalPalette, workingDIB
            
            'Apply the generated palette to our target image, using the method requested
            If (finalPaletteCount <> 0) Then
                If (ditherMethod = PDDM_None) Then
                    Palettes.ApplyPaletteToImage_KDTree workingDIB, finalPalette, toPreview, workingDIB.GetDIBHeight * 2, workingDIB.GetDIBHeight
                Else
                    Palettes.ApplyPaletteToImage_Dithered workingDIB, finalPalette, ditherMethod, ditherAmount, toPreview, workingDIB.GetDIBHeight * 2, workingDIB.GetDIBHeight
                End If
            End If
            
        End If
        
    End If
    
    EffectPrep.FinalizeImageData toPreview, dstPic, useRGBAQuantizer
    
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
    
    Dim ditherAmount As Double
    ditherAmount = cParams.GetDouble("ditheramount", 100#) / 100#
    
    Dim matchAlpha As Boolean
    matchAlpha = cParams.GetBool("palettefilematchalpha", False)
    
    Dim tmpSA As SafeArray2D
    EffectPrep.PrepImageData tmpSA, toPreview, dstPic, , , matchAlpha
    
    If (Not toPreview) Then
        SetProgBarMax workingDIB.GetDIBHeight
        SetProgBarVal 0
    End If
    
    'Branch according to internal or plugin-based quantization methods.  Note that if the user does *NOT* want
    ' dithering, we can use the plugin to apply the palette as well, trimming processing time a bit.
    Dim finalPalette() As RGBQuad
    If (m_Palette.GetPaletteColorCount(srcPaletteIndex) > 0) Then
        
        m_Palette.CopyPaletteToArray finalPalette, srcPaletteIndex
        
        If matchAlpha Then
            Palettes.SetPaletteAlphaPremultiplication True, finalPalette
        Else
            Palettes.SetFixedAlpha finalPalette, 255
        End If
        
        If (Not toPreview) Then Message "Applying new palette to image..."
        
        'Apply said palette to the image
        If (ditherMethod = PDDM_None) Then
            If matchAlpha Then
                Palettes.ApplyPaletteToImage_IncAlpha_KDTree workingDIB, finalPalette, toPreview, workingDIB.GetDIBHeight
            Else
                Palettes.ApplyPaletteToImage_KDTree workingDIB, finalPalette, toPreview, workingDIB.GetDIBHeight
            End If
        Else
            If matchAlpha Then
                Palettes.ApplyPaletteToImage_Dithered_IncAlpha workingDIB, finalPalette, ditherMethod, ditherAmount, toPreview, workingDIB.GetDIBHeight
            Else
                Palettes.ApplyPaletteToImage_Dithered workingDIB, finalPalette, ditherMethod, ditherAmount, toPreview, workingDIB.GetDIBHeight
            End If
        End If
        
    End If
    
    EffectPrep.FinalizeImageData toPreview, dstPic, matchAlpha
    
End Sub

Private Sub lstPalettes_Click()
    UpdatePreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub pdFxPreview_ViewportChanged()
    UpdatePreview
End Sub

Private Sub sldDitherAmount_Change(Index As Integer)
    UpdatePreview
End Sub

Private Sub sldPalette_Change()
    UpdatePreview
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
        
        '"Generate optimal" vs "from file"
        .AddParam "mode", btsOptions.ListIndex
        
        'RGB vs RGBA palette
        .AddParam "useAlpha", CBool(btsAlpha.ListIndex = 1)
        
        'Quantizer only matters for RGB palettes but we write it regardless.  (Perhaps in the future
        ' we can support different quantizers for RGBA palettes.)
        Select Case btsMethod.ListIndex
            Case 0
                .AddParam "quantizer", "MedianCut"
            Case 1
                .AddParam "quantizer", "Wu"
            Case 2
                .AddParam "quantizer", "NeuQuant"
        End Select
        
        .AddParam "palettesize", sldPalette.Value
        .AddParam "preservewhiteblack", chkPreserveWB.Value
        .AddParam "backgroundcolor", clsBackground.Color
        
        '"From file" data comes next
        .AddParam "palettefile", txtPalette.Text
        .AddParam "palettefileindex", lstPalettes.ListIndex
        .AddParam "palettefilematchalpha", chkMatchAlpha.Value
        
        'Some options are shared between the two methods
        .AddParam "dithering", cboDither(btsOptions.ListIndex).ListIndex
        .AddParam "ditheramount", sldDitherAmount(btsOptions.ListIndex).Value
        
    End With
    
    GetToolParamString = cParams.GetParamString

End Function

'Use this sub to update the on-screen preview
Private Sub UpdatePreview()
    If cmdBar.PreviewsAllowed Then Me.ApplyPalettizeEffect GetToolParamString, True, pdFxPreview
End Sub
