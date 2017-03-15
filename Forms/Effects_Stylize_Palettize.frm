VERSION 5.00
Begin VB.Form FormPalettize 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Palettize"
   ClientHeight    =   7125
   ClientLeft      =   45
   ClientTop       =   285
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
   ScaleHeight     =   475
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   821
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.pdButtonStrip btsOptions 
      Height          =   615
      Left            =   5880
      TabIndex        =   9
      Top             =   5640
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   1085
   End
   Begin PhotoDemon.pdCommandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   6375
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
      Height          =   5415
      Index           =   0
      Left            =   5880
      TabIndex        =   6
      Top             =   120
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   9551
      Begin PhotoDemon.pdDropDown cboDither 
         Height          =   855
         Left            =   120
         TabIndex        =   12
         Top             =   2640
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   1508
         Caption         =   "dithering"
      End
      Begin PhotoDemon.pdCheckBox chkReduceBleed 
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   3600
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   661
         Caption         =   "reduce color bleed"
         Value           =   0
      End
      Begin PhotoDemon.pdSlider sldPalette 
         Height          =   855
         Left            =   120
         TabIndex        =   7
         Top             =   1320
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   1508
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
         TabIndex        =   8
         Top             =   120
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   1931
         Caption         =   "quantization method"
      End
      Begin PhotoDemon.pdLabel lblWarning 
         Height          =   615
         Left            =   120
         Top             =   4800
         Visible         =   0   'False
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   1085
         Caption         =   ""
         ForeColor       =   4210752
         Layout          =   1
      End
      Begin PhotoDemon.pdCheckBox chkPreserveWB 
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   2085
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   661
         Caption         =   "preserve white and black"
      End
   End
   Begin PhotoDemon.pdContainer pnlQuantize 
      Height          =   5415
      Index           =   1
      Left            =   5880
      TabIndex        =   2
      Top             =   120
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   9551
      Begin PhotoDemon.pdButtonStrip btsAlpha 
         Height          =   1095
         Left            =   120
         TabIndex        =   3
         Top             =   1320
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   1931
         Caption         =   "transparency"
      End
      Begin PhotoDemon.pdSlider sldAlphaCutoff 
         Height          =   855
         Left            =   120
         TabIndex        =   4
         Top             =   2520
         Width           =   6135
         _ExtentX        =   10821
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
         TabIndex        =   5
         Top             =   120
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   1931
         Caption         =   "background color"
      End
   End
End
Attribute VB_Name = "FormPalettize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Color Reduction Form
'Copyright 2000-2017 by Tanner Helland
'Created: 4/October/00
'Last updated: 14/April/14
'Last update: rewrite function against layers; note that this will now flatten a layered image before proceeding
'
'In the original incarnation of PhotoDemon, this was a central part of the project. I have since not used it much
' (since the project is now centered around 24/32bpp imaging), but as it costs nothing to tie into FreeImage's advanced
' color reduction routines, I figure it's worth keeping this dialog around.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

Private m_CompositedImage As pdDIB

Private Sub btsAlpha_Click(ByVal buttonIndex As Long)
    UpdateTransparencyOptions
    UpdatePreview
End Sub

Private Sub UpdateTransparencyOptions()
    sldAlphaCutoff.Visible = CBool(btsAlpha.ListIndex = 1)
End Sub

Private Sub btsMethod_Click(ByVal buttonIndex As Long)
    UpdatePreview
End Sub

Private Sub btsOptions_Click(ByVal buttonIndex As Long)
    UpdateVisiblePanel
End Sub

Private Sub UpdateVisiblePanel()
    Dim i As Long
    For i = pnlQuantize.lBound To pnlQuantize.UBound
        pnlQuantize(i).Visible = CBool(i = btsOptions.ListIndex)
    Next i
End Sub

Private Sub cboDither_Click()
    UpdatePreview
    UpdateColorBleedVisibility
End Sub

Private Sub chkPreserveWB_Click()
    UpdatePreview
End Sub

Private Sub chkReduceBleed_Click()
    UpdatePreview
End Sub

Private Sub clsBackground_ColorChanged()
    UpdatePreview
End Sub

'OK button
Private Sub cmdBar_OKClick()
    Process "Palettize", , GetToolParamString, UNDO_LAYER
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

Private Sub Form_Load()
    
    'Suspend previews until the dialog has been fully initialized
    cmdBar.MarkPreviewStatus False
    
    btsOptions.AddItem "basic", 0
    btsOptions.AddItem "advanced", 1
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
    
    cboDither.Clear
    cboDither.AddItem "None", 0
    cboDither.AddItem "Ordered (Bayer 4x4)", 1
    cboDither.AddItem "Ordered (Bayer 8x8)", 2
    cboDither.AddItem "False (Fast) Floyd-Steinberg", 3
    cboDither.AddItem "Genuine Floyd-Steinberg", 4
    cboDither.AddItem "Jarvis, Judice, and Ninke", 5
    cboDither.AddItem "Stucki", 6
    cboDither.AddItem "Burkes", 7
    cboDither.AddItem "Sierra-3", 8
    cboDither.AddItem "Two-Row Sierra", 9
    cboDither.AddItem "Sierra Lite", 10
    cboDither.AddItem "Atkinson / Classic Macintosh", 11
    cboDither.ListIndex = 6
    UpdateColorBleedVisibility
    
    'Apply translations and visual themes
    ApplyThemeAndTranslations Me
    cmdBar.MarkPreviewStatus True
    UpdatePreview
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Function GetToolParamString() As String

    Dim cParams As pdParamXML
    Set cParams = New pdParamXML
    
    With cParams
        Select Case btsMethod.ListIndex
            Case 0
                .AddParam "IndexedColors_Method", "MedianCut"
            Case 1
                .AddParam "IndexedColors_Method", "Wu"
            Case 2
                .AddParam "IndexedColors_Method", "NeuQuant"
        End Select
        
        .AddParam "IndexedColors_PaletteSize", sldPalette.Value
        .AddParam "IndexedColors_PreserveWhiteBlack", CBool(chkPreserveWB.Value)
        .AddParam "IndexedColors_Dithering", cboDither.ListIndex
        .AddParam "IndexedColors_ReduceBleed", CBool(chkReduceBleed.Value)
        .AddParam "IndexedColors_BackgroundColor", clsBackground.Color
        
        Select Case btsAlpha.ListIndex
            Case 0
                .AddParam "IndexedColors_Alpha", "full"
            Case 1
                .AddParam "IndexedColors_Alpha", "binary"
            Case 2
                .AddParam "IndexedColors_Alpha", "none"
        End Select
        
        .AddParam "IndexedColors_AlphaCutoff", sldAlphaCutoff.Value
        
    End With
    
    GetToolParamString = cParams.GetParamString

End Function

'Automatic 8-bit color reduction.  Some option combinations require the FreeImage plugin.
Public Sub ApplyPalettizeEffect(ByVal toolParams As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)
    
    'Parse the parameter string and determine concrete values for our color conversion
    Dim cParams As pdParamXML
    Set cParams = New pdParamXML
    cParams.SetParamString toolParams
    
    Dim quantMethod As PD_COLOR_QUANTIZE
    If (StrComp(LCase$(cParams.GetString("IndexedColors_Method", "mediancut")), "neuquant", vbBinaryCompare) = 0) Then
        quantMethod = PDCQ_Neuquant
    ElseIf (StrComp(LCase$(cParams.GetString("IndexedColors_Method", "mediancut")), "wu", vbBinaryCompare) = 0) Then
        quantMethod = PDCQ_Wu
    Else
        quantMethod = PDCQ_MedianCut
    End If
    
    Dim paletteSize As Long
    paletteSize = cParams.GetLong("IndexedColors_PaletteSize", 256)
    
    Dim preserveWhiteBlack As Boolean
    preserveWhiteBlack = cParams.GetBool("IndexedColors_PreserveWhiteBlack", False)
    
    Dim DitherMethod As PD_DITHER_METHOD
    DitherMethod = cParams.GetLong("IndexedColors_Dithering", 0)
    
    Dim reduceBleed As Boolean
    reduceBleed = cParams.GetBool("IndexedColors_ReduceBleed", False)
    
    Dim finalBackColor As Long
    finalBackColor = cParams.GetLong("IndexedColors_BackgroundColor", vbWhite)
    
    Dim outputAlphaMode As PD_ALPHA_STATUS
    If (StrComp(LCase$(cParams.GetString("IndexedColors_Alpha", "full")), "full", vbBinaryCompare) = 0) Then
        outputAlphaMode = PDAS_ComplicatedAlpha
    ElseIf (StrComp(LCase$(cParams.GetString("IndexedColors_Alpha", "full")), "binary", vbBinaryCompare) = 0) Then
        outputAlphaMode = PDAS_BinaryAlpha
    Else
        outputAlphaMode = PDAS_NoAlpha
    End If
    
    Dim alphaCutoff As Long
    alphaCutoff = cParams.GetLong("IndexedColors_AlphaCutoff", 64)
    
    Dim tmpSA As SAFEARRAY2D
    PrepImageData tmpSA, toPreview, pdFxPreview
    
    If (Not toPreview) Then
        SetProgBarMax 3
        SetProgBarVal 1
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
    Dim usePDToApplyPalette As Boolean: usePDToApplyPalette = True
    Dim finalPalette() As RGBQUAD, finalPaletteCount As Long
    
    If (quantMethod = PDCQ_MedianCut) Then
    
        'Resize the target DIB to a smaller size
        Dim smallDIB As pdDIB
        DIBs.ResizeDIBByPixelCount workingDIB, smallDIB, 50000
        
        'Generate an optimal palette
        Palettes.GetOptimizedPalette smallDIB, finalPalette, paletteSize
        
        'Preserve black and white, as necessary
        If preserveWhiteBlack Then Palettes.EnsureBlackAndWhiteInPalette finalPalette, smallDIB
        
        If (Not toPreview) Then
            SetProgBarVal 2
            Message "Applying new palette to image..."
        End If
        
        'Apply said palette to the image
        If (DitherMethod = PDDM_None) Then
            Palettes.ApplyPaletteToImage_SysAPI workingDIB, finalPalette
        Else
            Palettes.ApplyPaletteToImage_Dithered workingDIB, finalPalette, DitherMethod, reduceBleed
        End If
    
    Else
        
        'Apply all color and transparency changes simultaneously
        Dim fiQuantMode As FREE_IMAGE_QUANTIZE
        If (quantMethod = PDCQ_Wu) Then fiQuantMode = FIQ_WUQUANT Else fiQuantMode = FIQ_NNQUANT
        
        Dim fi_DIB8 As Long
        fi_DIB8 = Plugin_FreeImage.GetFIDib_SpecificColorMode(workingDIB, 8, outputAlphaMode, currentAlphaState, alphaCutoff, finalBackColor, , paletteSize, , , fiQuantMode)
        FreeImage_FlipVertically fi_DIB8
        
        If (Not toPreview) Then
            SetProgBarVal 2
            Message "Applying new palette to image..."
        End If
        
        'If the caller does *not* want dithering, copy the (already palettized) FreeImage DIB over our
        ' original DIB.
        If (DitherMethod = PDDM_None) And (Not preserveWhiteBlack) Then
        
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
            ReDim Preserve finalPalette(0 To paletteSize - 1) As RGBQUAD
            FreeImage_Unload fi_DIB8
            
            'Preserve black and white, as necessary
            If preserveWhiteBlack Then Palettes.EnsureBlackAndWhiteInPalette finalPalette, smallDIB
            
            'Apply the generated palette to our target image, using the method requested
            If (finalPaletteCount <> 0) Then
                If (DitherMethod = PDDM_None) Then
                    Palettes.ApplyPaletteToImage_SysAPI workingDIB, finalPalette
                Else
                    Palettes.ApplyPaletteToImage_Dithered workingDIB, finalPalette, DitherMethod, reduceBleed
                End If
            End If
            
        End If
        
    End If
    
    EffectPrep.FinalizeImageData toPreview, pdFxPreview
    
End Sub

'Use this sub to update the on-screen preview
Private Sub UpdatePreview()
    If cmdBar.PreviewsAllowed Then ApplyPalettizeEffect GetToolParamString, True, pdFxPreview
End Sub

Private Sub UpdateColorBleedVisibility()
    chkReduceBleed.Visible = CBool(cboDither.ListIndex <> 0)
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
