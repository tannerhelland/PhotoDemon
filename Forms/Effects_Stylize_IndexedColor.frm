VERSION 5.00
Begin VB.Form FormIndexedColor 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Indexed color"
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
   End
   Begin PhotoDemon.pdCommandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   6375
      Width           =   12315
   End
   Begin PhotoDemon.pdFxPreviewCtl pdFxPreview 
      Height          =   6105
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5625
   End
   Begin PhotoDemon.pdContainer pnlQuantize 
      Height          =   5295
      Index           =   0
      Left            =   5880
      TabIndex        =   6
      Top             =   120
      Width           =   6375
      Begin PhotoDemon.pdSlider sldPalette 
         Height          =   855
         Left            =   120
         TabIndex        =   7
         Top             =   1320
         Width           =   6135
      End
      Begin PhotoDemon.pdButtonStrip btsMethod 
         Height          =   1095
         Left            =   120
         TabIndex        =   8
         Top             =   120
         Width           =   6135
      End
      Begin PhotoDemon.pdLabel lblWarning 
         Height          =   615
         Left            =   120
         Top             =   4800
         Visible         =   0   'False
         Width           =   6015
      End
   End
   Begin PhotoDemon.pdContainer pnlQuantize 
      Height          =   5295
      Index           =   1
      Left            =   5880
      TabIndex        =   2
      Top             =   120
      Width           =   6375
      Begin PhotoDemon.pdButtonStrip btsAlpha 
         Height          =   1095
         Left            =   120
         TabIndex        =   3
         Top             =   1320
         Width           =   6135
      End
      Begin PhotoDemon.pdSlider sldAlphaCutoff 
         Height          =   855
         Left            =   120
         TabIndex        =   4
         Top             =   2520
         Width           =   6135
      End
      Begin PhotoDemon.pdColorSelector clsBackground 
         Height          =   1095
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   6135
      End
   End
End
Attribute VB_Name = "FormIndexedColor"
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

Private Sub clsBackground_ColorChanged()
    UpdatePreview
End Sub

'OK button
Private Sub cmdBar_OKClick()
    Process "Reduce colors", , GetToolParamString, UNDO_IMAGE
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
    btsMethod.AddItem "NeuQuant neural network", 2
    btsMethod.ListIndex = 0
    
    btsAlpha.AddItem "full", 0
    btsAlpha.AddItem "binary", 1
    btsAlpha.AddItem "none", 2
    btsAlpha.ListIndex = 0
    UpdateTransparencyOptions
    
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
Public Sub ReduceImageColors_Auto(ByVal toolParams As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)
    
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
    
    'Branch according to internal or plugin-based quantization methods.  Note that if the user does *NOT* want
    ' dithering, we can use the plugin to apply the palette as well, trimming processing time a bit.
    Dim usePDToApplyPalette As Boolean: usePDToApplyPalette = True
    
    If (quantMethod = PDCQ_MedianCut) Then
    
        'Resize the target DIB to a smaller size
        Dim smallDIB As pdDIB
        DIB_Support.ResizeDIBByPixelCount workingDIB, smallDIB, 50000
        
        'Generate an optimal palette
        Dim finalPalette() As RGBQUAD
        Palettes.GetOptimizedPalette smallDIB, finalPalette, paletteSize
        
        If (Not toPreview) Then
            SetProgBarVal 2
            Message "Applying new palette to image..."
        End If
        
        'Apply said palette to the image
        Palettes.ApplyPaletteToImage_SysAPI workingDIB, finalPalette
    
    Else
        
        'Apply all color and transparency changes simultaneously
        Dim fiQuantMode As FREE_IMAGE_QUANTIZE
        If (quantMethod = PDCQ_Wu) Then fiQuantMode = FIQ_WUQUANT Else fiQuantMode = FIQ_NNQUANT
        
        Dim fi_DIB8 As Long
        fi_DIB8 = Plugin_FreeImage.GetFIDib_SpecificColorMode(workingDIB, 8, outputAlphaMode, PDAS_ComplicatedAlpha, alphaCutoff, finalBackColor, , paletteSize, , , fiQuantMode)
        FreeImage_FlipVertically fi_DIB8
        
        'If the caller does *not* want dithering, copy the (already palettized) FreeImage DIB over our
        ' original DIB.
        
        'TODO: implement dithering settings
        
        'Retrieve the generated palette, then free the FreeImage source
        'Dim finalPaletteCount As Long, finalPalette() As RGBQUAD
        'finalPaletteCount = Plugin_FreeImage.GetFreeImagePalette(fi_DIB8, finalPalette)
        'FreeImage_Unload fi_DIB8
        
        'Apply the generated palette to our target image, using the method requested
        'If (finalPaletteCount <> 0) Then
        '    Palettes.ApplyPaletteToImage_SysAPI workingDIB, finalPalette
        'Else
        '    Debug.Print "FAILURE!"
        'End If
        
        If (Not toPreview) Then
            SetProgBarVal 2
            Message "Applying new palette to image..."
        End If
        
        'Convert that DIB to 32-bpp
        Dim fi_DIB As Long
        fi_DIB = FreeImage_ConvertTo32Bits(fi_DIB8)
        FreeImage_Unload fi_DIB8
        
        'Paint the result to workingDIB
        workingDIB.ResetDIB 0
        Plugin_FreeImage.PaintFIDibToPDDib workingDIB, fi_DIB, 0, 0, workingDIB.GetDIBWidth, workingDIB.GetDIBHeight
        FreeImage_Unload fi_DIB
        
    End If
    
    FastDrawing.FinalizeImageData toPreview, pdFxPreview
    
End Sub

'Use this sub to update the on-screen preview
Private Sub UpdatePreview()
    If cmdBar.PreviewsAllowed Then ReduceImageColors_Auto GetToolParamString, True, pdFxPreview
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
