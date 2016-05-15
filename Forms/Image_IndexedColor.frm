VERSION 5.00
Begin VB.Form FormReduceColors 
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
   Begin PhotoDemon.pdSlider sldPalette 
      Height          =   855
      Left            =   6000
      TabIndex        =   3
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
      Left            =   6000
      TabIndex        =   2
      Top             =   120
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   1931
      Caption         =   "quantization method"
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
   Begin PhotoDemon.pdLabel lblWarning 
      Height          =   615
      Left            =   6000
      Top             =   5640
      Visible         =   0   'False
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   1085
      Caption         =   ""
      ForeColor       =   4210752
      Layout          =   1
   End
   Begin PhotoDemon.pdButtonStrip btsAlpha 
      Height          =   1095
      Left            =   6000
      TabIndex        =   4
      Top             =   3480
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   1931
      Caption         =   "transparency"
   End
   Begin PhotoDemon.pdSlider sldAlphaCutoff 
      Height          =   855
      Left            =   6000
      TabIndex        =   5
      Top             =   4680
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
      Left            =   6000
      TabIndex        =   6
      Top             =   2280
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   1931
      Caption         =   "background color"
   End
End
Attribute VB_Name = "FormReduceColors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Color Reduction Form
'Copyright 2000-2016 by Tanner Helland
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

Private Sub clsBackground_ColorChanged()
    UpdatePreview
End Sub

'OK button
Private Sub cmdBar_OKClick()
    Process "Reduce colors", , GetToolParamString, UNDO_IMAGE
End Sub

Private Sub cmdBar_ResetClick()
    btsMethod.ListIndex = 0
    sldPalette.Value = 256
    clsBackground.Color = vbWhite
    btsAlpha.ListIndex = 0
    sldAlphaCutoff.Value = sldAlphaCutoff.NotchValueCustom
    UpdatePreview
End Sub

Private Sub Form_Activate()
    UpdatePreview
End Sub

Private Sub Form_Load()

    'Suspend previews until the dialog has been fully initialized
    cmdBar.MarkPreviewStatus False
    
    btsMethod.AddItem "Xiaolin Wu", 0
    btsMethod.AddItem "NeuQuant neural network", 1
    btsMethod.ListIndex = 0
    
    btsAlpha.AddItem "full", 0
    btsAlpha.AddItem "binary", 1
    btsAlpha.AddItem "none", 2
    btsAlpha.ListIndex = 0
    UpdateTransparencyOptions
    
    'At present, FreeImage is required for conversion to indexed color
    If (Not g_ImageFormats.FreeImageEnabled) Then
    
        btsMethod.Enabled = False
        sldPalette.Enabled = False
        
        lblWarning.Caption = g_Language.TranslateMessage("The FreeImage plugin is missing.  Please install it if you wish to use this tool.")
        lblWarning.ForeColor = g_Themer.GetGenericUIColor(UI_UniversalErrorRed)
        lblWarning.UseCustomForeColor = True
        lblWarning.Visible = True
    
    'If the current image has more than one layer, warn the user that this action will flatten the image.
    Else
        If (pdImages(g_CurrentImage).GetNumOfLayers > 1) Then
            lblWarning.Caption = g_Language.TranslateMessage("Note: this operation will flatten the image before converting it to indexed color mode.")
            lblWarning.UseCustomForeColor = False
            lblWarning.Visible = True
        End If
    End If
    
    'Make a copy of the composited image; it takes time to composite layers, so we don't want to redo this except
    ' when absolutely necessary.
    Set m_CompositedImage = New pdDIB
    pdImages(g_CurrentImage).GetCompositedImage m_CompositedImage, True
    pdFxPreview.NotifyNonStandardSource m_CompositedImage.GetDIBWidth, m_CompositedImage.GetDIBHeight
    
    If (Not g_ImageFormats.FreeImageEnabled) Then Interface.ShowDisabledPreviewImage pdFxPreview
    
    'Apply translations and visual themes
    ApplyThemeAndTranslations Me
    cmdBar.MarkPreviewStatus True
    
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
                .AddParam "IndexedColors_Method", "Wu"
            Case 1
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

'Automatic 8-bit color reduction via the FreeImage DLL.
Public Sub ReduceImageColors_Auto(ByVal toolParams As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)
    
    'If this is not a preview, and a selection is active on the main image, remove it.
    If (Not toPreview) And pdImages(g_CurrentImage).selectionActive Then
        pdImages(g_CurrentImage).selectionActive = False
        pdImages(g_CurrentImage).mainSelection.lockRelease
    End If
    
    'Color reduction only works on a flat copy of the image, so retrieve a composited version now.
    If toPreview Then
        
        'Because the user can change the preview viewport, we can't guarantee that the preview region hasn't changed
        ' since the last preview.  Prep a new preview now.
        Dim tmpSafeArray As SAFEARRAY2D
        FastDrawing.PreviewNonStandardImage tmpSafeArray, m_CompositedImage, pdFxPreview, True
        
    'If this is not a preview, flatten the image before proceeding further
    Else
        SetProgBarMax 3
        SetProgBarVal 1
        Message "Flattening image..."
        Layer_Handler.FlattenImage
    End If
    
    'At this point, we have two potential sources of our temporary DIB:
    ' 1) During a preview, the global workingDIB object contains a section of the image relevant to the
    '     preview window.
    ' 2) During the processing of a full image, pdImages(g_CurrentImage) has what we need (the flattened image).
    '
    'To simplify the code from here, we are going to conditionally copy the current flattened image into
    ' the global workingLayer DIB.  That way, we can use the same code path regardless of previews or
    ' actual processing.
    If (Not toPreview) Then
        Set workingDIB = New pdDIB
        workingDIB.CreateFromExistingDIB pdImages(g_CurrentImage).GetLayerByIndex(0).layerDIB
    End If
    
    'Make sure we found the FreeImage plug-in when the program was loaded
    If g_ImageFormats.FreeImageEnabled Then
        
        If (Not toPreview) Then
            SetProgBarVal 2
            Message "Indexing colors..."
        End If
        
        'Parse the parameter string and determine concrete values for our color conversion
        Dim cParams As pdParamXML
        Set cParams = New pdParamXML
        cParams.SetParamString toolParams
        
        Dim quantMethod As FREE_IMAGE_QUANTIZE
        If (StrComp(LCase$(cParams.GetString("IndexedColors_Method", "wu")), "neuquant", vbBinaryCompare) = 0) Then
            quantMethod = FIQ_NNQUANT
        Else
            quantMethod = FIQ_WUQUANT
        End If
        
        Dim paletteSize As Long
        paletteSize = cParams.GetLong("IndexedColors_PaletteSize", 256)
        
        Dim backgroundColor As Long
        backgroundColor = cParams.GetLong("IndexedColors_BackgroundColor", vbWhite)
        
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
        
        'Convert our current DIB to a FreeImage-type DIB
        Dim fi_DIB8 As Long
        fi_DIB8 = Plugin_FreeImage.GetFIDib_SpecificColorMode(workingDIB, 8, outputAlphaMode, PDAS_ComplicatedAlpha, alphaCutoff, backgroundColor, , paletteSize, , , quantMethod)
        FreeImage_FlipVertically fi_DIB8
        
        'Convert that DIB to 32-bpp
        Dim fi_DIB As Long
        fi_DIB = FreeImage_ConvertTo32Bits(fi_DIB8)
        FreeImage_Unload fi_DIB8
        
        'Paint the result to workingDIB
        workingDIB.ResetDIB 0
        Plugin_FreeImage.PaintFIDibToPDDib workingDIB, fi_DIB, 0, 0, workingDIB.GetDIBWidth, workingDIB.GetDIBHeight
        FreeImage_Unload fi_DIB
        
        'If this is not a preview, overwrite the current active layer with the quantized FreeImage data.
        If (Not toPreview) Then
            
            SetProgBarVal 3
            pdImages(g_CurrentImage).GetLayerByIndex(0).layerDIB.CreateFromExistingDIB workingDIB
            pdImages(g_CurrentImage).GetLayerByIndex(0).layerDIB.ConvertTo32bpp
            
            'Notify the parent image of these changes
            pdImages(g_CurrentImage).NotifyImageChanged UNDO_LAYER, 0
            pdImages(g_CurrentImage).NotifyImageChanged UNDO_IMAGE
                
        End If
            
        'If this is a preview, draw the new image to the picture box and exit.  Otherwise, render the new main image DIB.
        If toPreview Then
            FinalizeNonstandardPreview dstPic, True
        Else
            Viewport_Engine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), FormMain.mainCanvas(0)
            SetProgBarVal 0
            ReleaseProgressBar
            Message "Image successfully quantized to %1 unique colors. ", paletteSize
        End If
        
    Else
        PDMsgBox "The FreeImage interface plug-in (FreeImage.dll) was marked as missing or disabled upon program initialization." & vbCrLf & vbCrLf & "To enable support for this feature, please copy the FreeImage.dll file into the plug-in directory and reload the program.", vbExclamation + vbOKOnly + vbApplicationModal, " FreeImage Interface Error"
        Exit Sub
    End If
    
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
