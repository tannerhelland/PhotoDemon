VERSION 5.00
Begin VB.Form dialog_ExportBMP 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " BMP export options"
   ClientHeight    =   6540
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
   ScaleHeight     =   436
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   873
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.pdCheckBox chkColorCount 
      Height          =   375
      Left            =   6000
      TabIndex        =   6
      Top             =   4200
      Width           =   6975
      _ExtentX        =   7858
      _ExtentY        =   661
      Caption         =   "restrict palette size"
      Value           =   0
   End
   Begin PhotoDemon.pdColorSelector clsBackground 
      Height          =   975
      Left            =   5880
      TabIndex        =   9
      Top             =   1860
      Width           =   7095
      _ExtentX        =   15690
      _ExtentY        =   1720
      Caption         =   "background color"
   End
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   375
      Index           =   0
      Left            =   9360
      Top             =   4740
      Width           =   3615
      _ExtentX        =   9340
      _ExtentY        =   661
      Caption         =   "unique colors"
   End
   Begin PhotoDemon.pdSlider sldColorCount 
      Height          =   375
      Left            =   6240
      TabIndex        =   7
      Top             =   4680
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   661
      Min             =   2
      Max             =   256
      Value           =   256
      NotchPosition   =   2
      NotchValueCustom=   256
   End
   Begin PhotoDemon.pdButtonStrip btsDepthRGB 
      Height          =   1095
      Left            =   5880
      TabIndex        =   4
      Top             =   3000
      Width           =   7095
      _ExtentX        =   15690
      _ExtentY        =   1931
      Caption         =   "depth"
   End
   Begin PhotoDemon.pdButtonStrip btsColorModel 
      Height          =   1095
      Left            =   5880
      TabIndex        =   3
      Top             =   120
      Width           =   7095
      _ExtentX        =   15690
      _ExtentY        =   1931
      Caption         =   "color model"
   End
   Begin PhotoDemon.pdCommandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   2
      Top             =   5790
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   1323
   End
   Begin PhotoDemon.pdCheckBox chkRLE 
      Height          =   375
      Left            =   6000
      TabIndex        =   0
      Top             =   5280
      Width           =   6975
      _ExtentX        =   7435
      _ExtentY        =   661
      Caption         =   "use RLE compression"
      Value           =   0
   End
   Begin PhotoDemon.pdButtonStrip btsDepthGrayscale 
      Height          =   1095
      Left            =   5880
      TabIndex        =   5
      Top             =   3000
      Width           =   7095
      _ExtentX        =   15690
      _ExtentY        =   1931
      Caption         =   "depth"
   End
   Begin PhotoDemon.pdCheckBox chkPremultiplyAlpha 
      Height          =   375
      Left            =   6000
      TabIndex        =   8
      Top             =   1320
      Width           =   6855
      _ExtentX        =   15478
      _ExtentY        =   661
      Caption         =   "premultiply alpha"
      Value           =   0
   End
   Begin PhotoDemon.pdCheckBox chkFlipRows 
      Height          =   375
      Left            =   6000
      TabIndex        =   10
      Top             =   1320
      Width           =   6975
      _ExtentX        =   15478
      _ExtentY        =   661
      Caption         =   "flip row order (top-down)"
      Value           =   0
   End
   Begin PhotoDemon.pdFxPreviewCtl pdFxPreview 
      Height          =   5625
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9922
   End
   Begin PhotoDemon.pdCheckBox chk16555 
      Height          =   375
      Left            =   6000
      TabIndex        =   1
      Top             =   4200
      Width           =   6975
      _ExtentX        =   15478
      _ExtentY        =   661
      Caption         =   "use legacy 15-bit encoding (X1-R5-G5-B5)"
      Value           =   0
   End
End
Attribute VB_Name = "dialog_ExportBMP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Bitmap export dialog
'Copyright 2012-2016 by Tanner Helland
'Created: 11/December/12
'Last updated: 16/March/16
'Last update: repurpose old color-depth dialog into a BMP-specific one
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

'The ShowDialog routine presents the user with this form.
Public Sub ShowDialog(Optional ByRef srcImage As pdImage = Nothing)

    'Provide a default answer of "cancel" (in the event that the user clicks the "x" button in the top-right)
    m_UserDialogAnswer = vbCancel
    
    Message "Waiting for user to specify export options... "
    
    btsColorModel.AddItem "auto", 0
    btsColorModel.AddItem "color + transparency", 1
    btsColorModel.AddItem "color only", 2
    btsColorModel.AddItem "grayscale", 3
    
    btsDepthRGB.AddItem "32-bpp XRGB (X8-R8-G8-B8)", 0
    btsDepthRGB.AddItem "24-bpp RGB (R8-G8-B8)", 1
    btsDepthRGB.AddItem "16-bpp (R5-G6-B5)", 2
    btsDepthRGB.AddItem "8-bpp (indexed)", 3
    
    btsDepthGrayscale.AddItem "8-bpp (256 shades)", 0
    btsDepthGrayscale.AddItem "4-bpp (16 shades)", 1
    btsDepthGrayscale.AddItem "1-bpp (monochrome)", 2
    
    'Prep a preview (if any)
    Set m_SrcImage = srcImage
    If Not (m_SrcImage Is Nothing) Then
        m_SrcImage.GetCompositedImage m_CompositedImage, True
        pdFxPreview.NotifyNonStandardSource m_CompositedImage.GetDIBWidth, m_CompositedImage.GetDIBHeight
    End If
    If (Not g_ImageFormats.FreeImageEnabled) Or (m_SrcImage Is Nothing) Then Interface.ShowDisabledPreviewImage pdFxPreview
    
    'Update the preview
    UpdateAllVisibility
    UpdatePreviewSource
    UpdatePreview
    
    'Apply translations and visual themes
    ApplyThemeAndTranslations Me
    
    'Display the dialog
    ShowPDDialog vbModal, Me, True

End Sub

Private Sub btsColorModel_Click(ByVal buttonIndex As Long)
    UpdateAllVisibility
    UpdatePreviewSource
    UpdatePreview
End Sub

Private Sub UpdateAllVisibility()

    Select Case btsColorModel.ListIndex
    
        'Auto
        Case 0
            btsDepthRGB.Visible = False
            btsDepthGrayscale.Visible = False
            chkPremultiplyAlpha.Visible = False
            clsBackground.Visible = False
            chkFlipRows.Visible = False
            
        'RGBA
        Case 1
            btsDepthRGB.Visible = False
            btsDepthGrayscale.Visible = False
            chkPremultiplyAlpha.Visible = True
            clsBackground.Visible = False
            chkFlipRows.Visible = True
        
        'RGB
        Case 2
            btsDepthRGB.Visible = True
            btsDepthGrayscale.Visible = False
            chkPremultiplyAlpha.Visible = False
            clsBackground.Visible = True
            chkFlipRows.Visible = True
        
        'Grayscale
        Case 3
            btsDepthRGB.Visible = False
            btsDepthGrayscale.Visible = True
            chkPremultiplyAlpha.Visible = False
            clsBackground.Visible = True
            chkFlipRows.Visible = True
    
    End Select
    
    EvaluateDepthRGBVisibility

End Sub

Private Sub EvaluateDepthRGBVisibility()
    If (Not btsDepthRGB.Visible) Then
        chk16555.Visible = False
        SetGroupVisibility_IndexedColor False
    Else
        Select Case btsDepthRGB.ListIndex
        
            '32-bpp XRGB
            Case 0
                chk16555.Visible = False
                SetGroupVisibility_IndexedColor False
                
            '24-bpp
            Case 1
                chk16555.Visible = False
                SetGroupVisibility_IndexedColor False
            
            '16-bpp
            Case 2
                chk16555.Visible = True
                SetGroupVisibility_IndexedColor False
            
            '8-bpp
            Case 3
                chk16555.Visible = False
                SetGroupVisibility_IndexedColor True
        
        End Select
    End If
End Sub

Private Sub SetGroupVisibility_IndexedColor(ByVal vState As Boolean)
    chkRLE.Visible = vState
    chkColorCount.Visible = vState
    sldColorCount.Visible = vState
    lblTitle(0).Visible = vState
End Sub

Private Sub btsDepthGrayscale_Click(ByVal buttonIndex As Long)
    UpdatePreviewSource
    UpdatePreview
End Sub

Private Sub btsDepthRGB_Click(ByVal buttonIndex As Long)
    EvaluateDepthRGBVisibility
    UpdatePreviewSource
    UpdatePreview
End Sub

Private Sub chk16555_Click()
    UpdatePreviewSource
    UpdatePreview
End Sub

Private Sub chkColorCount_Click()
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
    m_UserDialogAnswer = vbOK
    Me.Visible = False
End Sub

Private Sub cmdBar_ReadCustomPresetData()
    UpdateAllVisibility
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub cmdBar_ResetClick()
    chkPremultiplyAlpha.Value = vbUnchecked
    chk16555.Value = vbUnchecked
    chkColorCount.Value = vbUnchecked
    chkRLE = vbUnchecked
    chkFlipRows.Value = vbUnchecked
    sldColorCount.Value = 256
    btsDepthGrayscale.ListIndex = 0
    btsDepthRGB.ListIndex = 1
    btsColorModel.ListIndex = 0
    clsBackground.Color = vbWhite
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
    Plugin_FreeImage.ReleasePreviewCache m_FIHandle
End Sub

Private Function GetExportParamString() As String

    Dim cParams As pdParamXML
    Set cParams = New pdParamXML
    
    'Convert the color depth option buttons into a usable numeric value
    Dim outputDepth As String
    
    Select Case btsColorModel.ListIndex
        
        'Auto
        Case 0
            outputDepth = "Auto"
        
        'RGBA
        Case 1
            outputDepth = "32"
            cParams.AddParam "BMPUseXRGB", False
            cParams.AddParam "BMPPremultiplyAlpha", CBool(chkPremultiplyAlpha.Value)
        
        'RGB
        Case 2
            Select Case btsDepthRGB.ListIndex
                
                '32-bpp XRGB
                Case 0
                    outputDepth = "32"
                    cParams.AddParam "BMPUseXRGB", True
                    cParams.AddParam "BMPPremultiplyAlpha", False
                
                '24-bpp
                Case 1
                    outputDepth = "24"
                
                '16-bpp
                Case 2
                    outputDepth = "16"
                
                '8-bpp
                Case 3
                    outputDepth = "8"
                
            End Select
        
        'Grayscale
        Case 3
            Select Case btsDepthGrayscale.ListIndex
                
                '8-bpp
                Case 0
                    outputDepth = "8"
                
                '4-bpp
                Case 1
                    outputDepth = "4"
                
                '1-bpp
                Case 2
                    outputDepth = "1"
                
            End Select
    
    End Select
    
    cParams.AddParam "BMPColorDepth", outputDepth
    cParams.AddParam "BMPRLECompression", CBool(chkRLE.Value)
    cParams.AddParam "BMPForceGrayscale", CBool(btsColorModel.ListIndex = 3)
    cParams.AddParam "BMP16bpp555", CBool(chk16555.Value)
    If CBool(chkColorCount.Value) And (btsColorModel.ListIndex <> 3) Then cParams.AddParam "BMPIndexedColorCount", sldColorCount.Value Else cParams.AddParam "BMPIndexedColorCount", 256
    cParams.AddParam "BMPBackgroundColor", clsBackground.Color
    cParams.AddParam "BMPFlipRowOrder", CBool(chkFlipRows.Value)
    
    GetExportParamString = cParams.GetParamString
    
End Function

Private Sub pdFxPreview_ViewportChanged()
    UpdatePreviewSource
    UpdatePreview
End Sub

'When a parameter changes that requires a new source DIB for the preview (e.g. changing the background composite color),
' call this function to generate a new preview DIB.  Note that you *do not* need to call this function for format-specific
' changes (like quality, subsampling, etc).
Private Sub UpdatePreviewSource()

    If (Not (m_CompositedImage Is Nothing)) Then
        
        'Because the user can change the preview viewport, we can't guarantee that the preview region hasn't changed
        ' since the last preview.  Prep a new preview now.
        Dim tmpSafeArray As SAFEARRAY2D
        FastDrawing.PreviewNonStandardImage tmpSafeArray, m_CompositedImage, pdFxPreview, True
        
        'Convert the DIB to a FreeImage-compatible handle, at a color-depth that matches the current settings.
        ' (Note that we can completely skip this step for the "Auto" depth mode.)
        Dim prvColorDepth As Long, forceGrayscale As Boolean
        forceGrayscale = False
        
        If (btsColorModel.ListIndex = 0) Then
            prvColorDepth = 32
        Else
            
            If (btsColorModel.ListIndex = 1) Then
                prvColorDepth = 32
            ElseIf (btsColorModel.ListIndex = 2) Then
                Select Case btsDepthRGB.ListIndex
                    Case 0, 1
                        prvColorDepth = 24
                    Case 2
                        prvColorDepth = 16
                    Case 3
                        prvColorDepth = 8
                End Select
            Else
                forceGrayscale = True
                Select Case btsDepthGrayscale.ListIndex
                    Case 0
                        prvColorDepth = 8
                    Case 1
                        prvColorDepth = 4
                    Case 2
                        prvColorDepth = 1
                End Select
            End If
            
        End If
        
        Dim BMP16bpp555 As Boolean
        BMP16bpp555 = CBool(chk16555.Value)
        
        Dim BMPIndexedColorCount As Long
        If CBool(chkColorCount.Value) And (Not forceGrayscale) Then
            If sldColorCount.IsValid Then BMPIndexedColorCount = sldColorCount.Value Else BMPIndexedColorCount = 256
        Else
            BMPIndexedColorCount = 256
        End If
        
        If (m_FIHandle <> 0) Then Plugin_FreeImage.ReleaseFreeImageObject m_FIHandle
        If prvColorDepth = 32 Then
            m_FIHandle = Plugin_FreeImage.GetFIDib_SpecificColorMode(workingDIB, prvColorDepth, PDAS_ComplicatedAlpha, PDAS_ComplicatedAlpha)
        Else
            m_FIHandle = Plugin_FreeImage.GetFIDib_SpecificColorMode(workingDIB, prvColorDepth, PDAS_NoAlpha, , , clsBackground.Color, forceGrayscale, BMPIndexedColorCount, BMP16bpp555)
        End If
        
    End If
    
End Sub

Private Sub UpdatePreview()

    If (cmdBar.PreviewsAllowed And g_ImageFormats.FreeImageEnabled And sldColorCount.IsValid And (Not m_SrcImage Is Nothing)) Then
        
        'Make sure the preview source is up-to-date
        If (m_FIHandle = 0) Then UpdatePreviewSource
        
        'Retrieve a BMP-saved version of the current preview image
        workingDIB.ResetDIB
        If Plugin_FreeImage.GetExportPreview(m_FIHandle, workingDIB, PDIF_BMP) Then
            FinalizeNonstandardPreview pdFxPreview, True
        End If
        
    End If
    
End Sub

Private Sub sldColorCount_Change()
    If Not CBool(chkColorCount.Value) Then chkColorCount.Value = vbChecked
    UpdatePreviewSource
    UpdatePreview
End Sub
