VERSION 5.00
Begin VB.Form dialog_ExportTIFF 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   6900
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13095
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
   Icon            =   "File_Save_TIFF.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   460
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   873
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
   End
   Begin PhotoDemon.pdTitle ttlStandard 
      Height          =   375
      Index           =   0
      Left            =   5880
      TabIndex        =   2
      Top             =   240
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
      Left            =   5880
      TabIndex        =   4
      Top             =   600
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
      Left            =   5880
      TabIndex        =   5
      Top             =   960
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   661
      Caption         =   "metadata settings"
      FontBold        =   -1  'True
      FontSize        =   12
      Value           =   0   'False
   End
   Begin PhotoDemon.pdContainer picContainer 
      Height          =   3855
      Index           =   0
      Left            =   5880
      Top             =   1440
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   6800
      Begin PhotoDemon.pdButtonStrip btsCompressionColor 
         Height          =   1095
         Left            =   0
         TabIndex        =   6
         Top             =   120
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   1931
         Caption         =   "color compression"
      End
      Begin PhotoDemon.pdButtonStrip btsCompressionMono 
         Height          =   1095
         Left            =   0
         TabIndex        =   7
         Top             =   1320
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   1931
         Caption         =   "monochrome compression"
      End
      Begin PhotoDemon.pdButtonStrip btsMultipage 
         Height          =   1095
         Left            =   0
         TabIndex        =   8
         Top             =   2520
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   1931
         Caption         =   "page format"
      End
   End
   Begin PhotoDemon.pdContainer picContainer 
      Height          =   3975
      Index           =   1
      Left            =   5880
      Top             =   1560
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   7011
      Begin PhotoDemon.pdColorDepth clrDepth 
         Height          =   2055
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   3625
      End
   End
   Begin PhotoDemon.pdContainer picContainer 
      Height          =   3975
      Index           =   2
      Left            =   5880
      Top             =   1440
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   7011
      Begin PhotoDemon.pdMetadataExport mtdManager 
         Height          =   3255
         Left            =   120
         TabIndex        =   3
         Top             =   0
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   5741
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
'Copyright 2012-2026 by Tanner Helland
'Created: 11/December/12
'Last updated: 29/April/16
'Last update: repurpose old color-depth dialog into a TIFF-specific one
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

Private Sub clrDepth_Change()
    UpdatePreviewSource
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
    
    cmdBar.SetPreviewStatus False
    
    'General panel settings
    btsCompressionColor.ListIndex = 0
    btsCompressionMono.ListIndex = 0
    btsMultipage.ListIndex = 0
    
    'Metadata settings
    mtdManager.Reset
    
    cmdBar.SetPreviewStatus True
    UpdatePreviewSource
    UpdatePreview
    
End Sub

'The ShowDialog routine presents the user with this form.
Public Sub ShowDialog(Optional ByRef srcImage As pdImage = Nothing)

    'Provide a default answer of "cancel" (in the event that the user clicks the "x" button in the top-right)
    m_UserDialogAnswer = vbCancel
    
    Message "Waiting for user to specify export options... "
    
    'Basic options
    btsMultipage.AddItem "single page (composited image)", 0
    btsMultipage.AddItem "multipage (one page per layer)", 1
    btsMultipage.ListIndex = 0
    
    If (Not srcImage Is Nothing) Then btsMultipage.Visible = (srcImage.GetNumOfLayers > 1) Else btsMultipage.Visible = True
    
    'Synchronize the size of the first panel to match whatever elements are still visible
    Dim srcImageIsMultipage As Boolean
    If (Not srcImage Is Nothing) Then srcImageIsMultipage = (srcImage.GetNumOfLayers > 1) Else srcImageIsMultipage = True
    
    If srcImageIsMultipage Then
        picContainer(0).SetHeight btsMultipage.GetTop + btsMultipage.GetHeight + Interface.FixDPI(8)
    Else
        picContainer(0).SetHeight btsCompressionMono.GetTop + btsCompressionMono.GetHeight + Interface.FixDPI(8)
    End If
    
    'TIFF exposes a number of specialized compression settings
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
    
    'Standard settings are accessed via pdTitle controls.  Because the panels are so large, only one panel
    ' is allowed open at a time.
    Dim i As Long
    For i = picContainer.lBound To picContainer.UBound
        picContainer(i).SetLeft pdFxPreview.GetLeft + pdFxPreview.GetWidth + Interface.FixDPI(8)
    Next i
    
    clrDepth.SyncToIdealSize
    ttlStandard(0).Value = True
    m_ActiveTitleBar = 0
    UpdateStandardTitlebars
    
    'Prep a preview (if any)
    Set m_SrcImage = srcImage
    If (Not m_SrcImage Is Nothing) Then
        m_SrcImage.GetCompositedImage m_CompositedImage, True
        pdFxPreview.NotifyNonStandardSource m_CompositedImage.GetDIBWidth, m_CompositedImage.GetDIBHeight
    End If
    If (Not ImageFormats.IsFreeImageEnabled()) Or (m_SrcImage Is Nothing) Then Interface.ShowDisabledPreviewImage pdFxPreview
    
    'Next, prepare various controls on the metadata panel
    mtdManager.SetParentImage m_SrcImage, PDIF_TIFF
    
    'Update the preview
    UpdatePreviewSource
    UpdatePreview
    
    'Apply translations and visual themes
    ApplyThemeAndTranslations Me, True, True
    Interface.SetFormCaptionW Me, g_Language.TranslateMessage("%1 options", "TIFF")
    
    'Many UI options are dynamically shown/hidden depending on other settings; make sure their initial state is correct
    UpdateStandardPanelVisibility
    
    'Display the dialog
    ShowPDDialog vbModal, Me, True

End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
    Plugin_FreeImage.ReleasePreviewCache m_FIHandle
End Sub

Private Function GetExportParamString() As String

    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    
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
    
    cParams.AddParam "tiff-compression-color", compressName
    
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
    
    cParams.AddParam "tiff-compression-mono", compressName
    
    If (btsMultipage.ListIndex <> 0) Then
        cParams.AddParam "tiff-multipage", True
    Else
        cParams.AddParam "tiff-multipage", False
    End If
        
    'Next come all the messy color-depth possibilities
    cParams.AddParam "tiff-color-depth", clrDepth.GetAllSettings
    
    GetExportParamString = cParams.GetParamString
    
End Function

Private Sub pdFxPreview_ColorSelected()
    clrDepth.NotifyNewAlphaColor pdFxPreview.SelectedColor
End Sub

Private Sub pdFxPreview_ViewportChanged()
    UpdatePreviewSource
    UpdatePreview
End Sub

'When a parameter changes that requires a new source DIB for the preview (e.g. changing the background composite color,
' changing the output color depth), you must call this function to generate a new preview DIB.  Note that you *do not*
' need to call this function for format-specific changes (e.g. compression settings).
Private Sub UpdatePreviewSource()
    
    If (Not m_CompositedImage Is Nothing) Then
        
        'Because the user can change the preview viewport, we can't guarantee that the preview region hasn't changed
        ' since the last preview.  Prep a new preview now.
        Dim tmpSafeArray As SafeArray2D
        EffectPrep.PreviewNonStandardImage tmpSafeArray, m_CompositedImage, pdFxPreview, True
        
        'To reduce the chance of bugs, we use the same parameter parsing technique as the core TIFF encoder
        Dim cParams As pdSerialize
        Set cParams = New pdSerialize
        cParams.SetParamString GetExportParamString()
        
        'The color-depth-specific options are embedded as a single option, so extract them into their
        ' own parser.
        Dim cParamsDepth As pdSerialize
        Set cParamsDepth = New pdSerialize
        cParamsDepth.SetParamString cParams.GetString("tiff-color-depth", vbNullString)
        
        'Color and grayscale modes require different processing, so start there
        Dim forceGrayscale As Boolean
        forceGrayscale = ParamsEqual(cParamsDepth.GetString("cd-color-model", "auto"), "gray")
        
        'For 8-bit modes, grab a palette size.  (This parameter will be ignored in other color modes.)
        Dim newPaletteSize As Long
        newPaletteSize = cParamsDepth.GetLong("cd-palette-size", 256)
        
        'Convert the text-only descriptors of color depth into a meaningful bpp value
        Dim newColorDepth As Long
        
        If ParamsEqual(cParamsDepth.GetString("cd-color-model", "auto"), "auto") Then
            newColorDepth = 32
        Else
            
            'HDR modes do not need to be previewed, so we forcibly downsample them here
            If forceGrayscale Then
                
                newColorDepth = 8
                
                If ParamsEqual(cParamsDepth.GetString("cd-gray-depth", "auto"), "gray-monochrome") Then
                    newPaletteSize = 2
                End If
                
            Else
                
                If ParamsEqual(cParamsDepth.GetString("cd-color-depth", "color-standard"), "color-indexed") Then
                    newColorDepth = 8
                Else
                    newColorDepth = 32
                End If
                
            End If
        
        End If
        
        'Next comes transparency, which is somewhat messy because PNG alpha behavior deviates significantly from normal alpha behavior.
        Dim desiredAlphaMode As PD_ALPHA_STATUS, desiredAlphaCutoff As Long
        
        If ParamsEqual(cParamsDepth.GetString("cd-alpha-model", "auto"), "auto") Or ParamsEqual(cParamsDepth.GetString("cd-alpha-model", "auto"), "full") Then
            desiredAlphaMode = PDAS_ComplicatedAlpha
            If (newColorDepth = 24) Then newColorDepth = 32
        ElseIf ParamsEqual(cParamsDepth.GetString("cd-alpha-model", "auto"), "none") Then
            desiredAlphaMode = PDAS_NoAlpha
            If (newColorDepth = 32) Then newColorDepth = 24
            desiredAlphaCutoff = 0
        ElseIf ParamsEqual(cParamsDepth.GetString("cd-alpha-model", "auto"), "by-cutoff") Then
            desiredAlphaMode = PDAS_BinaryAlpha
            desiredAlphaCutoff = cParamsDepth.GetLong("cd-alpha-cutoff", PD_DEFAULT_ALPHA_CUTOFF)
            If (newColorDepth = 24) Then newColorDepth = 32
        ElseIf ParamsEqual(cParamsDepth.GetString("cd-alpha-model", "auto"), "by-color") Then
            desiredAlphaMode = PDAS_NewAlphaFromColor
            desiredAlphaCutoff = cParamsDepth.GetLong("cd-alpha-color", vbBlack)
            If (newColorDepth = 24) Then newColorDepth = 32
        End If
        
        If (m_FIHandle <> 0) Then Plugin_FreeImage.ReleaseFreeImageObject m_FIHandle
        m_FIHandle = Plugin_FreeImage.GetFIDib_SpecificColorMode(workingDIB, newColorDepth, desiredAlphaMode, PDAS_ComplicatedAlpha, desiredAlphaCutoff, cParamsDepth.GetLong("cd-matte-color", vbWhite), forceGrayscale, newPaletteSize, , True)
        
    End If
    
End Sub

Private Function ParamsEqual(ByVal param1 As String, ByVal param2 As String) As Boolean
    ParamsEqual = Strings.StringsEqual(param1, param2, True)
End Function

Private Sub UpdatePreview()

    If (cmdBar.PreviewsAllowed And ImageFormats.IsFreeImageEnabled() And clrDepth.IsValid And (Not m_SrcImage Is Nothing)) Then
        
        'Make sure the preview source is up-to-date
        If (m_FIHandle = 0) Then UpdatePreviewSource
        
        'Retrieve a PNG-saved version of the current preview image
        workingDIB.ResetDIB
        If Plugin_FreeImage.GetExportPreview(m_FIHandle, workingDIB, PDIF_TIFF, FISO_TIFF_NONE) Then FinalizeNonstandardPreview pdFxPreview, True
        
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
    yPos = FixDPI(8)
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

