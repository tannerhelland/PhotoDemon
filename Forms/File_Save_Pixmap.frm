VERSION 5.00
Begin VB.Form dialog_ExportPixmap 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   6540
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12495
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
   Icon            =   "File_Save_Pixmap.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   436
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   833
   Begin PhotoDemon.pdCheckBox chkFileExtension 
      Height          =   375
      Left            =   6000
      TabIndex        =   6
      Top             =   1440
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   661
      Caption         =   "change file extension to match color model"
   End
   Begin PhotoDemon.pdCommandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5790
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   1323
      DontAutoUnloadParent=   -1  'True
   End
   Begin PhotoDemon.pdFxPreviewCtl pdFxPreview 
      Height          =   5625
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9922
   End
   Begin PhotoDemon.pdButtonStrip btsFormat 
      Height          =   1095
      Left            =   5880
      TabIndex        =   2
      Top             =   3120
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   1931
      Caption         =   "format"
   End
   Begin PhotoDemon.pdColorSelector clsBackground 
      Height          =   1095
      Left            =   5880
      TabIndex        =   3
      Top             =   4320
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   1931
      Caption         =   "background color"
   End
   Begin PhotoDemon.pdButtonStrip btsDepth 
      Height          =   1095
      Left            =   5880
      TabIndex        =   4
      Top             =   1920
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   1931
      Caption         =   "depth"
   End
   Begin PhotoDemon.pdButtonStrip btsColorModel 
      Height          =   1095
      Left            =   5880
      TabIndex        =   5
      Top             =   240
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   1931
      Caption         =   "color model"
   End
End
Attribute VB_Name = "dialog_ExportPixmap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Portable Pixmap Export Dialog
'Copyright 2000-2026 by Tanner Helland
'Created: 01/May/16
'Last updated: 11/August/17
'Last update: improve flow of export dialog (by auto-hiding the background color selector when the source
'             image doesn't contain meaningful transparency data)
'
'Dialog for presenting the user various options related to PBM/PGM/PPM/PFM exporting.  All features
' rely on FreeImage for implementation, and this format will simply not be available for export or
' import if FreeImage cannot be found.
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
' cannot write any pixmap-specific data.
Private m_MetadataParamString As String

'If the source image contains transparency, this will be set to TRUE.  Various export options can be disabled
' or hidden if we don't have to deal with transparency in the source image.
Private m_ImageHasTransparency As Boolean

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

Private Sub btsColorModel_Click(ByVal buttonIndex As Long)
    UpdateComponentVisibility
    UpdatePreviewSource
    UpdatePreview
End Sub

Private Sub UpdateComponentVisibility()
    
    'There is no "color depth" option for monochrome images
    If (btsColorModel.ListIndex = 3) Then
        btsFormat.SetTop btsDepth.GetTop
        btsDepth.Visible = False
    Else
        btsDepth.Visible = True
        btsFormat.SetTop (btsDepth.GetTop + btsDepth.GetHeight) + FixDPI(8)
    End If
    
    'There is no "format" option for float images
    btsFormat.Visible = (btsDepth.ListIndex <> 3)
    
    'Show/hide the background color option if the current image has meaningful transparency data
    If m_ImageHasTransparency Then
        If (btsDepth.ListIndex <> 3) Then
            clsBackground.SetTop btsFormat.GetTop + btsFormat.GetHeight + Interface.FixDPI(8)
        Else
            clsBackground.SetTop btsFormat.GetTop
        End If
        clsBackground.Visible = True
    Else
        clsBackground.Visible = False
    End If
    
End Sub

Private Sub btsDepth_Click(ByVal buttonIndex As Long)
    UpdateComponentVisibility
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
    
    'Store all parameters inside an XML string
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    
    Dim pnmColorModel As String
    Select Case btsColorModel.ListIndex
        
        'Auto
        Case 0
            pnmColorModel = "auto"
        
        'RGB
        Case 1
            pnmColorModel = "color"
        
        'Gray
        Case 2
            pnmColorModel = "gray"
        
        'Monochrome
        Case 3
            pnmColorModel = "monochrome"
    
    End Select
    
    Dim pnmColorDepth As String
    If (btsColorModel.ListIndex = 3) Then
        pnmColorDepth = "standard"
    Else
        Select Case btsDepth.ListIndex
        
            '"auto" depth just corresponds to "standard" depth, at present
            Case 0, 1
                pnmColorDepth = "standard"
                
            Case 2
                pnmColorDepth = "HDR"
            
            Case 3
                pnmColorDepth = "float"
        
        End Select
    End If
    
    cParams.AddParam "pnm-color-model", pnmColorModel
    cParams.AddParam "pnm-color-depth", pnmColorDepth
    cParams.AddParam "pnm-change-extension", chkFileExtension.Value
    cParams.AddParam "pnm-use-ascii", (btsFormat.ListIndex = 1)
    cParams.AddParam "pnm-background-color", clsBackground.Color
    
    m_FormatParamString = cParams.GetParamString
    
    'Free resources that are no longer required
    Set m_CompositedImage = Nothing
    Set m_SrcImage = Nothing
    
    'Hide but *DO NOT UNLOAD* the form.  The dialog manager needs to retrieve the setting strings before unloading us
    m_UserDialogAnswer = vbOK
    Me.Visible = False
    
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreviewSource
End Sub

Private Sub cmdBar_ResetClick()
    clsBackground.Color = vbWhite
    btsColorModel.ListIndex = 0
    chkFileExtension.Value = True
    btsDepth.ListIndex = 0
    btsFormat.ListIndex = 0
End Sub

'The ShowDialog routine presents the user with this form.
Public Sub ShowDialog(Optional ByRef srcImage As pdImage = Nothing)
    
    'Provide a default answer of "cancel" (in the event that the user clicks the "x" button in the top-right)
    m_UserDialogAnswer = vbCancel
    
    'Make sure that the proper cursor is set
    Screen.MousePointer = 0
    Message "Waiting for user to specify export options... "
    
    'Populate the color model button strip
    btsColorModel.AddItem "auto", 0
    btsColorModel.AddItem "RGB", 1
    btsColorModel.AddItem "grayscale", 2
    btsColorModel.AddItem "monochrome", 3
    btsColorModel.ListIndex = 0
    
    'Populate available color depths
    btsDepth.AddItem "auto", 0
    btsDepth.AddItem "standard", 1
    btsDepth.AddItem "HDR", 2
    btsDepth.AddItem "floating-point", 3
    btsDepth.ListIndex = 0
    
    'Populate format options
    btsFormat.AddItem "binary", 0
    btsFormat.AddItem "ASCII", 1
    btsFormat.ListIndex = 0
    
    'Create a local reference to our parent image; we need this for generating live previews
    Set m_SrcImage = srcImage
    
    'Make a copy of the composited image; it takes time to composite layers, so we don't want to redo this except
    ' when absolutely necessary.
    If ((m_SrcImage Is Nothing) Or (Not ImageFormats.IsFreeImageEnabled())) Then
        Interface.ShowDisabledPreviewImage pdFxPreview
    Else
        
        m_SrcImage.GetCompositedImage m_CompositedImage, True
        pdFxPreview.NotifyNonStandardSource m_CompositedImage.GetDIBWidth, m_CompositedImage.GetDIBHeight
        
        'Detect the source image's transparency state
        m_ImageHasTransparency = DIBs.IsDIBTransparent(m_CompositedImage)
        
    End If
    
    'Update the preview
    UpdateComponentVisibility
    UpdatePreviewSource
    UpdatePreview True
    
    'Apply translations and visual themes
    ApplyThemeAndTranslations Me, True, True
    Interface.SetFormCaptionW Me, g_Language.TranslateMessage("%1 options", "Pixmap")
    
    'Display the dialog
    ShowPDDialog vbModal, Me, True

End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
    Plugin_FreeImage.ReleasePreviewCache m_FIHandle
End Sub

Private Sub pdFxPreview_ViewportChanged()
    UpdatePreviewSource
    UpdatePreview
End Sub

'When a parameter changes that requires a new source DIB for the preview (e.g. changing the background composite color),
' call this function to generate a new preview DIB.  Note that you *do not* need to call this function for format-specific
' changes (like quality, subsampling, etc).
Private Sub UpdatePreviewSource()

    If (Not m_CompositedImage Is Nothing) Then
        
        'Because the user can change the preview viewport, we can't guarantee that the preview region hasn't changed
        ' since the last preview.  Prep a new preview now.
        Dim tmpSafeArray As SafeArray2D
        EffectPrep.PreviewNonStandardImage tmpSafeArray, m_CompositedImage, pdFxPreview, True
        
        'PNM formats don't support transparency, so we can save some time by forcibly converting to 24-bpp in advance
        If (workingDIB.GetDIBColorDepth = 32) Then workingDIB.ConvertTo24bpp clsBackground.Color
        
        'Finally, convert that preview copy to a FreeImage-compatible handle.  Because PNM formats are so limited,
        ' this step is very simple.
        If (m_FIHandle <> 0) Then Plugin_FreeImage.ReleaseFreeImageObject m_FIHandle
        If (btsColorModel.ListIndex = 0) Or (btsColorModel.ListIndex = 1) Then
            m_FIHandle = Plugin_FreeImage.GetFIDib_SpecificColorMode(workingDIB, 24, PDAS_NoAlpha)
        ElseIf (btsColorModel.ListIndex = 2) Then
            m_FIHandle = Plugin_FreeImage.GetFIDib_SpecificColorMode(workingDIB, 8, PDAS_NoAlpha, , , , True)
        Else
            m_FIHandle = Plugin_FreeImage.GetFIDib_SpecificColorMode(workingDIB, 1, PDAS_NoAlpha)
        End If
        
    End If
    
End Sub

Private Sub UpdatePreview(Optional ByVal forceUpdate As Boolean = False)

    If ((cmdBar.PreviewsAllowed Or forceUpdate) And ImageFormats.IsFreeImageEnabled() And (Not m_SrcImage Is Nothing)) Then
        
        'Make sure the preview source is up-to-date
        If (m_FIHandle = 0) Then UpdatePreviewSource
        
        Dim outputFormat As PD_IMAGE_FORMAT
        If (btsColorModel.ListIndex = 0) Or (btsColorModel.ListIndex = 1) Then
            outputFormat = PDIF_PPMRAW
        ElseIf (btsColorModel.ListIndex = 2) Then
            outputFormat = PDIF_PGMRAW
        Else
            outputFormat = PDIF_PBMRAW
        End If
        
        workingDIB.ResetDIB
        If Plugin_FreeImage.GetExportPreview(m_FIHandle, workingDIB, outputFormat) Then
            FinalizeNonstandardPreview pdFxPreview, True
        Else
            Debug.Print "WARNING: PNM export previews failed for reasons unknown."
        End If
        
    End If
    
End Sub
