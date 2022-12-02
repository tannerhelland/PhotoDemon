VERSION 5.00
Begin VB.Form dialog_ExportJXL 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   390
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   436
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   874
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.pdButtonStrip btsCategory 
      Height          =   615
      Left            =   6000
      TabIndex        =   1
      Top             =   120
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   1085
      FontSize        =   11
   End
   Begin PhotoDemon.pdCommandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5790
      Width           =   13110
      _ExtentX        =   23125
      _ExtentY        =   1323
      DontAutoUnloadParent=   -1  'True
   End
   Begin PhotoDemon.pdFxPreviewCtl pdFxPreview 
      Height          =   5625
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9922
   End
   Begin PhotoDemon.pdContainer picContainer 
      Height          =   4695
      Index           =   0
      Left            =   5880
      Top             =   1080
      Width           =   7215
      _ExtentX        =   0
      _ExtentY        =   0
      Begin PhotoDemon.pdSlider sldEffort 
         Height          =   975
         Left            =   120
         TabIndex        =   8
         Top             =   1320
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   1720
         Caption         =   "compression effort"
         Min             =   1
         Max             =   9
         Value           =   7
         NotchPosition   =   2
         NotchValueCustom=   7
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   375
         Index           =   0
         Left            =   120
         Top             =   120
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   450
         Caption         =   "quality"
         FontSize        =   12
      End
      Begin PhotoDemon.pdDropDown cboSaveQuality 
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   600
         Width           =   2610
         _ExtentX        =   4604
         _ExtentY        =   661
      End
      Begin PhotoDemon.pdSlider sltQuality 
         Height          =   405
         Left            =   2880
         TabIndex        =   3
         Top             =   600
         Width           =   4335
         _ExtentX        =   7223
         _ExtentY        =   873
         Min             =   1
         Max             =   100
         Value           =   90
         NotchPosition   =   1
         DefaultValue    =   90
      End
      Begin PhotoDemon.pdLabel lblHint 
         Height          =   255
         Index           =   0
         Left            =   480
         Top             =   2280
         Width           =   2340
         _ExtentX        =   4128
         _ExtentY        =   450
         Caption         =   "fast, larger file"
         FontItalic      =   -1  'True
         FontSize        =   9
      End
      Begin PhotoDemon.pdLabel lblHint 
         Height          =   255
         Index           =   1
         Left            =   2880
         Top             =   2280
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
      Height          =   4695
      Index           =   1
      Left            =   5880
      Top             =   1080
      Width           =   7215
      _ExtentX        =   0
      _ExtentY        =   0
      Begin PhotoDemon.pdButtonStrip btsSubsampling 
         Height          =   1095
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   1931
         Caption         =   "chroma subsampling"
      End
      Begin PhotoDemon.pdButtonStrip btsDepth 
         Height          =   1095
         Left            =   120
         TabIndex        =   7
         Top             =   1440
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   1931
         Caption         =   "depth"
      End
   End
   Begin PhotoDemon.pdContainer picContainer 
      Height          =   4695
      Index           =   2
      Left            =   5880
      Top             =   1080
      Width           =   7215
      _ExtentX        =   0
      _ExtentY        =   0
      Begin PhotoDemon.pdMetadataExport mtdManager 
         Height          =   4215
         Left            =   240
         TabIndex        =   6
         Top             =   120
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   7435
      End
   End
End
Attribute VB_Name = "dialog_ExportJXL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'JPEG XL Export Dialog
'Copyright 2022-2022 by Tanner Helland
'Created: 08/November/22
'Last updated: 08/November/22
'Last update: initial build
'
'Dialog for presenting the user various options related to JPEG XL exporting.  All export options rely on
' libjxl for their actual implementation.
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

'Current original preview DIB, cropped and zoomed as necessary (but otherwise unmodified).
Private m_PreviewDIB As pdDIB

'The quality checkboxes work as toggles.  To prevent infinite looping while they update each other, a module-level
' variable controls access to the toggle code.
Private m_DisableUIUpdates As Boolean

'OK or CANCEL result
Private m_UserDialogAnswer As VbMsgBoxResult

'Final format-specific XML packet, with all format-specific settings defined as tag+value pairs
Private m_FormatParamString As String

'Final metadata XML packet, with all metadata settings defined as tag+value pairs
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

Private Sub btsCategory_Click(ByVal buttonIndex As Long)
    UpdatePanelVisibility
End Sub

Private Sub UpdatePanelVisibility()
    Dim i As Long
    For i = 0 To btsCategory.ListCount - 1
        picContainer(i).Visible = (i = btsCategory.ListIndex)
    Next i
End Sub

Private Sub btsDepth_Click(ByVal buttonIndex As Long)
    UpdatePreviewSource
    UpdatePreview
End Sub

Private Sub btsSubsampling_Click(ByVal buttonIndex As Long)
    UpdatePreview
End Sub

Private Sub cboSaveQuality_Click()
    
    If (Not m_DisableUIUpdates) Then
        Select Case cboSaveQuality.ListIndex
            Case 0
                sltQuality.Value = 100
            Case 1
                sltQuality.Value = 92
            Case 2
                sltQuality = 80
            Case 3
                sltQuality = 65
            Case 4
                sltQuality = 40
        End Select
    End If
    
End Sub

Private Sub cmdBar_CancelClick()
    m_UserDialogAnswer = vbCancel
    Me.Hide
End Sub

Private Sub cmdBar_OKClick()
    
    'Store all parameters inside an XML string
    m_FormatParamString = GetParamString_JXL()
    
    'The metadata panel manages its own XML string
    m_MetadataParamString = mtdManager.GetMetadataSettings
    
    'Free resources that are no longer required
    Set m_CompositedImage = Nothing
    Set m_SrcImage = Nothing
    Set m_PreviewDIB = Nothing
    
    'Hide but *DO NOT UNLOAD* the form.  The dialog manager needs to retrieve the setting strings before unloading us
    m_UserDialogAnswer = vbOK
    Me.Visible = False
    
End Sub

Private Function GetParamString_JXL() As String
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.AddParam "jxl-quality", sltQuality.Value
    cParams.AddParam "jxl-effort", sldEffort.Value
    cParams.AddParam "jxl-subsampling", btsSubsampling.ListIndex
    
    Select Case btsDepth.ListIndex
        Case 0
            cParams.AddParam "jxl-color-format", "auto"
        Case 1
            cParams.AddParam "jxl-color-format", "color"
        Case 2
            cParams.AddParam "jxl-color-format", "gray"
    End Select
    
    GetParamString_JXL = cParams.GetParamString
    
End Function

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub cmdBar_ResetClick()
    
    'Default save quality is "Excellent"
    cboSaveQuality.ListIndex = 1
    sldEffort.Value = 7     'Default per libjxl
    
    'Default to 4:2:2 subsampling.  (Photoshop sets this automatically, depending on the selected quality, but it's
    ' too fiddly and prone to large jumps between otherwise small quality measurements.)
    btsSubsampling.ListIndex = 1
    
    'Auto color detection
    btsDepth.ListIndex = 0
    
    mtdManager.Reset
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'Used to keep the "image quality" text box, scroll bar, and combo box in sync
Private Sub UpdateDropDown()
    
    Select Case sltQuality.Value
        Case 40
            If (cboSaveQuality.ListIndex <> 4) Then cboSaveQuality.ListIndex = 4
        Case 65
            If (cboSaveQuality.ListIndex <> 3) Then cboSaveQuality.ListIndex = 3
        Case 80
            If (cboSaveQuality.ListIndex <> 2) Then cboSaveQuality.ListIndex = 2
        Case 92
            If (cboSaveQuality.ListIndex <> 1) Then cboSaveQuality.ListIndex = 1
        Case 100
            If (cboSaveQuality.ListIndex <> 0) Then cboSaveQuality.ListIndex = 0
        Case Else
            If (cboSaveQuality.ListIndex <> 5) Then cboSaveQuality.ListIndex = 5
    End Select
    
    If (Not m_DisableUIUpdates) Then UpdatePreview
    
End Sub

'The ShowDialog routine presents the user with this form.
Public Sub ShowDialog(Optional ByRef srcImage As pdImage = Nothing)
    
    'Provide a default answer of "cancel" (in the event that the user clicks the "x" button in the top-right)
    m_UserDialogAnswer = vbCancel
    
    'Make sure that the proper cursor is set
    Screen.MousePointer = 0
    Message "Waiting for user to specify export options... "
    
    'Populate the category button strip
    btsCategory.AddItem "basic", 0
    btsCategory.AddItem "advanced", 1
    btsCategory.AddItem "metadata", 2
    
    'Populate the "basic" options panel
    cboSaveQuality.SetAutomaticRedraws False
    cboSaveQuality.Clear
    cboSaveQuality.AddItem "lossless (100)", 0
    cboSaveQuality.AddItem "excellent (92)", 1
    cboSaveQuality.AddItem "good (80)", 2
    cboSaveQuality.AddItem "average (65)", 3
    cboSaveQuality.AddItem "low (40)", 4
    cboSaveQuality.AddItem "custom quality", 5
    cboSaveQuality.ListIndex = 1
    cboSaveQuality.SetAutomaticRedraws True, True
    
    'Populate the "advanced" options panel
    btsSubsampling.AddItem "none", 0
    btsSubsampling.AddItem "low (default)", 1
    btsSubsampling.AddItem "medium", 2
    btsSubsampling.AddItem "high", 3
    btsSubsampling.ListIndex = 1
    
    btsDepth.AddItem "auto", 0
    btsDepth.AddItem "color (24-bpp)", 1
    btsDepth.AddItem "black and white (8-bpp)", 2
    btsDepth.ListIndex = 0
    
    'Next, prepare various controls on the metadata panel
    Set m_SrcImage = srcImage
    mtdManager.SetParentImage m_SrcImage, PDIF_JXL
    
    'By default, the basic options panel is always shown.
    btsCategory.ListIndex = 0
    UpdatePanelVisibility
    
    'Make a copy of the composited image; it takes time to composite layers, so we don't want to redo this except
    ' when absolutely necessary.
    If (Not m_SrcImage Is Nothing) Then
        
        m_SrcImage.GetCompositedImage m_CompositedImage, True
        pdFxPreview.NotifyNonStandardSource m_CompositedImage.GetDIBWidth, m_CompositedImage.GetDIBHeight
        
    End If
    
    'In batch process mode, we won't have a sample image to preview
    If (m_SrcImage Is Nothing) Then Interface.ShowDisabledPreviewImage pdFxPreview
    
    'Update the preview
    UpdatePreviewSource
    UpdatePreview True
    
    'Apply translations and visual themes
    ApplyThemeAndTranslations Me
    Interface.SetFormCaptionW Me, g_Language.TranslateMessage("%1 options", "JPEG XL")
    If (Not g_WindowManager Is Nothing) Then g_WindowManager.SetFocusAPI cmdBar.hWnd
    
    'Display the dialog
    ShowPDDialog vbModal, Me, True

End Sub

Private Sub pdFxPreview_ViewportChanged()
    UpdatePreviewSource
    UpdatePreview
End Sub

Private Sub sldEffort_Change()
    UpdatePreview
End Sub

Private Sub sltQuality_Change()
    If (Not m_DisableUIUpdates) Then UpdateDropDown
End Sub

'When a parameter changes that requires a new base image for the preview (e.g. changing the background composite color),
' call this function to generate a new preview DIB.  Note that you *do not* need to call this function for format-specific
' changes (like quality, subsampling, etc).
Private Sub UpdatePreviewSource()

    If (Not m_CompositedImage Is Nothing) Then
        
        'Because the user can change the preview viewport, we can't guarantee that the preview region hasn't changed
        ' since the last preview.  Prep a new preview now.
        Dim tmpSafeArray As SafeArray2D
        EffectPrep.PreviewNonStandardImage tmpSafeArray, m_CompositedImage, pdFxPreview, False
        
        'The public workingDIB object now contains the preview area image.  Clone it locally.
        If (m_PreviewDIB Is Nothing) Then Set m_PreviewDIB = New pdDIB
        m_PreviewDIB.CreateFromExistingDIB workingDIB
        
        'Perform a one-time swizzle here (from BGRA to RGBA order)
        DIBs.SwizzleBR m_PreviewDIB
        
        'TODO: forcible color-depth changes here?
        
    End If
    
End Sub

Private Sub UpdatePreview(Optional ByVal forceUpdate As Boolean = False)

    If (cmdBar.PreviewsAllowed Or forceUpdate) And (Not m_SrcImage Is Nothing) And (Not m_PreviewDIB Is Nothing) Then
        
        'Retrieve a JPEG-XL version of the current preview image.
        If Plugin_jxl.PreviewJXL(m_PreviewDIB, workingDIB, GetParamString_JXL()) Then
            FinalizeNonstandardPreview pdFxPreview, True
        Else
            PDDebug.LogAction "WARNING: JPEG-XL EXPORT PREVIEW PROBLEM!"
        End If
        
    End If
    
End Sub
