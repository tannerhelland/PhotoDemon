VERSION 5.00
Begin VB.Form dialog_ExportPalette 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Palette export options"
   ClientHeight    =   6585
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10950
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
   ScaleHeight     =   439
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   730
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.pdButtonStrip btsTargetFile 
      Height          =   1050
      Left            =   4560
      TabIndex        =   6
      Top             =   3960
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   1852
      Caption         =   "target palette file"
   End
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   255
      Index           =   0
      Left            =   4560
      Top             =   3120
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   450
      Caption         =   "palette name"
      FontSize        =   12
   End
   Begin PhotoDemon.pdTextBox txtPaletteName 
      Height          =   375
      Left            =   4680
      TabIndex        =   5
      Top             =   3480
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   661
   End
   Begin PhotoDemon.pdSlider sldColorCount 
      Height          =   735
      Left            =   4560
      TabIndex        =   4
      Top             =   2280
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   1296
      Caption         =   "color count"
      Min             =   2
      Max             =   256
      Value           =   256
      NotchPosition   =   2
      NotchValueCustom=   256
   End
   Begin PhotoDemon.pdButtonStrip btsColorCount 
      Height          =   1050
      Left            =   4560
      TabIndex        =   3
      Top             =   1200
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   1852
      Caption         =   "palette color count"
   End
   Begin PhotoDemon.pdPaletteUI palPreview 
      Height          =   5655
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   9975
      Caption         =   "palette contents"
      UseRGBA         =   -1  'True
   End
   Begin PhotoDemon.pdButtonStrip btsPalette 
      Height          =   1050
      Left            =   4560
      TabIndex        =   1
      Top             =   120
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   1852
      Caption         =   "palette to export"
   End
   Begin PhotoDemon.pdCommandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5835
      Width           =   10950
      _ExtentX        =   19315
      _ExtentY        =   1323
      DontAutoUnloadParent=   -1  'True
   End
End
Attribute VB_Name = "dialog_ExportPalette"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Palette Export Dialog
'Copyright 2018-2026 by Tanner Helland
'Created: 25/March/18
'Last updated: 25/March/18
'Last update: initial build
'
'Dialog for presenting the user a number of options related to palette exporting.  In the future,
' I'd love to find time to write a full palette editor, but for now, this stripped-down dialog
' will have to do.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'This form can (and should!) be notified of the image being exported.  The only exception to this rule is invoking
' the dialog from the batch process dialog, as no image is associated with that preview.
Private m_SrcImage As pdImage

'Composited copy of the current image
Private m_CompositedImage As pdDIB

'If the current image is too big, we'll use a mini-copy instead
Private m_CurrentImageHuge As Boolean, m_MiniImage As pdDIB

'Desired format; some formats support optional bonus features
Private m_DstFormat As PD_PaletteFormat

'Target filename; some formats support "append" behavior, so we need to know if the target file exists
' and has a matching format.
Private m_DstFilename As String

'Previews use a persistent pdPalette object; this may be generated on-the-fly for the current image.
Private m_PreviewPalette As pdPalette

'OK or CANCEL result
Private m_UserDialogAnswer As VbMsgBoxResult

'Final format-specific XML packet, with all format-specific settings defined as tag+value pairs
Private m_FormatParamString As String

'The user's answer is returned via this property
Public Function GetDialogResult() As VbMsgBoxResult
    GetDialogResult = m_UserDialogAnswer
End Function

Public Function GetFormatParams() As String
    GetFormatParams = m_FormatParamString
End Function

Private Sub btsColorCount_Click(ByVal buttonIndex As Long)
    ReflowInterface
    UpdatePreview
End Sub

Private Sub btsPalette_Click(ByVal buttonIndex As Long)
    ReflowInterface
    UpdatePreview
End Sub

Private Sub cmdBar_CancelClick()
    m_UserDialogAnswer = vbCancel
    Me.Visible = False
End Sub

Private Sub cmdBar_OKClick()

    'Exit if anything fails validation!
    'If (Not sltQuality.IsValid) Then Exit Sub
    
    m_FormatParamString = GetCurrentParamString()
    
    'Free resources that are no longer required
    Set m_SrcImage = Nothing
    
    'Hide but *DO NOT UNLOAD* the form.  The dialog manager needs to retrieve the setting strings before unloading us
    m_UserDialogAnswer = vbOK
    Me.Visible = False
    
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub Form_Load()
    
    cmdBar.SetPreviewStatus False
    
    btsPalette.AddItem "current palette", 0
    btsPalette.AddItem "original embedded palette", 1
    btsPalette.ListIndex = 0
    
    btsColorCount.AddItem "auto", 0
    btsColorCount.AddItem "custom", 1
    btsColorCount.ListIndex = 0
    
    btsTargetFile.AddItem "overwrite", 0
    btsTargetFile.AddItem "append", 1
    btsTargetFile.AssignTooltip "Adobe Swatch Exchange (ASE) files can store multiple palettes inside a single file.  If you select the ""append"" option, PhotoDemon will add this palette to your existing ASE file.  Any palettes already inside the file will not be modified."
    btsTargetFile.ListIndex = 0
    
    cmdBar.SetPreviewStatus True
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'The ShowDialog routine presents the user with this form.
Public Sub ShowDialog(Optional ByRef srcImage As pdImage = Nothing, Optional ByVal palFormat As PD_PaletteFormat = pdpf_AdobeSwatchExchange, Optional ByRef dstFilename As String = vbNullString)

    cmdBar.SetPreviewStatus False

    'Provide a default answer of "cancel" (in the event that the user clicks the "x" button in the top-right)
    m_UserDialogAnswer = vbCancel
    Set m_SrcImage = srcImage
    m_DstFormat = palFormat
    m_DstFilename = dstFilename
    
    'Set the preview window's alpha handling status based on the output format; many palette formats
    ' do not support opacity, so we want to preview opacity conditionally.
    palPreview.UseRGBA = (palFormat = pdpf_PhotoDemon) Or (palFormat = pdpf_PaintDotNet)
    
    'Cache a copy of the fully composited image (if any)
    If (Not srcImage Is Nothing) Then
        
        Set m_CompositedImage = New pdDIB
        m_SrcImage.GetCompositedImage m_CompositedImage, False
        
        'If the current image is too large, create a mini-version that we can use for faster palette generation
        Dim numPixels As Long
        numPixels = srcImage.Width * srcImage.Height
        
        Const HUGE_IMAGE_THRESHOLD As Long = 100000
        m_CurrentImageHuge = (numPixels > HUGE_IMAGE_THRESHOLD)
        If m_CurrentImageHuge Then
            Set m_MiniImage = New pdDIB
            DIBs.ResizeDIBByPixelCount m_CompositedImage, m_MiniImage, HUGE_IMAGE_THRESHOLD, GP_IM_NearestNeighbor
        End If
        
        'Suggest an automatic palette name.  (Some palette file formats don't support names; the edit box
        ' will be automatically hidden by the layout function, as necessary.)
        If (LenB(m_DstFilename) <> 0) Then txtPaletteName.Text = Files.FileGetName(dstFilename, True) Else txtPaletteName.Text = g_Language.TranslateMessage("New palette")
        
    End If
    
    'Make sure that the proper cursor is set
    Screen.MousePointer = 0
    Message "Waiting for user to specify export options... "
    
    'Reflow the available interface options and update the preview
    ReflowInterface
    cmdBar.SetPreviewStatus True
    UpdatePreview True
    
    'Apply translations and visual themes
    ApplyThemeAndTranslations Me
    
    'Display the dialog
    ShowPDDialog vbModal, Me, True
    
End Sub

Private Sub ReflowInterface()

    Dim yOffset As Long
    yOffset = Interface.FixDPI(4)
    
    Dim ctlPadding As Long
    ctlPadding = Interface.FixDPI(8)
    
    If m_SrcImage.HasOriginalPalette Then
        btsPalette.SetTop yOffset
        btsPalette.Visible = True
        yOffset = yOffset + btsPalette.GetHeight + ctlPadding
    Else
        btsPalette.Visible = False
    End If
    
    'Different palette formats have maximum color count limits
    If (m_DstFormat = pdpf_PaintDotNet) Then
        sldColorCount.Max = 96
    ElseIf (m_DstFormat = pdpf_AdobeSwatchExchange) Then
        sldColorCount.Max = 4096
    Else
        sldColorCount.Max = 256
    End If
    
    btsColorCount.SetTop yOffset
    btsColorCount.Visible = True
    yOffset = yOffset + btsColorCount.GetHeight + ctlPadding
    
    If (btsColorCount.ListIndex <> 0) Then
        sldColorCount.SetTop yOffset
        sldColorCount.Visible = True
        yOffset = yOffset + sldColorCount.GetHeight + ctlPadding
    Else
        sldColorCount.Visible = False
    End If
    
    'Some palette formats support a separate embedded "name"
    If (m_DstFormat = pdpf_AdobeColorSwatch) Or (m_DstFormat = pdpf_AdobeSwatchExchange) Or (m_DstFormat = pdpf_GIMP) Then
        lblTitle(0).SetTop yOffset
        lblTitle(0).Visible = True
        yOffset = yOffset + lblTitle(0).GetHeight + ctlPadding
        txtPaletteName.SetTop yOffset
        txtPaletteName.Visible = True
        yOffset = yOffset + txtPaletteName.GetHeight + ctlPadding
    Else
        lblTitle(0).Visible = False
        txtPaletteName.Visible = False
    End If
    
    'ASE palettes are unique in supporting multiple palettes within a single file.  To make it easier
    ' to save a bunch of palettes to a single file, let's offer an overwrite vs merge feature when
    ' saving to an existing ASE file.
    If (m_DstFormat = pdpf_AdobeSwatchExchange) And (Files.FileExists(m_DstFilename)) Then
        btsTargetFile.SetTop yOffset
        btsTargetFile.Visible = True
        yOffset = yOffset + btsTargetFile.GetHeight + ctlPadding
    Else
        btsTargetFile.Visible = False
    End If
    
End Sub

Private Sub UpdatePreview(Optional ByVal forceUpdate As Boolean = False)

    If (cmdBar.PreviewsAllowed Or forceUpdate) And (Not m_SrcImage Is Nothing) Then
        
        'Before proceeding, figure out what our source image is
        Dim tmpDIB As pdDIB
        If m_CurrentImageHuge Then Set tmpDIB = m_MiniImage Else Set tmpDIB = m_CompositedImage
        
        Dim tmpQuad() As RGBQuad, numColors As Long
        If (btsColorCount.ListIndex = 0) Then numColors = 256 Else numColors = sldColorCount.Value
        
        'The user can choose to export the image's current palette (e.g. including any changes),
        ' or if the image contained an embedded palette at load-time, we can export that, instead.
        Select Case btsPalette.ListIndex
        
            'Current
            Case 0
                If Palettes.GetOptimizedPaletteIncAlpha(tmpDIB, tmpQuad, numColors, pdqs_Variance) Then
                    Set m_PreviewPalette = New pdPalette
                    m_PreviewPalette.CreateFromPaletteArray tmpQuad, UBound(tmpQuad) + 1
                End If
            
            'Original
            Case 1
                If m_SrcImage.HasOriginalPalette Then m_SrcImage.GetOriginalPalette m_PreviewPalette
        
        End Select
        
        'Render the new palette
        palPreview.SetPDPalette m_PreviewPalette
        
    End If
    
End Sub

Private Sub sldColorCount_Change()
    UpdatePreview
End Sub

Private Function GetCurrentParamString() As String

    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    
    With cParams
        
        .AddParam "srcPalette", btsPalette.ListIndex
        .AddParam "colorCount", btsColorCount.ListIndex
        
        'In auto mode, the number of target colors varies based on the target palette format
        Dim numColors As Long
        If (btsColorCount.ListIndex = 0) Then
            numColors = -1
        Else
            numColors = sldColorCount.Value
        End If
        
        .AddParam "numColors", numColors
        
        'Some settings, like embedded palette name, are only supported by certain formats, but we
        ' always pass the name - the encoder will deal with this.
        .AddParam "palName", txtPaletteName.Text
        .AddParam "embedPaletteASE", btsTargetFile.ListIndex
        
    End With
    
    GetCurrentParamString = cParams.GetParamString()

End Function
