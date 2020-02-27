VERSION 5.00
Begin VB.Form dialog_ExportPSD 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " PSD Export Options"
   ClientHeight    =   6585
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   12630
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
   ScaleHeight     =   439
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   842
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.pdCommandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5835
      Width           =   12630
      _ExtentX        =   22278
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
   Begin PhotoDemon.pdButtonStrip btsCategory 
      Height          =   615
      Left            =   5880
      TabIndex        =   2
      Top             =   120
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   1085
      FontSize        =   11
   End
   Begin PhotoDemon.pdContainer picContainer 
      Height          =   4815
      Index           =   0
      Left            =   5880
      TabIndex        =   5
      Top             =   840
      Width           =   6615
      _ExtentX        =   0
      _ExtentY        =   0
      Begin PhotoDemon.pdButtonStrip btsCompression 
         Height          =   975
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   1720
         Caption         =   "compression"
      End
      Begin PhotoDemon.pdButtonStrip btsCompatibility 
         Height          =   975
         Left            =   240
         TabIndex        =   7
         Top             =   1320
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   1720
         Caption         =   "maximize compatibility"
      End
   End
   Begin PhotoDemon.pdContainer picContainer 
      Height          =   4815
      Index           =   1
      Left            =   5880
      TabIndex        =   3
      Top             =   840
      Width           =   6615
      _ExtentX        =   0
      _ExtentY        =   0
      Begin PhotoDemon.pdMetadataExport mtdManager 
         Height          =   4215
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   7435
      End
   End
End
Attribute VB_Name = "dialog_ExportPSD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Adobe Photoshop (PSD) Export Dialog
'Copyright 2019-2020 by Tanner Helland
'Created: 18/February/19
'Last updated: 19/February/19
'Last update: wrap up initial build
'
'This dialog works as a simple relay to the pdPSD class (and its associated child classes).  Look there for specific
' encoding details.
'
'Given the breadth of features supported by Photoshop, potential PSD export settings are many and varied.  I have
' tried to pare down the UI toggles to only the most essential elements.  If you find that exported PSD files are
' not what you expect, please notify me so I can improve PhotoDemon's PSD engine!
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit https://photodemon.org/license/
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

Private Sub cmdBar_CancelClick()
    m_UserDialogAnswer = vbCancel
    Me.Visible = False
End Sub

Private Sub cmdBar_OKClick()

    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.AddParam "compression", btsCompression.ListIndex
    cParams.AddParam "max-compatibility", CBool(btsCompatibility.ListIndex = 1)
    
    m_FormatParamString = cParams.GetParamString
    
    'The metadata panel manages its own XML string
    m_MetadataParamString = mtdManager.GetMetadataSettings
    
    'Free resources that are no longer required
    Set m_CompositedImage = Nothing
    Set m_SrcImage = Nothing
    
    'Hide but *DO NOT UNLOAD* the form.  The dialog manager needs to retrieve the setting strings before unloading us
    m_UserDialogAnswer = vbOK
    Me.Visible = False

End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub cmdBar_ResetClick()
    mtdManager.Reset
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
    Plugin_FreeImage.ReleasePreviewCache m_FIHandle
End Sub

Private Sub pdFxPreview_ViewportChanged()
    UpdatePreviewSource
    UpdatePreview
End Sub

'The ShowDialog routine presents the user with this form.
Public Sub ShowDialog(Optional ByRef srcImage As pdImage = Nothing)
    
    'Provide a default answer of "cancel" (in the event that the user clicks the "x" button in the top-right)
    m_UserDialogAnswer = vbCancel
    
    'Make sure that the proper cursor is set
    Screen.MousePointer = 0
    Message "Waiting for user to specify export options... "
    
    'Next, prepare various controls on the metadata panel
    Set m_SrcImage = srcImage
    mtdManager.SetParentImage m_SrcImage, PDIF_PSD
    
    'By default, the basic options panel is always shown.
    btsCategory.AddItem "basic", 0
    btsCategory.AddItem "advanced", 1
    btsCategory.ListIndex = 0
    UpdatePanelVisibility
    
    'Populate any other list elements
    btsCompression.AddItem "none", 0
    btsCompression.AddItem "PackBits", 1
    btsCompression.ListIndex = 1
    
    btsCompatibility.AddItem "no", 0
    btsCompatibility.AddItem "yes", 1
    btsCompatibility.ListIndex = 1
    
    'Make a copy of the composited image; it takes time to composite layers, so we don't want to redo
    ' this except when absolutely necessary.
    If ((m_SrcImage Is Nothing) Or (Not ImageFormats.IsFreeImageEnabled())) Then
        Interface.ShowDisabledPreviewImage pdFxPreview
    Else
        m_SrcImage.GetCompositedImage m_CompositedImage, True
        pdFxPreview.NotifyNonStandardSource m_CompositedImage.GetDIBWidth, m_CompositedImage.GetDIBHeight
    End If
    
    'Update the preview
    UpdatePreviewSource
    UpdatePreview True
    
    'Apply translations and visual themes
    ApplyThemeAndTranslations Me
    
    'Display the dialog
    ShowPDDialog vbModal, Me, True
    
End Sub

'When a parameter changes that requires a new source DIB for the preview (e.g. changing the background composite color),
' call this function to generate a new preview DIB.  Note that you *do not* need to call this function for format-specific
' changes (like quality, subsampling, etc).
Private Sub UpdatePreviewSource()

    If (Not m_CompositedImage Is Nothing) Then
        
        'Because the user can change the preview viewport, we can't guarantee that the preview region hasn't changed
        ' since the last preview.  Prep a new preview now.
        Dim tmpSafeArray As SafeArray2D
        EffectPrep.PreviewNonStandardImage tmpSafeArray, m_CompositedImage, pdFxPreview, False
        
        'Finally, convert that preview copy to a FreeImage-compatible handle.
        If (m_FIHandle <> 0) Then Plugin_FreeImage.ReleaseFreeImageObject m_FIHandle
        
        'During previews, we can always use 32-bpp mode
        m_FIHandle = Plugin_FreeImage.GetFIDib_SpecificColorMode(workingDIB, 32, PDAS_ComplicatedAlpha)
        
    End If
    
End Sub

Private Sub UpdatePreview(Optional ByVal forceUpdate As Boolean = False)

    If ((cmdBar.PreviewsAllowed Or forceUpdate) And ImageFormats.IsFreeImageEnabled() And (Not m_SrcImage Is Nothing)) Then
        
        'Make sure the preview source is up-to-date
        If (m_FIHandle = 0) Then UpdatePreviewSource
        
        'Prep all relevant FreeImage flags
        Dim fi_Flags As FREE_IMAGE_SAVE_OPTIONS
        
        'Retrieve a JPEG-saved version of the current preview image
        workingDIB.ResetDIB
        If Plugin_FreeImage.GetExportPreview(m_FIHandle, workingDIB, PDIF_BMP, fi_Flags) Then
            workingDIB.SetAlphaPremultiplication True, True
            FinalizeNonstandardPreview pdFxPreview, True
        Else
            Debug.Print "WARNING: JP2 EXPORT PREVIEW IS HORRIBLY BROKEN!"
        End If
        
    End If

End Sub

Private Sub UpdatePanelVisibility()
    Dim i As Long
    For i = 0 To btsCategory.ListCount - 1
        picContainer(i).Visible = (i = btsCategory.ListIndex)
    Next i
End Sub
