VERSION 5.00
Begin VB.Form dialog_ExportJPEG 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " JPEG Export Options"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   13110
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
   ScaleHeight     =   436
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   874
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.pdButtonStrip btsCategory 
      Height          =   615
      Left            =   5880
      TabIndex        =   1
      Top             =   120
      Width           =   7095
      _ExtentX        =   12515
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
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      HasDC           =   0   'False
      Height          =   4695
      Index           =   1
      Left            =   5880
      ScaleHeight     =   313
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   481
      TabIndex        =   10
      Top             =   1080
      Width           =   7215
      Begin PhotoDemon.pdMetadataExport mtdManager 
         Height          =   4215
         Left            =   240
         TabIndex        =   13
         Top             =   120
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   7435
      End
   End
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4695
      Index           =   0
      Left            =   5880
      ScaleHeight     =   313
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   481
      TabIndex        =   3
      Top             =   1080
      Width           =   7215
      Begin PhotoDemon.pdDropDown cboSaveQuality 
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   2730
         _ExtentX        =   4815
         _ExtentY        =   661
      End
      Begin PhotoDemon.pdDropDown cboSubsample 
         Height          =   375
         Left            =   720
         TabIndex        =   11
         Top             =   2910
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   661
      End
      Begin VB.ComboBox cmbAutoQuality 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   360
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   4005
         Visible         =   0   'False
         Width           =   6735
      End
      Begin PhotoDemon.pdCheckBox chkOptimize 
         Height          =   330
         Left            =   360
         TabIndex        =   5
         Top             =   1350
         Width           =   6690
         _ExtentX        =   11800
         _ExtentY        =   582
         Caption         =   "optimize compression tables"
      End
      Begin PhotoDemon.pdCheckBox chkProgressive 
         Height          =   330
         Left            =   360
         TabIndex        =   6
         Top             =   1830
         Width           =   6690
         _ExtentX        =   11800
         _ExtentY        =   582
         Caption         =   "use progressive encoding"
      End
      Begin PhotoDemon.pdCheckBox chkSubsample 
         Height          =   330
         Left            =   360
         TabIndex        =   7
         Top             =   2310
         Width           =   6690
         _ExtentX        =   11800
         _ExtentY        =   582
         Caption         =   "use specific subsampling:"
      End
      Begin PhotoDemon.pdSlider sltQuality 
         Height          =   405
         Left            =   2880
         TabIndex        =   8
         Top             =   180
         Width           =   4335
         _ExtentX        =   7223
         _ExtentY        =   873
         Min             =   1
         Max             =   99
         Value           =   90
         NotchPosition   =   1
      End
      Begin PhotoDemon.pdCheckBox chkColorMatching 
         Height          =   300
         Left            =   360
         TabIndex        =   9
         Top             =   4440
         Visible         =   0   'False
         Width           =   6690
         _ExtentX        =   11800
         _ExtentY        =   582
         Caption         =   "use perceptive color matching (slower, more accurate)"
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   285
         Index           =   3
         Left            =   120
         Top             =   3570
         Visible         =   0   'False
         Width           =   6900
         _ExtentX        =   0
         _ExtentY        =   503
         Caption         =   "automatic quality detection"
         FontSize        =   12
         ForeColor       =   4210752
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   285
         Index           =   1
         Left            =   120
         Top             =   900
         Width           =   6945
         _ExtentX        =   0
         _ExtentY        =   503
         Caption         =   "advanced quality settings"
         FontSize        =   12
         ForeColor       =   4210752
      End
   End
End
Attribute VB_Name = "dialog_ExportJPEG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'JPEG Export Dialog
'Copyright 2000-2016 by Tanner Helland
'Created: 5/8/00
'Last updated: 17/January/14
'Last update: separate metadata panel.  (See issue #113 on GitHub.)  Users can use this to override program-wide
'              metadata handling for a single image.
'
'Dialog for presenting the user various options related to JPEG exporting.  The advanced features all currently
' rely on FreeImage for implementation, and will be disabled and/or ignored if FreeImage cannot be found.
'
'IMPORTANT NOTE MARCH 2016: this dialog is being blown to bits while I sort out PD's new export engine.
'  Please ignore the dust!
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'The user input from the dialog
Private userAnswer As VbMsgBoxResult

'This form can be notified of the image being exported.  This may be used in the future to provide a preview.
Public imageBeingExported As pdImage

'The quality checkboxes work as toggles.  To prevent infinite looping while they update each other, a module-level
' variable controls access to the toggle code.
Private m_CheckBoxUpdatingDisabled As Boolean

'Final JPEG XML packet, with all JPEG settings defined as tag+value pairs
Public xmlParamString As String

'Final metadata XML packet, with all metadata settings defined as tag+value pairs
Public metadataParamString As String

'The user's answer is returned via this property
Public Property Get DialogResult() As VbMsgBoxResult
    DialogResult = userAnswer
End Property

Private Sub btsCategory_Click(ByVal buttonIndex As Long)

    Dim i As Long
    
    For i = 0 To btsCategory.ListCount - 1
        If i = buttonIndex Then
            picContainer(i).Visible = True
        Else
            picContainer(i).Visible = False
        End If
    Next i

End Sub

Private Sub chkColorMatching_Click()
    UpdatePreview
End Sub

Private Sub chkSubsample_Click()
    UpdatePreview
End Sub

'As of 16 January '14, PD can now choose a quality value for the user, using an RMSD comparison between the base image and
' its JPEG transformation.
Private Sub cmbAutoQuality_Click()
    
    If cmbAutoQuality.ListIndex > 0 Then
        cboSaveQuality.Enabled = False
        sltQuality.Enabled = False
        chkColorMatching.Enabled = True
    Else
        cboSaveQuality.Enabled = True
        sltQuality.Enabled = True
        chkColorMatching.Enabled = False
    End If
    
    UpdatePreview
    
End Sub

'QUALITY combo box - when adjusted, change the scroll bar to match
Private Sub cboSaveQuality_Click()
    
    If Not m_CheckBoxUpdatingDisabled Then
    
        Select Case cboSaveQuality.ListIndex
            
            Case 0
                sltQuality.Value = 99
                    
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

Private Sub cboSubsample_Click()
    
    'Update the specific subsampling box to match
    If Not CBool(chkSubsample) Then chkSubsample.Value = vbChecked
    UpdatePreview
    
End Sub

Private Sub cmdBar_CancelClick()
    userAnswer = vbCancel
    Me.Hide
End Sub

Private Sub cmdBar_OKClick()
    
    'Determine the compression quality for the quantization tables
    If (Not sltQuality.IsValid) Then Exit Sub
    
    'Auto-compare mode is currently deactivated for performance and reliability reasons
    'If the user has requested that PD choose a quality value for them, do so now
    'm_JPEGAutoQuality = cmbAutoQuality.ListIndex
    'Also pass along the color matching technique, which may or may not be useful
    'm_JPEGAdvancedColorMatching = CBool(chkColorMatching)
    
    Dim cParams As pdParamXML
    Set cParams = New pdParamXML
    cParams.AddParam "JPEGQuality", sltQuality.Value
    cParams.AddParam "JPEGOptimizeTables", CBool(chkOptimize)
    cParams.AddParam "JPEGProgressive", CBool(chkProgressive)
    cParams.AddParam "JPEGCustomSubsampling", CBool(chkSubsample)
    cParams.AddParam "JPEGCustomSubsamplingValue", GetSubsampleConstantFromComboBox()
    xmlParamString = cParams.GetParamString
    
    'cParams.AddParam "JPEGThumbnail", CBool(chkThumbnail)
    
    'The metadata panel manages its own XML string
    metadataParamString = mtdManager.GetMetadataSettings
    
    'Hide but *DO NOT UNLOAD* the form.  The dialog manager needs to retrieve the setting strings before unloading us
    userAnswer = vbOK
    Me.Hide
    
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub cmdBar_ResetClick()
    
    'Default is to let the user choose JPEG quality
    cmbAutoQuality.ListIndex = 0
    
    'Default save quality is "Excellent"
    cboSaveQuality.ListIndex = 1
    
    'By default, the only advanced setting is Optimize compression tables
    chkOptimize.Value = vbChecked
    'chkThumbnail.Value = vbUnchecked
    chkProgressive.Value = vbUnchecked
    chkSubsample.Value = vbUnchecked
    
    'By default, automatic color matching prefers speed over accuracy
    chkColorMatching.Value = vbUnchecked
    
    mtdManager.Reset
    
End Sub

Private Sub Form_Activate()
    UpdatePreview
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'Used to keep the "image quality" text box, scroll bar, and combo box in sync
Private Sub updateComboBox()
    
    Select Case sltQuality.Value
        
        Case 40
            If cboSaveQuality.ListIndex <> 4 Then cboSaveQuality.ListIndex = 4
                            
        Case 65
            If cboSaveQuality.ListIndex <> 3 Then cboSaveQuality.ListIndex = 3
                
        Case 80
            If cboSaveQuality.ListIndex <> 2 Then cboSaveQuality.ListIndex = 2
                
        Case 92
            If cboSaveQuality.ListIndex <> 1 Then cboSaveQuality.ListIndex = 1
                
        Case 99
            If cboSaveQuality.ListIndex <> 0 Then cboSaveQuality.ListIndex = 0
                
        Case Else
            If cboSaveQuality.ListIndex <> 5 Then cboSaveQuality.ListIndex = 5
                
    End Select
    
    If Not m_CheckBoxUpdatingDisabled Then UpdatePreview
    
End Sub

'The ShowDialog routine presents the user with this form.
Public Sub ShowDialog()

    'Provide a default answer of "cancel" (in the event that the user clicks the "x" button in the top-right)
    userAnswer = vbCancel
    
    'Make sure that the proper cursor is set
    Screen.MousePointer = 0
    
    'Populate the category button strip
    btsCategory.AddItem "Quality", 0
    btsCategory.AddItem "Metadata", 1
    
    'Populate the quality drop-down box with presets corresponding to the JPEG format
    cboSaveQuality.Clear
    cboSaveQuality.AddItem " Perfect (99)", 0
    cboSaveQuality.AddItem " Excellent (92)", 1
    cboSaveQuality.AddItem " Good (80)", 2
    cboSaveQuality.AddItem " Average (65)", 3
    cboSaveQuality.AddItem " Low (40)", 4
    cboSaveQuality.AddItem " Custom value", 5
    cboSaveQuality.ListIndex = 1
    Message "Waiting for user to specify JPEG export options... "
    
    'Populate the "auto quality" drop-down next
    cmbAutoQuality.Clear
    cmbAutoQuality.AddItem " Do not set quality automatically", 0
    cmbAutoQuality.AddItem " No noticeable differences allowed", 1
    cmbAutoQuality.AddItem " Slight differences allowed", 2
    cmbAutoQuality.AddItem " Some minor differences allowed", 3
    cmbAutoQuality.AddItem " Many minor differences allowed", 4
    cmbAutoQuality.ListIndex = 0
    
    'By default, let the user pick JPEG quality, and use sloppy color matching
    cmbAutoQuality.ToolTipText = g_Language.TranslateMessage("PhotoDemon can automatically choose a JPEG quality setting for you.  The statistical analyses it uses are designed around photographs; synthetic images or images with large regions of solid color may not work as well.")
    
    chkColorMatching.Value = vbUnchecked
    chkColorMatching.AssignTooltip "Perceptive color matching uses the CIE L*a*b* color space for highly accurate color modeling.  Enabling this setting may increase processing time by several seconds."
    
    'Populate the custom subsampling combo box as well
    cboSubsample.Clear
    cboSubsample.AddItem " 4:4:4 (best quality)", 0
    cboSubsample.AddItem " 4:2:2 (good quality)", 1
    cboSubsample.AddItem " 4:2:0 (moderate quality)", 2
    cboSubsample.AddItem " 4:1:1 (low quality)", 3
    cboSubsample.ListIndex = 2
    
    'Next, prepare various controls on the metadata panel
    mtdManager.SetParentImage imageBeingExported
    
    'By default, the quality panel is always shown.
    btsCategory.ListIndex = 0
    Dim i As Long
    For i = 0 To btsCategory.ListCount - 1
        If i = 0 Then
            picContainer(i).Visible = True
        Else
            picContainer(i).Visible = False
        End If
    Next i
    
    'If FreeImage is not available, disable all the advanced settings
    If Not g_ImageFormats.FreeImageEnabled Then
        chkOptimize.Enabled = False
        chkProgressive.Enabled = False
        chkSubsample.Enabled = False
        'chkThumbnail.Enabled = False
        cboSubsample.AddItem "n/a", 4
        cboSubsample.ListIndex = 4
        cboSubsample.Enabled = False
        lblTitle(1).Caption = "advanced settings require the FreeImage plugin"
    End If
        
    'Apply some checkbox tooltips manually (so the translation engine can find them)
    chkOptimize.AssignTooltip "Optimization is highly recommended.  This option allows the JPEG encoder to compute an optimal Huffman coding table for the file.  It does not affect image quality - only file size."
    chkProgressive.AssignTooltip "Progressive encoding is sometimes used for JPEG files that will be used on the Internet.  It saves the image in three steps, which can be used to gradually fade-in the image on a slow Internet connection."
    'chkThumbnail.AssignTooltip "Embedded thumbnails increase file size, but they help previews of the image appear more quickly in other software (e.g. Windows Explorer)."
    chkSubsample.AssignTooltip "Subsampling affects the way the JPEG encoder compresses image luminance.  4:2:0 (moderate) is the default value."
    
    'FreeImage is required to perform the JPEG transformation.  We could use GDI+, but FreeImage is
    ' much easier to interface with.  If FreeImage is not available, warn the user.
    If Not g_ImageFormats.FreeImageEnabled Then
        
        Dim tmpDIB As pdDIB
        Set tmpDIB = New pdDIB
        tmpDIB.createBlank pdFxPreview.GetPreviewWidth, pdFxPreview.GetPreviewHeight
    
        Dim notifyFont As pdFont
        Set notifyFont = New pdFont
        notifyFont.SetFontFace g_InterfaceFont
        notifyFont.SetFontSize 14
        notifyFont.SetFontColor 0
        notifyFont.SetFontBold True
        notifyFont.SetTextAlignment vbCenter
        notifyFont.CreateFontObject
        notifyFont.AttachToDC tmpDIB.getDIBDC
    
        notifyFont.FastRenderText tmpDIB.getDIBWidth \ 2, tmpDIB.getDIBHeight \ 2, g_Language.TranslateMessage("JPEG previews require the FreeImage plugin.")
        pdFxPreview.SetOriginalImage tmpDIB
        pdFxPreview.SetFXImage tmpDIB
        
        notifyFont.ReleaseFromDC
        Set tmpDIB = Nothing
        
    End If
    
    'Apply translations and visual themes
    ApplyThemeAndTranslations Me
    
    'Update the preview
    UpdatePreview
    
    'Display the dialog
    ShowPDDialog vbModal, Me, True

End Sub

Private Sub pdFxPreview_ViewportChanged()
    UpdatePreview
End Sub

Private Sub sltQuality_Change()
    If Not m_CheckBoxUpdatingDisabled Then updateComboBox
End Sub

Private Sub UpdatePreview()

    If cmdBar.PreviewsAllowed And g_ImageFormats.FreeImageEnabled And sltQuality.IsValid Then
        
        'Retrieve a composited version of the target image
        Dim tmpDIB As pdDIB
        Set tmpDIB = New pdDIB
        imageBeingExported.GetCompositedImage tmpDIB, True
        
        'Start by retrieving the relevant portion of the image, according to the preview window
        Dim tmpSafeArray As SAFEARRAY2D
        previewNonStandardImage tmpSafeArray, tmpDIB, pdFxPreview
        
        If workingDIB.getDIBColorDepth = 32 Then workingDIB.convertTo24bpp
        
        Dim JPEGQuality As Long
        JPEGQuality = sltQuality.Value
        
        'If the user wants PhotoDemon to determine a save value for them, let's do that now for the working copy.
        ' While not 100% true to the final image, it should give them a good idea of how far the compressor can go.
        If cmbAutoQuality.ListIndex > 0 Then
            JPEGQuality = FindQualityForDesiredJPEGPerception(workingDIB, cmbAutoQuality.ListIndex, CBool(chkColorMatching))
            m_CheckBoxUpdatingDisabled = True
            sltQuality.Value = JPEGQuality
            updateComboBox
            m_CheckBoxUpdatingDisabled = False
        End If
        
        'The public workingDIB object now contains the relevant portion of the preview window.  Use that to
        ' obtain a JPEG-ified version of the image data.
        FillDIBWithJPEGVersion workingDIB, workingDIB, JPEGQuality, IIf(CBool(chkSubsample), GetSubsampleConstantFromComboBox(), JPEG_SUBSAMPLING_422)
        
        'Paint the final image to screen and release all temporary objects
        finalizeNonstandardPreview pdFxPreview
                
    End If
    
End Sub

Private Function GetSubsampleConstantFromComboBox() As Long
    
    Select Case cboSubsample.ListIndex
            
        Case 0
            GetSubsampleConstantFromComboBox = JPEG_SUBSAMPLING_444
        Case 1
            GetSubsampleConstantFromComboBox = JPEG_SUBSAMPLING_422
        Case 2
            GetSubsampleConstantFromComboBox = JPEG_SUBSAMPLING_420
        Case 3
            GetSubsampleConstantFromComboBox = JPEG_SUBSAMPLING_411
                    
    End Select
    
End Function
