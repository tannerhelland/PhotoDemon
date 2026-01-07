VERSION 5.00
Begin VB.Form dialog_ToneMapping 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " HDR image identified"
   ClientHeight    =   6660
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11655
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
   ScaleHeight     =   444
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   777
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.pdPictureBox picPreview 
      Height          =   4500
      Left            =   120
      Top             =   1200
      Width           =   4500
      _ExtentX        =   7938
      _ExtentY        =   7938
   End
   Begin PhotoDemon.pdCommandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5910
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   1323
      DontAutoUnloadParent=   -1  'True
   End
   Begin PhotoDemon.pdCheckBox chkRemember 
      Height          =   330
      Left            =   4920
      TabIndex        =   1
      Top             =   5490
      Width           =   6660
      _ExtentX        =   11748
      _ExtentY        =   582
      Caption         =   "in the future, automatically apply these settings"
      Value           =   0   'False
   End
   Begin PhotoDemon.pdButtonStrip btsMethod 
      Height          =   960
      Left            =   4920
      TabIndex        =   2
      Top             =   1200
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   1693
      Caption         =   "tone-mapping operator"
      FontSizeCaption =   10
   End
   Begin PhotoDemon.pdLabel lblWarning 
      Height          =   765
      Left            =   975
      Top             =   240
      Width           =   10440
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   ""
      ForeColor       =   2105376
      Layout          =   1
   End
   Begin PhotoDemon.pdPictureBox picWarning 
      Height          =   615
      Left            =   120
      Top             =   210
      Width           =   615
      _ExtentX        =   873
      _ExtentY        =   1085
   End
   Begin PhotoDemon.pdContainer picContainer 
      Height          =   3135
      Index           =   1
      Left            =   4800
      Top             =   2280
      Visible         =   0   'False
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   5530
      Begin PhotoDemon.pdSlider sltGamma 
         Height          =   690
         Index           =   2
         Left            =   120
         TabIndex        =   3
         Top             =   0
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   1244
         Caption         =   "gamma"
         FontSizeCaption =   11
         Min             =   1
         Max             =   5
         SigDigits       =   2
         Value           =   2.2
         NotchPosition   =   2
         NotchValueCustom=   2.2
      End
      Begin PhotoDemon.pdSlider sltExposure 
         Height          =   690
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   1244
         Caption         =   "exposure"
         FontSizeCaption =   11
         Min             =   0.01
         Max             =   8
         SigDigits       =   2
         Value           =   2
         NotchPosition   =   2
         NotchValueCustom=   2
      End
      Begin PhotoDemon.pdSlider sltWhitepoint 
         Height          =   690
         Left            =   120
         TabIndex        =   6
         Top             =   1680
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   1244
         Caption         =   "white point"
         FontSizeCaption =   11
         Min             =   1
         Max             =   40
         SigDigits       =   2
         Value           =   11.2
         NotchPosition   =   2
         NotchValueCustom=   11.2
      End
   End
   Begin PhotoDemon.pdContainer picContainer 
      Height          =   3135
      Index           =   0
      Left            =   4800
      Top             =   2280
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   5530
      Begin PhotoDemon.pdButtonStrip btsNormalize 
         Height          =   975
         Left            =   120
         TabIndex        =   12
         Top             =   840
         Width           =   6615
         _ExtentX        =   11456
         _ExtentY        =   1085
         Caption         =   "normalize"
         FontSizeCaption =   11
      End
      Begin PhotoDemon.pdSlider sltGamma 
         Height          =   690
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   0
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   1244
         Caption         =   "gamma"
         FontSizeCaption =   11
         Min             =   1
         Max             =   5
         SigDigits       =   2
         Value           =   2.2
         NotchPosition   =   2
         NotchValueCustom=   2.2
      End
   End
   Begin PhotoDemon.pdContainer picContainer 
      Height          =   3135
      Index           =   2
      Left            =   4800
      Top             =   2280
      Visible         =   0   'False
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   5530
      Begin PhotoDemon.pdSlider sltGamma 
         Height          =   690
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   0
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   1244
         Caption         =   "gamma"
         FontSizeCaption =   11
         Min             =   1
         Max             =   5
         SigDigits       =   2
         Value           =   1
         NotchPosition   =   2
         NotchValueCustom=   1
      End
      Begin PhotoDemon.pdSlider sltExposure 
         Height          =   690
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   1244
         Caption         =   "exposure"
         FontSizeCaption =   11
         Min             =   -8
         Max             =   8
         SigDigits       =   2
      End
   End
   Begin PhotoDemon.pdContainer picContainer 
      Height          =   3135
      Index           =   3
      Left            =   4800
      Top             =   2280
      Visible         =   0   'False
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   5530
      Begin PhotoDemon.pdSlider sltIntensity 
         Height          =   690
         Left            =   120
         TabIndex        =   9
         Top             =   0
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   1244
         Caption         =   "intensity"
         FontSizeCaption =   11
         Min             =   -4
         Max             =   4
         SigDigits       =   2
      End
      Begin PhotoDemon.pdSlider sltAdaptation 
         Height          =   690
         Left            =   120
         TabIndex        =   10
         Top             =   840
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   1244
         Caption         =   "adaptation"
         FontSizeCaption =   11
         Max             =   1
         SigDigits       =   2
         SliderTrackStyle=   1
         Value           =   1
         NotchPosition   =   2
         NotchValueCustom=   1
      End
      Begin PhotoDemon.pdSlider sltColorCorrection 
         Height          =   690
         Left            =   120
         TabIndex        =   11
         Top             =   1680
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   1244
         Caption         =   "color correction"
         FontSizeCaption =   11
         Max             =   1
         SigDigits       =   2
         SliderTrackStyle=   1
      End
   End
End
Attribute VB_Name = "dialog_ToneMapping"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Tone Mapping (e.g. high-bit-depth image import) Dialog
'Copyright 2014-2026 by Tanner Helland
'Created: 04/December/14
'Last updated: 08/Augusts/17
'Last update: migrate to XML params, many performance improvements
'
'Images with more than 8-bits per channel cannot be displayed on conventional monitors.  These images must undergo a
' process called Tone Mapping (https://en.wikipedia.org/wiki/Tone_mapping) which reduces their color count to a range
' acceptable for display.
'
'Unfortunately, there are many nuances to tone-mapping, which makes handling difficult to automate.  HDR images
' (I'm using the literal meaning here, not the .HDR file format) come in many shapes and sizes: some are pre-normalized.
' Some include infrared channels.  Some use non-floating point HDR formats, and thus require conversion to RGBF.
' Some are already gamma-corrected.  My frequent use of "some" here is not accidental - there are variations in all these
' parameters, and a lack of good metadata handling for these formats means that handling is not clear-cut.
'
'Hence the need for this dialog.  When an HDR image is loaded (including many RAW formats), and the file does not contain
' ICC data that explicitly states how to tone-map, this dialog will be triggered.  It provides a way for the user to
' control how the image is ultimately mapped to their display.
'
'The linear and filmic options are custom coded for PD.  Drago and Reinhard simply wrap the matching FreeImage.dll functions.
'
'Many thanks to Hans Nolte for his invaluable help on this topic.  (https://github.com/tannerhelland/PhotoDemon/issues/149)
'
'Thanks also to John Hable (formerly of Naughty Dog) for his great presentation on tone-mapping, including a breakdown
' of his fast and reliable Filmic Tonemapping method: http://fr.slideshare.net/ozlael/hable-john-uncharted2-hdr-lighting
' Additional filmic tonemapping references include:
'  - http://filmicgames.com/archives/190
'  - http://filmicgames.com/archives/75
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'The user input from the dialog
Private userAnswer As VbMsgBoxResult

'A copy of the FreeImage handle to the incoming image.  THIS HANDLE CANNOT BE RELEASED!  It is only provided so that we
' have a way to generate a smaller, preview-friendly image.
Private src_FIHandle As Long

'A miniaturized version of the incoming image, with the same bit-depth as the source image, but a smaller size.  It is
' computationally expensive to generate this image, so please do not destroy it.  Instead, make a temporary copy prior
' to applying tone-mapping operations for preview purposes.
Private mini_FIHandle As Long

'Theme-specific icons are fully supported
Private m_warningDIB As pdDIB

'To reduce memory churn, we reuse a temporary DIB object
Private m_tmpDIB As pdDIB

'The user's dialog answer is returned via this property
Public Property Get DialogResult() As VbMsgBoxResult
    DialogResult = userAnswer
End Property

'Tone-mapping settings are returned via this property
Public Property Get ToneMapSettings() As String
    ToneMapSettings = GetToneMapParamString()
End Property

'If the user wants us to auto-set these parameters in the future, without raising the dialog, this property will be
' set to TRUE.
Public Property Get RememberSettings() As Boolean
    RememberSettings = chkRemember.Value
End Property

'This dialog will be given access to the FreeImage handle of the image being imported.  This is crucial because only FreeImage
' can handle certain data types (e.g. unsigned 16-bit per channel images).
Public Property Let fi_HandleCopy(ByVal srcHandle As Long)
    src_FIHandle = srcHandle
End Property

'The ShowDialog routine presents the user with this form.
Public Sub ShowDialog()
    
    'Prevent preview images from rendering until all initialization has finished
    cmdBar.SetPreviewStatus False
    
    'Provide a default answer of "cancel" (in the event that the user clicks the "x" button in the top-right)
    userAnswer = vbCancel
    
    lblWarning.Caption = g_Language.TranslateMessage("This image contains more than 16 million colors.  Before it can be displayed on your screen, it must be converted to a simpler format.  This process is called tone mapping.")
    Screen.MousePointer = 0
    
    'Create a small copy of the image, for preview purposes.  Tone-mapping can be hideously slow, so we'll want to limit the size
    ' of the image in question.  Note that we do not free the source handle - we still need it for the loading process!!
    ' (Also, we should check the case of the FreeImage handle being 0, as that will cause uncatchable crashes.)
    Dim newWidth As Long, newHeight As Long
    If (src_FIHandle <> 0) Then
        PDMath.ConvertAspectRatio FreeImage_GetWidth(src_FIHandle), FreeImage_GetHeight(src_FIHandle), picPreview.GetWidth * 2, picPreview.GetHeight * 2, newWidth, newHeight
        If (FreeImage_GetWidth(src_FIHandle) < newWidth) Or (FreeImage_GetHeight(src_FIHandle) < newHeight) Then
            newWidth = FreeImage_GetWidth(src_FIHandle)
            newHeight = FreeImage_GetHeight(src_FIHandle)
        End If
        mini_FIHandle = Outside_FreeImageV3.FreeImage_Rescale(src_FIHandle, newWidth, newHeight, FILTER_CATMULLROM)
    End If
    
    'Render a preview of the current settings, if any
    cmdBar.SetPreviewStatus True
        
    Message "Waiting for tone mapping instructions..."
    
    'Apply translations and visual themes
    ApplyThemeAndTranslations Me
    
    'Generate a preview
    UpdatePreview
    
    'Display the dialog
    ShowPDDialog vbModal, Me, True

End Sub

'Render a preview of the current alpha cut-off to the large picture box on the form
Private Sub UpdatePreview()
    
    'Ignore redraws while the dialog is not visible or disabled
    If (Not Me.Enabled) Or (Not cmdBar.PreviewsAllowed) Then Exit Sub
    
    'As a failsafe against rapid clicking by the user, disable the form prior to applying any tone-mapping operations
    Me.Enabled = False
    
    Dim tmp_FIHandle As Long
    
    'Retrieve a tone-mapped image, using the central tone-map function
    If (mini_FIHandle <> 0) Then tmp_FIHandle = Plugin_FreeImage.ApplyToneMapping(mini_FIHandle, GetToneMapParamString())
    
    'If successful, create a pdDIB copy, render it to the screen, then kill our temporary FreeImage handle
    If (tmp_FIHandle <> 0) Then
        
        If (m_tmpDIB Is Nothing) Then Set m_tmpDIB = New pdDIB
        If Plugin_FreeImage.GetPDDibFromFreeImageHandle(tmp_FIHandle, m_tmpDIB) Then
            
            'Premultiply as necessary
            If (m_tmpDIB.GetDIBColorDepth = 32) Then m_tmpDIB.SetAlphaPremultiplication True
            
            'Render to the screen
            picPreview.CopyDIB m_tmpDIB, , True, , True
            
        Else
            Debug.Print "Can't preview tone-mapping; could not create pdDIB object."
        End If
        
        'Release our FreeImage handle
        FreeImage_Unload tmp_FIHandle
    
    'Tone mapping failed; abandon the preview attempt
    Else
        PDDebug.LogAction "Can't preview tone-mapping; unspecified error returned by central tone-map function."
    End If
    
    'Re-enable the form
    Me.Enabled = True
    
End Sub

Private Sub btsMethod_Click(ByVal buttonIndex As Long)
    
    Dim i As Long
    For i = 0 To picContainer.UBound
        picContainer(i).Visible = (i = buttonIndex)
    Next i
    
    UpdatePreview
    
End Sub

Private Sub btsNormalize_Click(ByVal buttonIndex As Long)
    UpdatePreview
End Sub

'CANCEL button
Private Sub cmdBar_CancelClick()
    Message vbNullString
    userAnswer = vbCancel
    Me.Hide
End Sub

'OK button
Private Sub cmdBar_OKClick()
    Message "Proceeding with tone mapping..."
    userAnswer = vbOK
    Me.Hide
End Sub

Private Sub cmdBar_RandomizeClick()
    chkRemember.Value = False
End Sub

Private Sub cmdBar_ReadCustomPresetData()
    chkRemember.Value = False
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub cmdBar_ResetClick()
    
    cmdBar.SetPreviewStatus False
    btsMethod.ListIndex = 0
    sltGamma(0) = 2.2
    sltGamma(1) = 1#        'FreeImage documentation is unclear on the correct behavior for Drago gamma.  2.2 is recommended as
                            ' a "starting place", but this seems to blow out the image, so I'm changing PD to recommend 1.0 as
                            ' the default Drago value.
    sltGamma(2) = 2.2
    sltExposure(1) = 2#
    sltWhitepoint = 11.2
    chkRemember.Value = False
    cmdBar.SetPreviewStatus True
    
    UpdatePreview
    
End Sub

Private Sub Form_Load()
    
    'Initialize any/all controls
    btsMethod.AddItem "Linear", 0
    btsMethod.AddItem "Filmic", 1
    btsMethod.AddItem "Drago", 2
    btsMethod.AddItem "Reinhard", 3
    btsMethod.ListIndex = 0
    
    btsNormalize.AddItem "none", 0
    btsNormalize.AddItem "visible spectrum", 1
    btsNormalize.AddItem "full spectrum", 2
    
    'Prep a warning icon
    Dim warningIconSize As Long
    warningIconSize = Interface.FixDPI(32)
    
    If Not IconsAndCursors.LoadResourceToDIB("generic_warning", m_warningDIB, warningIconSize, warningIconSize, 0) Then
        Set m_warningDIB = Nothing
        picWarning.Visible = False
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    'Release our mini-FreeImage handle
    If (mini_FIHandle <> 0) Then FreeImage_Unload mini_FIHandle
    ReleaseFormTheming Me
    
End Sub

'Assemble the current settings into a parameter string
Private Function GetToneMapParamString() As String
    
    'Like everything else in PD, parameters are returned as XML strings
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    
    With cParams
        
        'As you can imagine, the parameters we care about are highly variable, depending on the user's selected values.
        ' To keep this function simple, we potentially write out parameters that may not be required, but the load
        ' function doesn't care because it just grabs the parameters relevant to the active tone-mapping method.
        .AddParam "method", btsMethod.ListIndex
        
        If (btsMethod.ListIndex = PDTM_LINEAR) Then
        
            .AddParam "gamma", sltGamma(0).Value
            
            'Normalization is a little weird, because it controls two values in the destination
            If (btsNormalize.ListIndex = 0) Then
                .AddParam "normalize", PD_BOOL_FALSE
                .AddParam "ignorenegative", False
            ElseIf (btsNormalize.ListIndex = 1) Then
                .AddParam "normalize", PD_BOOL_AUTO
                .AddParam "ignorenegative", True
            Else
                .AddParam "normalize", PD_BOOL_TRUE
                .AddParam "ignorenegative", False
            End If
            
        ElseIf (btsMethod.ListIndex = PDTM_FILMIC) Then
            .AddParam "gamma", sltGamma(2).Value
            .AddParam "exposure", sltExposure(1).Value
            .AddParam "whitepoint", sltWhitepoint.Value
        
        ElseIf (btsMethod.ListIndex = PDTM_DRAGO) Then
            .AddParam "gamma", sltGamma(1).Value
            .AddParam "exposure", sltExposure(0).Value
        
        ElseIf (btsMethod.ListIndex = PDTM_REINHARD) Then
            .AddParam "intensity", sltIntensity.Value
            .AddParam "adaptation", sltAdaptation.Value
            .AddParam "colorcorrection", sltColorCorrection.Value
        
        End If
    
    End With
    
    GetToneMapParamString = cParams.GetParamString()
    
End Function

Private Sub picPreview_DrawMe(ByVal targetDC As Long, ByVal ctlWidth As Long, ByVal ctlHeight As Long)
    If (Not m_tmpDIB Is Nothing) Then picPreview.CopyDIB m_tmpDIB, True, True, False, True
End Sub

Private Sub sltAdaptation_Change()
    UpdatePreview
End Sub

Private Sub sltColorCorrection_Change()
    UpdatePreview
End Sub

Private Sub sltExposure_Change(Index As Integer)
    UpdatePreview
End Sub

Private Sub sltGamma_Change(Index As Integer)
    UpdatePreview
End Sub

Private Sub sltIntensity_Change()
    UpdatePreview
End Sub

Private Sub sltWhitepoint_Change()
    UpdatePreview
End Sub

Private Sub picWarning_DrawMe(ByVal targetDC As Long, ByVal ctlWidth As Long, ByVal ctlHeight As Long)
    GDI.FillRectToDC targetDC, 0, 0, ctlWidth, ctlHeight, g_Themer.GetGenericUIColor(UI_Background)
    If (Not m_warningDIB Is Nothing) Then m_warningDIB.AlphaBlendToDC targetDC, , (ctlWidth - m_warningDIB.GetDIBWidth) \ 2, (ctlHeight - m_warningDIB.GetDIBHeight) \ 2
End Sub
