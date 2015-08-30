VERSION 5.00
Begin VB.Form dialog_ToneMapping 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " HDR image identified"
   ClientHeight    =   6660
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11655
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
   ScaleHeight     =   444
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   777
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.commandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   2
      Top             =   5910
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   1323
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14802140
      dontAutoUnloadParent=   -1  'True
   End
   Begin VB.PictureBox picPreview 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4500
      Left            =   120
      ScaleHeight     =   298
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   298
      TabIndex        =   1
      Top             =   1200
      Width           =   4500
   End
   Begin PhotoDemon.smartCheckBox chkRemember 
      Height          =   330
      Left            =   4920
      TabIndex        =   3
      Top             =   5490
      Width           =   6660
      _ExtentX        =   11748
      _ExtentY        =   582
      Caption         =   "in the future, automatically apply these settings"
      Value           =   0
   End
   Begin PhotoDemon.buttonStrip btsMethod 
      Height          =   720
      Left            =   4920
      TabIndex        =   4
      Top             =   1200
      Width           =   6615
      _ExtentX        =   9790
      _ExtentY        =   1058
   End
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3375
      Index           =   0
      Left            =   4800
      ScaleHeight     =   225
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   457
      TabIndex        =   5
      Top             =   2040
      Width           =   6855
      Begin PhotoDemon.sliderTextCombo sltGamma 
         Height          =   705
         Index           =   0
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
         Value           =   2.2
         NotchPosition   =   2
         NotchValueCustom=   2.2
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   330
         Index           =   9
         Left            =   120
         Top             =   960
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   582
         Caption         =   "normalize"
         FontSize        =   11
      End
      Begin PhotoDemon.smartOptionButton optNormalize 
         Height          =   330
         Index           =   0
         Left            =   360
         TabIndex        =   18
         Top             =   1380
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   582
         Caption         =   "none"
         Value           =   -1  'True
      End
      Begin PhotoDemon.smartOptionButton optNormalize 
         Height          =   330
         Index           =   1
         Left            =   360
         TabIndex        =   19
         Top             =   1740
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   582
         Caption         =   "visible spectrum"
      End
      Begin PhotoDemon.smartOptionButton optNormalize 
         Height          =   330
         Index           =   2
         Left            =   360
         TabIndex        =   20
         Top             =   2100
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   582
         Caption         =   "full spectrum"
      End
   End
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3375
      Index           =   2
      Left            =   4800
      ScaleHeight     =   225
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   457
      TabIndex        =   8
      Top             =   2040
      Visible         =   0   'False
      Width           =   6855
      Begin PhotoDemon.sliderTextCombo sltGamma 
         Height          =   705
         Index           =   1
         Left            =   120
         TabIndex        =   9
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
      Begin PhotoDemon.sliderTextCombo sltExposure 
         Height          =   705
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   960
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
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3375
      Index           =   3
      Left            =   4800
      ScaleHeight     =   225
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   457
      TabIndex        =   6
      Top             =   2040
      Visible         =   0   'False
      Width           =   6855
      Begin PhotoDemon.sliderTextCombo sltIntensity 
         Height          =   705
         Left            =   120
         TabIndex        =   11
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
      Begin PhotoDemon.sliderTextCombo sltAdaptation 
         Height          =   705
         Left            =   120
         TabIndex        =   12
         Top             =   960
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
      Begin PhotoDemon.sliderTextCombo sltColorCorrection 
         Height          =   705
         Left            =   120
         TabIndex        =   13
         Top             =   1920
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
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3375
      Index           =   1
      Left            =   4800
      ScaleHeight     =   225
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   457
      TabIndex        =   14
      Top             =   2040
      Visible         =   0   'False
      Width           =   6855
      Begin PhotoDemon.sliderTextCombo sltGamma 
         Height          =   705
         Index           =   2
         Left            =   120
         TabIndex        =   15
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
      Begin PhotoDemon.sliderTextCombo sltExposure 
         Height          =   705
         Index           =   1
         Left            =   120
         TabIndex        =   16
         Top             =   960
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
      Begin PhotoDemon.sliderTextCombo sltWhitepoint 
         Height          =   705
         Left            =   120
         TabIndex        =   17
         Top             =   1920
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
   Begin VB.Label lblWarning 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00202020&
      Height          =   765
      Left            =   975
      TabIndex        =   0
      Top             =   240
      Width           =   10440
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "dialog_ToneMapping"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Tone Mapping (e.g. high-bit-depth image import) Dialog
'Copyright 2014-2015 by Tanner Helland
'Created: 04/December/14
'Last updated: 07/December/14
'Last update: add new filmic tone-mapping mode; see http://fr.slideshare.net/ozlael/hable-john-uncharted2-hdr-lighting for details
'
'Images with more than 8-bits per channel cannot be displayed on conventional monitors.  These images must undergo a
' process called Tone Mapping (http://en.wikipedia.org/wiki/Tone_mapping) which reduces their color count to a range
' acceptable for display.
'
'Unfortunately, there are many nuances to tone-mapping, which makes handling difficult to automate.  HDR images
' (I'm using the literal meaning here, not the literal HDR format) come in many shapes and sizes: some are pre-normalized.
' Some include infrared channels.  Some are in non-floating point HDR formats, and thus require conversion to RGBF.
' Some are already gamma corrected.  My frequent use of "some" here is not accidental - there are variations in all these
' parameters, and a lack of good metadata handling for these formats means that handling is not clear-cut.
'
'Hence the need for this dialog.  When an HDR image is loaded (including many RAW formats), this dialog will be triggered,
' giving the user a way to control how the image is mapped to their display.
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
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
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

'Param string assembler
Private cParams As pdParamString

'Current tone-mapping mode (set by clicking an option button)
Private m_curToneMapMode As Long

'The user's dialog answer is returned via this property
Public Property Get DialogResult() As VbMsgBoxResult
    DialogResult = userAnswer
End Property

'Tone-mapping settings are returned via this property
Public Property Get toneMapSettings() As String
    toneMapSettings = getToneMapParamString()
End Property

'If the user wants us to auto-set these parameters in the future, without raising the dialog, this property will be
' set to TRUE.
Public Property Get RememberSettings() As Boolean
    RememberSettings = CBool(chkRemember.Value)
End Property

'This dialog will be given access to the FreeImage handle of the image being imported.  This is crucial because only FreeImage
' can handle certain data types (e.g. unsigned 16-bit per channel images).
Public Property Let fi_HandleCopy(ByVal srcHandle As Long)
    src_FIHandle = srcHandle
End Property

'The ShowDialog routine presents the user with this form.
Public Sub showDialog()
    
    'Prevent preview images from rendering until all initialization has finished
    cmdBar.markPreviewStatus False
    
    'Provide a default answer of "cancel" (in the event that the user clicks the "x" button in the top-right)
    userAnswer = vbCancel
    
    lblWarning.Caption = g_Language.TranslateMessage("This image contains more than 16 million colors.  Before it can be displayed on your screen, it must be converted to a simpler format.  This process is called tone mapping.")
    Screen.MousePointer = 0
        
    'Automatically draw a question icon using the system icon set
    Dim iconY As Long
    iconY = fixDPI(18)
    If g_UseFancyFonts Then iconY = iconY + fixDPI(2)
    DrawSystemIcon IDI_ASTERISK, Me.hDC, fixDPI(22), iconY
    
    'Create a small copy of the image, for preview purposes.  Tone-mapping can be hideously slow, so we'll want to limit the size
    ' of the image in question.  Note that we do not free the source handle - we still need it for the loading process!!
    ' (Also, we should check the case of the FreeImage handle being 0, as that will cause uncatchable crashes.)
    Dim newWidth As Long, newHeight As Long
    If src_FIHandle <> 0 Then
        convertAspectRatio FreeImage_GetWidth(src_FIHandle), FreeImage_GetHeight(src_FIHandle), picPreview.ScaleWidth * 2, picPreview.ScaleHeight * 2, newWidth, newHeight
        mini_FIHandle = Outside_FreeImageV3.FreeImage_Rescale(src_FIHandle, newWidth, newHeight, FILTER_CATMULLROM)
    End If
    
    'Render a preview of the current settings, if any
    cmdBar.markPreviewStatus True
    updatePreview
        
    Message "Waiting for tone mapping instructions..."
    
    'Apply translations and visual themes
    makeFormPretty Me
    
    'Display the dialog
    showPDDialog vbModal, Me, True

End Sub

'Render a preview of the current alpha cut-off to the large picture box on the form
Private Sub updatePreview()
    
    'Ignore redraws while the dialog is not visible or disabled
    If (Not Me.Enabled) Or (Not Me.Visible) Or (Not cmdBar.previewsAllowed) Then Exit Sub
    
    'As a failsafe against rapid clicking by the user, disable the form prior to applying any tone-mapping operations
    Me.Enabled = False
    
    Dim tmp_FIHandle As Long
    
    'Retrieve a tone-mapped image, using the master tone-map function
    If mini_FIHandle <> 0 Then
        tmp_FIHandle = Plugin_FreeImage_Interface.applyToneMapping(mini_FIHandle, getToneMapParamString())
    End If
    
    'If successful, create a pdDIB copy, render it to the screen, then kill our temporary FreeImage handle
    If tmp_FIHandle <> 0 Then
    
        Dim tmpDIB As pdDIB
        Set tmpDIB = New pdDIB
        If Plugin_FreeImage_Interface.getPDDibFromFreeImageHandle(tmp_FIHandle, tmpDIB) Then
            
            'Premultiply as necessary
            If tmpDIB.getDIBColorDepth = 32 Then tmpDIB.setAlphaPremultiplication True
            tmpDIB.renderToPictureBox picPreview
            
            'Release our DIB
            Set tmpDIB = Nothing
            
        Else
            Debug.Print "Can't preview tone-mapping; could not create pdDIB object."
        End If
        
        'Release our FreeImage handle
        FreeImage_Unload tmp_FIHandle
    
    'Tone mapping failed; abandon the preview attempt
    Else
        Debug.Print "Can't preview tone-mapping; unspecified error returned by master tone-map function."
    End If
    
    'Re-enable the form
    Me.Enabled = True
    
End Sub

Private Sub btsMethod_Click(ByVal buttonIndex As Long)
    
    Dim i As Long
    For i = 0 To picContainer.UBound
        If i = buttonIndex Then
            picContainer(i).Visible = True
        Else
            picContainer(i).Visible = False
        End If
    Next i
    
    updatePreview
    
End Sub

'CANCEL button
Private Sub cmdBar_CancelClick()
    
    Message "Cancelling image import..."
    
    'Hide the dialog and return a value of "Cancel"
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
    chkRemember.Value = vbUnchecked
End Sub

Private Sub cmdBar_ReadCustomPresetData()
    chkRemember.Value = vbUnchecked
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    updatePreview
End Sub

Private Sub cmdBar_ResetClick()
    
    cmdBar.markPreviewStatus False
    btsMethod.ListIndex = 0
    sltGamma(0) = 2.2
    sltGamma(1) = 1#        'FreeImage documentation is unclear on the correct behavior for Drago gamma.  2.2 is recommended as
                            ' a "starting place", but this seems to blow out the image, so I'm changing PD to recommend 1.0 as
                            ' the default Drago value.
    sltGamma(2) = 2.2
    sltExposure(1) = 2#
    sltWhitepoint = 11.2
    chkRemember.Value = vbUnchecked
    cmdBar.markPreviewStatus True
    
    updatePreview
    
End Sub

Private Sub Form_Load()
    
    'Initialize a parameter assembler
    Set cParams = New pdParamString
    
    'Initialize any/all controls
    btsMethod.AddItem "Linear", 0
    btsMethod.AddItem "Filmic", 1
    btsMethod.AddItem "Drago", 2
    btsMethod.AddItem "Reinhard", 3
    btsMethod.ListIndex = 0
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    'Release our mini-FreeImage handle
    If mini_FIHandle <> 0 Then FreeImage_Unload mini_FIHandle
    ReleaseFormTheming Me
    
End Sub

'Assemble the current settings into a parameter string
Private Function getToneMapParamString() As String
    
    'The param string's functions are highly variable, depending on the selected values
    Dim vParams() As Variant
    ReDim vParams(0 To 4) As Variant
    
    'First comes the actual conversion method
    vParams(0) = btsMethod.ListIndex
    
    'Subsequent parameters vary by method
    Select Case vParams(0)
    
        Case PDTM_LINEAR
            vParams(1) = sltGamma(0).Value
            
            'Normalization is a little weird, because it controls two values in the destination
            If optNormalize(0) Then
                vParams(2) = PD_BOOL_FALSE
                vParams(3) = False
            ElseIf optNormalize(1) Then
                vParams(2) = PD_BOOL_AUTO
                vParams(3) = True
            Else
                vParams(2) = PD_BOOL_TRUE
                vParams(3) = False
            End If
        
        Case PDTM_FILMIC
            vParams(1) = sltGamma(2).Value
            vParams(2) = sltExposure(1).Value
            vParams(3) = sltWhitepoint.Value
        
        Case PDTM_DRAGO
            vParams(1) = sltGamma(1).Value
            vParams(2) = sltExposure(0).Value
        
        Case PDTM_REINHARD
            vParams(1) = sltIntensity.Value
            vParams(2) = sltAdaptation.Value
            vParams(3) = sltColorCorrection.Value
        
    End Select
    
    getToneMapParamString = buildParams(vParams(0), vParams(1), vParams(2), vParams(3), vParams(4))
    
End Function

Private Sub optNormalize_Click(Index As Integer)
    updatePreview
End Sub

Private Sub sltAdaptation_Change()
    updatePreview
End Sub

Private Sub sltColorCorrection_Change()
    updatePreview
End Sub

Private Sub sltExposure_Change(Index As Integer)
    updatePreview
End Sub

Private Sub sltGamma_Change(Index As Integer)
    updatePreview
End Sub

Private Sub sltIntensity_Change()
    updatePreview
End Sub

Private Sub sltWhitepoint_Change()
    updatePreview
End Sub
