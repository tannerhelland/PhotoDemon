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
      _extentx        =   20558
      _extenty        =   1323
      font            =   "VBP_FormToneMapping.frx":0000
      dontautounloadparent=   -1
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
      _extentx        =   11748
      _extenty        =   582
      caption         =   "in the future, automatically apply these settings"
      font            =   "VBP_FormToneMapping.frx":0028
      value           =   2
   End
   Begin PhotoDemon.buttonStrip btsMethod 
      Height          =   720
      Left            =   4920
      TabIndex        =   4
      Top             =   1200
      Width           =   6615
      _extentx        =   9790
      _extenty        =   1058
      font            =   "VBP_FormToneMapping.frx":0050
   End
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3255
      Index           =   2
      Left            =   4800
      ScaleHeight     =   217
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   457
      TabIndex        =   6
      Top             =   2040
      Visible         =   0   'False
      Width           =   6855
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   330
         Index           =   4
         Left            =   120
         Top             =   0
         Width           =   6495
         _extentx        =   11456
         _extenty        =   582
         caption         =   "intensity"
         fontsize        =   11
      End
      Begin PhotoDemon.sliderTextCombo sltIntensity 
         Height          =   495
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   6615
         _extentx        =   11668
         _extenty        =   873
         forecolor       =   0
         min             =   -4
         max             =   4
         sigdigits       =   2
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   330
         Index           =   5
         Left            =   120
         Top             =   960
         Width           =   6495
         _extentx        =   11456
         _extenty        =   582
         caption         =   "adaptation"
         fontsize        =   11
      End
      Begin PhotoDemon.sliderTextCombo sltAdaptation 
         Height          =   495
         Left            =   120
         TabIndex        =   12
         Top             =   1320
         Width           =   6615
         _extentx        =   11456
         _extenty        =   873
         forecolor       =   0
         max             =   1
         sigdigits       =   2
         slidertrackstyle=   1
         value           =   1
         notchposition   =   2
         notchvaluecustom=   1
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   330
         Index           =   1
         Left            =   120
         Top             =   1920
         Width           =   6495
         _extentx        =   11456
         _extenty        =   582
         caption         =   "color correction"
         fontsize        =   11
      End
      Begin PhotoDemon.sliderTextCombo sltColorCorrection 
         Height          =   495
         Left            =   120
         TabIndex        =   13
         Top             =   2280
         Width           =   6615
         _extentx        =   11456
         _extenty        =   873
         forecolor       =   0
         max             =   1
         sigdigits       =   2
         slidertrackstyle=   1
      End
   End
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3255
      Index           =   0
      Left            =   4800
      ScaleHeight     =   217
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   457
      TabIndex        =   5
      Top             =   2040
      Width           =   6855
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   330
         Index           =   0
         Left            =   120
         Top             =   0
         Width           =   6495
         _extentx        =   11456
         _extenty        =   582
         caption         =   "gamma"
         fontsize        =   11
      End
      Begin PhotoDemon.sliderTextCombo sltGamma 
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   6615
         _extentx        =   11668
         _extenty        =   873
         forecolor       =   0
         min             =   1
         max             =   5
         sigdigits       =   2
         value           =   2.2
         notchposition   =   2
         notchvaluecustom=   2.2
      End
   End
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3255
      Index           =   1
      Left            =   4800
      ScaleHeight     =   217
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   457
      TabIndex        =   8
      Top             =   2040
      Width           =   6855
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   330
         Index           =   2
         Left            =   120
         Top             =   0
         Width           =   6495
         _extentx        =   11456
         _extenty        =   582
         caption         =   "gamma"
         fontsize        =   11
      End
      Begin PhotoDemon.sliderTextCombo sltGamma 
         Height          =   495
         Index           =   1
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   6615
         _extentx        =   11668
         _extenty        =   873
         forecolor       =   0
         min             =   1
         max             =   5
         sigdigits       =   2
         value           =   2.2
         notchposition   =   2
         notchvaluecustom=   2.2
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   330
         Index           =   3
         Left            =   120
         Top             =   960
         Width           =   6495
         _extentx        =   11456
         _extenty        =   582
         caption         =   "exposure"
         fontsize        =   11
      End
      Begin PhotoDemon.sliderTextCombo sltExposure 
         Height          =   495
         Left            =   120
         TabIndex        =   10
         Top             =   1320
         Width           =   6615
         _extentx        =   11668
         _extenty        =   873
         forecolor       =   0
         min             =   -8
         max             =   8
         sigdigits       =   2
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
'Copyright ©2014-2015 by Tanner Helland
'Created: 04/December/14
'Last updated: 04/December/14
'Last update: merge existing tone-mapping code into this dialog, and expose options to the user.
'
'Dialog for presenting the user a choice of alpha cut-off.  When reducing complex (32bpp)
' alpha channels to the simple ones required by 8bpp images, there is no fool-proof
' heuristic for maximizing quality.  In these cases, some user intervention is required
' to inspect the image and make sure everything looks acceptable.
'
'Thus this dialog.  It should only be called when a 32bpp image has a non-binary alpha
' channel.  The individual save functions automatically check for binary alpha channels,
' and if one is found, it handles the alpha-cutoff on its own (on account of there only
' being "fully transparent" and "fully opaque" pixels).
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

'Custom tooltip class allows for things like multiline, theming, and multiple monitor support
Dim m_ToolTip As clsToolTip

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
        
    'Assign the system hand cursor to all relevant objects
    Set m_ToolTip = New clsToolTip
    makeFormPretty Me, m_ToolTip
    
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
        tmp_FIHandle = Plugin_FreeImage_Expanded_Interface.applyToneMapping(mini_FIHandle, getToneMapParamString())
    End If
    
    'If successful, create a pdDIB copy, render it to the screen, then kill our temporary FreeImage handle
    If tmp_FIHandle <> 0 Then
    
        Dim tmpDIB As pdDIB
        Set tmpDIB = New pdDIB
        If Plugin_FreeImage_Expanded_Interface.getPDDibFromFreeImageHandle(tmp_FIHandle, tmpDIB) Then
            
            'Premultiply as necessary
            If tmpDIB.getDIBColorDepth = 32 Then tmpDIB.fixPremultipliedAlpha True
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
    sltGamma(1) = 2.2
    chkRemember.Value = vbUnchecked
    cmdBar.markPreviewStatus True
    
    updatePreview
    
End Sub

Private Sub Form_Load()
    
    'Initialize a parameter assembler
    Set cParams = New pdParamString
    
    'Initialize any/all controls
    btsMethod.AddItem "linear", 0
    btsMethod.AddItem "Drago", 1
    btsMethod.AddItem "Reinhard", 2
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
    Dim strParams() As Variant
    ReDim strParams(0 To 4) As Variant
    
    'First comes the actual conversion method
    strParams(0) = btsMethod.ListIndex
    
    'Subsequent parameters vary by method
    Select Case strParams(0)
    
        Case PDTM_LINEAR
            strParams(1) = sltGamma(0).Value
            
            'If normalize is exposed in the future, set it as the second parameter; the tone-mapping function is
            ' automatically prepared to operate on this value if supplied.
            strParams(2) = 0
        
        Case PDTM_ADAPTIVE_LOGARITHMIC
            strParams(1) = sltGamma(1).Value
            strParams(2) = sltExposure.Value
        
        Case PDTM_PHOTORECEPTOR
            strParams(1) = sltIntensity.Value
            strParams(2) = sltAdaptation.Value
            strParams(3) = sltColorCorrection.Value
        
        Case PDTM_MANUAL
    
    End Select
    
    getToneMapParamString = buildParams(strParams(0), strParams(1), strParams(2), strParams(3), strParams(4))
    
End Function

Private Sub sltAdaptation_Change()
    updatePreview
End Sub

Private Sub sltColorCorrection_Change()
    updatePreview
End Sub

Private Sub sltExposure_Change()
    updatePreview
End Sub

Private Sub sltGamma_Change(Index As Integer)
    updatePreview
End Sub

Private Sub sltIntensity_Change()
    updatePreview
End Sub
