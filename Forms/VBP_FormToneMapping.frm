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
      TabIndex        =   3
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
      TabIndex        =   2
      Top             =   1200
      Width           =   4500
   End
   Begin PhotoDemon.smartOptionButton optToneMap 
      Height          =   375
      Index           =   1
      Left            =   4920
      TabIndex        =   4
      Top             =   2160
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   661
      Caption         =   "adaptive logarithm"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PhotoDemon.smartOptionButton optToneMap 
      Height          =   375
      Index           =   0
      Left            =   4920
      TabIndex        =   5
      Top             =   1680
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   661
      Caption         =   "linear"
      Value           =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PhotoDemon.smartOptionButton optToneMap 
      Height          =   375
      Index           =   2
      Left            =   4920
      TabIndex        =   6
      Top             =   2640
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   661
      Caption         =   "photoreceptor modeling"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PhotoDemon.smartOptionButton optToneMap 
      Height          =   375
      Index           =   3
      Left            =   4920
      TabIndex        =   7
      Top             =   3120
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   661
      Caption         =   "manual settings"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PhotoDemon.smartCheckBox chkRemember 
      Height          =   330
      Left            =   4920
      TabIndex        =   8
      Top             =   5280
      Width           =   6660
      _ExtentX        =   11748
      _ExtentY        =   582
      Caption         =   "in the future, automatically apply these settings"
      Value           =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      TabIndex        =   1
      Top             =   270
      Width           =   10440
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "tone mapping options:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Index           =   0
      Left            =   4920
      TabIndex        =   0
      Top             =   1200
      Width           =   2415
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
    
    'Ignore redraws while the dialog is not visible
    If (Not Me.Enabled) Or (Not Me.Visible) Then Exit Sub
    
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

Private Sub cmdBar_RequestPreviewUpdate()
    updatePreview
End Sub

Private Sub cmdBar_ResetClick()
    
    'By default, we want to suggest logarithmic tone mapping.  It provides good results at reasonable speed.
    optToneMap(1).Value = True
    
End Sub

Private Sub Form_Load()
    
    'Initialize a parameter assembler
    Set cParams = New pdParamString
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    'Release our mini-FreeImage handle
    If mini_FIHandle <> 0 Then FreeImage_Unload mini_FIHandle
    ReleaseFormTheming Me
    
End Sub

'Assemble the current settings into a parameter string
Private Function getToneMapParamString() As String
    getToneMapParamString = buildParams(m_curToneMapMode)
End Function

Private Sub optToneMap_Click(Index As Integer)
    m_curToneMapMode = Index
    If Me.Visible Then updatePreview
End Sub
