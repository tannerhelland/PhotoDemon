VERSION 5.00
Begin VB.Form dialog_AlphaCutoff 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Please Choose A Transparency Threshold"
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
   Begin PhotoDemon.colorSelector csComposite 
      Height          =   495
      Left            =   5040
      TabIndex        =   8
      Top             =   3900
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   873
   End
   Begin PhotoDemon.commandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   7
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
   End
   Begin PhotoDemon.sliderTextCombo sltThreshold 
      Height          =   495
      Left            =   4920
      TabIndex        =   0
      Top             =   2400
      Width           =   6615
      _ExtentX        =   11456
      _ExtentY        =   873
      Max             =   255
      Value           =   127
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
      TabIndex        =   3
      Top             =   1200
      Width           =   4500
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "background color for compositing:"
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
      Index           =   1
      Left            =   4920
      TabIndex        =   6
      Top             =   3480
      Width           =   3675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "maximum transparency "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   8640
      TabIndex        =   5
      Top             =   2940
      Width           =   1710
   End
   Begin VB.Label lblAfter 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "no transparency "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   5400
      TabIndex        =   4
      Top             =   2940
      Width           =   1230
   End
   Begin VB.Label lblWarning 
      BackStyle       =   0  'Transparent
      Caption         =   "This image has a complex alpha channel.  Before it can be saved as a paletted image (8bpp), the alpha channel must be simplified."
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
      TabIndex        =   2
      Top             =   270
      Width           =   10440
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "transparency cut-off:"
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
      TabIndex        =   1
      Top             =   2040
      Width           =   2205
   End
End
Attribute VB_Name = "dialog_AlphaCutoff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Alpha Cut-Off Dialog
'Copyright ©2012-2014 by Tanner Helland
'Created: 15/December/12
'Last updated: 29/January/14
'Last update: add an option for composite background color; with thanks to Kroc of camendesign.com for the suggestion
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

'A reference to the image being saved (actually, a temporary copy of the image being saved - but whatever).
Private srcDIB As pdDIB

'Our copy of the image being saved.  This will be created and destroyed frequently as the alpha values are updated.
Private tmpDIB As pdDIB

'Custom tooltip class allows for things like multiline, theming, and multiple monitor support
Dim m_ToolTip As clsToolTip

'The user's answer is returned via this property
Public Property Get DialogResult() As VbMsgBoxResult
    DialogResult = userAnswer
End Property

'This form can be notified of the image being exported.  This may be used in the future to provide a preview.
Public Property Let refDIB(ByRef refDIB As pdDIB)
    Set srcDIB = refDIB
End Property

'The ShowDialog routine presents the user with this form.
Public Sub showDialog()

    'Provide a default answer of "cancel" (in the event that the user clicks the "x" button in the top-right)
    userAnswer = vbCancel
    
    'Make sure that the proper cursor is set
    Screen.MousePointer = 0
        
    'Automatically draw a question icon using the system icon set
    Dim iconY As Long
    iconY = fixDPI(18)
    If g_UseFancyFonts Then iconY = iconY + fixDPI(2)
    DrawSystemIcon IDI_ASTERISK, Me.hDC, fixDPI(22), iconY
        
    'Initialize our temporary DIB render object
    Set tmpDIB = New pdDIB
        
    'Render a preview of this threshold value
    updatePreview
        
    Message "Waiting for user to specify alpha threshold... "
        
    'Assign the system hand cursor to all relevant objects
    Set m_ToolTip = New clsToolTip
    makeFormPretty Me, m_ToolTip
    
    'Display the dialog
    showPDDialog vbModal, Me, True

End Sub

'Render a preview of the current alpha cut-off to the large picture box on the form
Private Sub updatePreview()

    tmpDIB.eraseDIB
    
    tmpDIB.createFromExistingDIB srcDIB
    tmpDIB.applyAlphaCutoff sltThreshold.Value, False, csComposite.Color
    
    tmpDIB.renderToPictureBox picPreview

End Sub

'CANCEL button
Private Sub cmdBar_CancelClick()

    'Free up memory
    tmpDIB.eraseDIB
    
    'Hide the dialog and return a value of "Cancel"
    userAnswer = vbCancel
    Me.Hide

End Sub

'OK button
Private Sub cmdBar_OKClick()
    
    'Save the selected color depth to the corresponding global variable (so other functions can access it
    ' after this form is unloaded)
    g_AlphaCutoff = sltThreshold.Value
    
    'Similarly, save the selected composite color to its corresponding global variable
    g_AlphaCompositeColor = csComposite.Color
    
    'Free up memory
    tmpDIB.eraseDIB
    
    userAnswer = vbOK
    Me.Hide
        
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    updatePreview
End Sub

Private Sub cmdBar_ResetClick()
    
    'Give the transparency slider a default value of 127 instead of 0
    sltThreshold.Value = 127
    
End Sub

Private Sub csComposite_ColorChanged()
    updatePreview
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'Redraw the preview when the scroll bar is moved
Private Sub sltThreshold_Change()
    If sltThreshold.IsValid(False) Then updatePreview
End Sub
