VERSION 5.00
Begin VB.Form dialog_AlphaCutoff 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Please Choose A Transparency Threshold"
   ClientHeight    =   9135
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7035
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
   ScaleHeight     =   609
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   469
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin PhotoDemon.sliderTextCombo sltThreshold 
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   7080
      Width           =   6495
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
      Height          =   5100
      Left            =   615
      ScaleHeight     =   338
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   382
      TabIndex        =   5
      Top             =   1200
      Width           =   5760
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   5550
      TabIndex        =   1
      Top             =   8520
      Width           =   1365
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   4080
      TabIndex        =   0
      Top             =   8520
      Width           =   1365
   End
   Begin VB.Label lblBackground 
      Height          =   855
      Left            =   0
      TabIndex        =   8
      Top             =   8370
      Width           =   7095
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
      Left            =   3960
      TabIndex        =   7
      Top             =   7560
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
      Left            =   720
      TabIndex        =   6
      Top             =   7560
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
      TabIndex        =   4
      Top             =   270
      Width           =   5775
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
      Left            =   240
      TabIndex        =   3
      Top             =   6720
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
'Copyright ©2012-2013 by Tanner Helland
'Created: 15/December/12
'Last updated: 24/April/13
'Last update: rewrote the dialog against the new slider/text combo control
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
' projects IF you provide attribution.  For more information, please visit http://www.tannerhelland.com/photodemon/#license
'
'***************************************************************************

Option Explicit

'The user input from the dialog
Private userAnswer As VbMsgBoxResult

'A reference to the image being saved (actually, a temporary copy of the image being saved - but whatever).
Private srcLayer As pdLayer

'Our copy of the image being saved.  This will be created and destroyed frequently as the alpha values are updated.
Private tmpLayer As pdLayer

'Custom tooltip class allows for things like multiline, theming, and multiple monitor support
Dim m_ToolTip As clsToolTip

'The user's answer is returned via this property
Public Property Get DialogResult() As VbMsgBoxResult
    DialogResult = userAnswer
End Property

'This form can be notified of the image being exported.  This may be used in the future to provide a preview.
Public Property Let refLayer(ByRef refLayer As pdLayer)
    Set srcLayer = refLayer
End Property

'CANCEL button
Private Sub CmdCancel_Click()
    
    'Free up memory
    tmpLayer.eraseLayer
    
    userAnswer = vbCancel
    Me.Hide
    
End Sub

'OK button
Private Sub CmdOK_Click()
        
    'Make sure the input value is valid before continuing
    If sltThreshold.IsValid Then
        
        'Save the selected color depth to the corresponding global variable (so other functions can access it
        ' after this form is unloaded)
        g_AlphaCutoff = sltThreshold.Value
        
        'Free up memory
        tmpLayer.eraseLayer
        
        userAnswer = vbOK
        Me.Hide
        
    End If
    
End Sub

'The ShowDialog routine presents the user with this form.
Public Sub ShowDialog()

    'Provide a default answer of "cancel" (in the event that the user clicks the "x" button in the top-right)
    userAnswer = vbCancel
    
    'Make sure that the proper cursor is set
    Screen.MousePointer = 0
        
    'Automatically draw a question icon using the system icon set
    Dim iconY As Long
    iconY = 18
    If g_UseFancyFonts Then iconY = iconY + 2
    DrawSystemIcon IDI_ASTERISK, Me.hDC, 22, iconY
        
    'Initialize our temporary layer render object
    Set tmpLayer = New pdLayer
        
    'Render a preview of this threshold value
    renderPreview
        
    Message "Waiting for user to specify alpha threshold... "
        
    'Assign the system hand cursor to all relevant objects
    Set m_ToolTip = New clsToolTip
    makeFormPretty Me, m_ToolTip
    
    'Display the dialog
    Me.Show vbModal, FormMain

End Sub

'Render a preview of the current alpha cut-off to the large picture box on the form
Private Sub renderPreview()

    tmpLayer.eraseLayer
    
    tmpLayer.createFromExistingLayer srcLayer
    tmpLayer.applyAlphaCutoff sltThreshold.Value, False
    
    DrawPreviewImage picPreview, True, tmpLayer

End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'Redraw the preview when the scroll bar is moved
Private Sub sltThreshold_Change()
    If sltThreshold.IsValid(False) Then renderPreview
End Sub
