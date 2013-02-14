VERSION 5.00
Begin VB.Form dialog_FirstRun 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Please select which language to use"
   ClientHeight    =   9135
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   14070
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
   ScaleWidth      =   938
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lstLanguages 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   2460
      Left            =   360
      TabIndex        =   5
      Top             =   1560
      Width           =   7335
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   5550
      TabIndex        =   1
      Top             =   5640
      Width           =   1365
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   4080
      TabIndex        =   0
      Top             =   5640
      Width           =   1365
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000D&
      X1              =   8
      X2              =   528
      Y1              =   56
      Y2              =   56
   End
   Begin VB.Label lblAvailableLanguages 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "currently available languages:"
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
      Left            =   240
      TabIndex        =   6
      Top             =   1080
      Width           =   3150
   End
   Begin VB.Label lblBackground 
      Height          =   855
      Left            =   0
      TabIndex        =   4
      Top             =   5490
      Width           =   7095
   End
   Begin VB.Label lblIntro 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "please select your preferred language"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00202020&
      Height          =   285
      Left            =   975
      TabIndex        =   3
      Top             =   270
      UseMnemonic     =   0   'False
      Width           =   3960
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "why are you seeing this dialog?"
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
      TabIndex        =   2
      Top             =   4320
      Width           =   3345
   End
End
Attribute VB_Name = "dialog_FirstRun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'First-Run Dialog (including language selection)
'Copyright ©2012-2013 by Tanner Helland
'Created: 12/February/12
'Last updated: 12/February/12
'Last update: initial build
'
'Dialog for presenting the user a choice of language.  At present, this dialog is only presented
' the first time the program is run.  (The one exception is if the requested language file goes
' missing - in this case, this dialog is presented again.)  After first run, the user can always
' change the language via the Tools -> Language menu.
'
'For convenience, an explanation is given to the user about why the form has been presented.  In
' most cases, this dialog will appear in English - unless a near-perfect match was found (language
' but not region); in that case, the closest language will be used for convenience.
'
'***************************************************************************

Option Explicit

'The user input from the dialog
Private userAnswer As VbMsgBoxResult

'A reference to the image being saved (actually, a temporary copy of the image being saved - but whatever).
Private srcLayer As pdLayer

'Our copy of the image being saved.  This will be created and destroyed frequently as the alpha values are updated.
Private tmpLayer As pdLayer

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
Private Sub cmdOK_Click()
        
    'Save the selected color depth to the corresponding global variable (so other functions can access it
    ' after this form is unloaded)
    g_AlphaCutoff = hsThreshold.Value
    
    'Free up memory
    tmpLayer.eraseLayer
    
    userAnswer = vbOK
    Me.Hide
    
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
    makeFormPretty Me
    
    'Display the dialog
    Me.Show vbModal, FormMain

End Sub

'Render a preview of the current alpha cut-off to the large picture box on the form
Private Sub renderPreview()

    tmpLayer.eraseLayer
    
    tmpLayer.createFromExistingLayer srcLayer
    tmpLayer.applyAlphaCutoff hsThreshold.Value, False
    
    DrawPreviewImage picPreview, True, tmpLayer

End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'Redraw the preview when the scroll bar is moved
Private Sub hsThreshold_Change()
    renderPreview
End Sub

Private Sub hsThreshold_Scroll()
    renderPreview
End Sub
