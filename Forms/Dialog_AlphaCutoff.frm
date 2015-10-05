VERSION 5.00
Begin VB.Form dialog_AlphaCutoff 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Please choose a transparency threshold"
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
      Height          =   915
      Left            =   4920
      TabIndex        =   1
      Top             =   3420
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   1614
      Caption         =   "background color for compositing"
   End
   Begin PhotoDemon.commandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   3
      Top             =   5910
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   1323
      BackColor       =   14802140
      dontAutoUnloadParent=   -1  'True
   End
   Begin PhotoDemon.sliderTextCombo sltThreshold 
      Height          =   720
      Left            =   4920
      TabIndex        =   0
      Top             =   2160
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   1270
      Caption         =   "transparency cut-off"
      Max             =   255
      Value           =   127
      NotchPosition   =   2
      NotchValueCustom=   127
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
   Begin PhotoDemon.pdLabel lblGuide 
      Height          =   240
      Index           =   1
      Left            =   7680
      Top             =   2940
      Width           =   2670
      _ExtentX        =   4710
      _ExtentY        =   503
      Alignment       =   1
      Caption         =   "maximum transparency "
      FontItalic      =   -1  'True
      FontSize        =   8
      ForeColor       =   4210752
   End
   Begin PhotoDemon.pdLabel lblGuide 
      Height          =   240
      Index           =   0
      Left            =   5040
      Top             =   2940
      Width           =   2535
      _ExtentX        =   0
      _ExtentY        =   503
      Caption         =   "no transparency "
      FontItalic      =   -1  'True
      FontSize        =   8
      ForeColor       =   4210752
   End
   Begin PhotoDemon.pdLabel lblWarning 
      Height          =   765
      Left            =   975
      Top             =   270
      Width           =   10440
      _ExtentX        =   0
      _ExtentY        =   503
      Caption         =   "This image has a complex alpha channel.  Before it can be saved as a paletted image (8bpp), the alpha channel must be simplified."
      ForeColor       =   2105376
      Layout          =   1
   End
End
Attribute VB_Name = "dialog_AlphaCutoff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Alpha Cut-Off Dialog
'Copyright 2012-2015 by Tanner Helland
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
    iconY = FixDPI(18)
    If g_UseFancyFonts Then iconY = iconY + FixDPI(2)
    DrawSystemIcon IDI_ASTERISK, Me.hDC, FixDPI(22), iconY
        
    'Initialize our temporary DIB render object
    Set tmpDIB = New pdDIB
        
    'Render a preview of this threshold value
    updatePreview
        
    Message "Waiting for user to specify alpha threshold... "
        
    'Apply translations and visual themes
    MakeFormPretty Me
    
    'Display the dialog
    ShowPDDialog vbModal, Me, True

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

