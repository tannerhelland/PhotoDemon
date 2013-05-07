VERSION 5.00
Begin VB.Form FormSplash 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2550
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   9000
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "VBP_FormSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   170
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   600
   ShowInTaskbar   =   0   'False
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   2280
      Width           =   2925
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblMessage 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Live updates..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   2880
      TabIndex        =   0
      Top             =   2280
      Width           =   6045
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "FormSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Simple Splash Screen
'Copyright ©2000-2013 by Tanner Helland
'Created: 15/April/01
'Last updated: 03/January/13
'Last update: move code from the LoadProgram routine to here (to help tidy up that function)
'
'Responsible for checking the runtime environment and building paths
' accordingly.  Also shows a nice little loading message while it does its thing.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://www.tannerhelland.com/photodemon/#license
'
'***************************************************************************

Option Explicit

'We use these to ensure that the splash shows for at least 1 second
Private Const LOADTIME As Double = 1#
Dim OT As Double

'Custom tooltip class allows for things like multiline, theming, and multiple monitor support
Dim m_ToolTip As clsToolTip

'The form is loaded invisibly, so this code is placed in the _Activate event instead of the more common _Load event
Private Sub Form_Activate()
    
    'We want to make sure the splash screen shows for at least 1 second, so make a note of the time when the form is loaded
    OT = Timer
        
End Sub

Public Sub prepareSplash()

    'Before we can display the splash screen, we need to paint the program logo to it.  (This is done dynamically
    ' for several reasons; it allows us to keep just one copy of the logo in the project, and it guarantees proper
    ' painting regardless of screen DPI.)
    Dim logoWidth As Long, logoHeight As Long
    Dim logoAspectRatio As Double
    
    logoWidth = FormMain.picLogo.ScaleWidth
    logoHeight = FormMain.picLogo.ScaleHeight
    logoAspectRatio = CDbl(logoWidth) / CDbl(logoHeight)
    
    SetStretchBltMode Me.hDC, STRETCHBLT_HALFTONE
    StretchBlt Me.hDC, 0, 0, Me.ScaleWidth, Me.ScaleWidth / logoAspectRatio, FormMain.picLogo.hDC, 0, 0, logoWidth, logoHeight, vbSrcCopy
    Me.Picture = Me.Image
    
End Sub

'When the form is unloaded, pause until at least a second has passed
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Timer - OT < LOADTIME Then
        Do While Timer - OT < LOADTIME
            DoEvents
        Loop
    End If

End Sub
