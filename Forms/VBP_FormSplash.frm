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
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version (populated at run-time)"
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
      Caption         =   "Initializing software..."
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
'Copyright ©2000-2012 by Tanner Helland
'Created: 15/April/01
'Last updated: 10/June/12
'Last update: no more "force to top" z-order code.  It was never necessary in
'             the first place, and it's bad UI behavior (especially if the program
'             has to throw a message box for some reason, because it gets hidden
'             behind the forced-to-top form).
'
'Responsible for checking the runtime environment and building paths
'accordingly.  Also shows a nice little loading message while it does its thing.
'
'***************************************************************************

Option Explicit

'We use this to ensure that the splash shows for at least 1 second
Private Const LOADTIME As Single = 1#

'The form is loaded invisibly, so this code is placed in the _Activate event instead of the more common _Load event
Private Sub Form_Activate()

    'Check to see if we're running in the IDE or as a compiled EXE (see below)
    CheckEnvironment
    
    'Make sure the splash screen shows for at least 1 second
    Me.Show
    Dim OT As Single
    OT = Timer
    Do While Timer - OT < LOADTIME
        DoEvents
    Loop
    
End Sub

'Check for IDE or compiled EXE, and set program parameters accordingly
Private Sub CheckEnvironment()
    
    'Check the run-time environment.
    
    'App is compiled:
    If App.LogMode = 1 Then
        
        IsProgramCompiled = True
        
        'Determine the version automatically from the EXE information
        lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
        
        'Disable the "Test" menu (that I use for debugging)
        FormMain.MnuTest.Visible = False
        
    'App is not compiled:
    Else
    
        IsProgramCompiled = False

        'Add a gentle reminder to compile the program
        lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision & " - please compile!"
        
    End If
    
End Sub
