VERSION 5.00
Begin VB.Form toolbar_File 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "File"
   ClientHeight    =   8715
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   1050
   ClipControls    =   0   'False
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
   ScaleHeight     =   581
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   70
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin PhotoDemon.jcbutton cmdOpen 
      Height          =   615
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   1085
      ButtonStyle     =   13
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   15199212
      Caption         =   ""
      HandPointer     =   -1  'True
      PictureNormal   =   "VBP_ToolbarPrimary.frx":0000
      DisabledPictureMode=   1
      CaptionEffects  =   0
      TooltipTitle    =   "Open"
   End
   Begin PhotoDemon.jcbutton cmdSave 
      Height          =   615
      Left            =   60
      TabIndex        =   1
      Top             =   1440
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   1085
      ButtonStyle     =   13
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   15199212
      Caption         =   ""
      HandPointer     =   -1  'True
      PictureNormal   =   "VBP_ToolbarPrimary.frx":1452
      DisabledPictureMode=   1
      CaptionEffects  =   0
      TooltipTitle    =   "Save"
   End
   Begin PhotoDemon.jcbutton cmdUndo 
      Height          =   615
      Left            =   60
      TabIndex        =   2
      Top             =   2820
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   1085
      ButtonStyle     =   13
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   15199212
      Caption         =   ""
      HandPointer     =   -1  'True
      PictureNormal   =   "VBP_ToolbarPrimary.frx":26B4
      DisabledPictureMode=   1
      CaptionEffects  =   0
      TooltipTitle    =   "Undo"
      TooltipBackColor=   -2147483643
   End
   Begin PhotoDemon.jcbutton cmdRedo 
      Height          =   615
      Left            =   60
      TabIndex        =   3
      Top             =   3450
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   1085
      ButtonStyle     =   13
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   15199212
      Caption         =   ""
      HandPointer     =   -1  'True
      PictureNormal   =   "VBP_ToolbarPrimary.frx":3706
      DisabledPictureMode=   1
      CaptionEffects  =   0
      TooltipTitle    =   "Redo"
      TooltipBackColor=   -2147483643
   End
   Begin PhotoDemon.jcbutton cmdClose 
      Height          =   615
      Left            =   60
      TabIndex        =   4
      Top             =   690
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   1085
      ButtonStyle     =   13
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   15199212
      Caption         =   ""
      HandPointer     =   -1  'True
      PictureNormal   =   "VBP_ToolbarPrimary.frx":4758
      DisabledPictureMode=   1
      CaptionEffects  =   0
      TooltipTitle    =   "Close"
   End
   Begin PhotoDemon.jcbutton cmdSaveAs 
      Height          =   615
      Left            =   60
      TabIndex        =   5
      Top             =   2070
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   1085
      ButtonStyle     =   13
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   15199212
      Caption         =   ""
      HandPointer     =   -1  'True
      PictureNormal   =   "VBP_ToolbarPrimary.frx":57AA
      DisabledPictureMode=   1
      CaptionEffects  =   0
      TooltipTitle    =   "Save As"
   End
   Begin VB.Label lblRecording 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "macro recording in progress..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   2160
      Left            =   30
      TabIndex        =   6
      Top             =   4200
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   960
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "toolbar_File"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Primary Toolbar
'Copyright ©2013-2014 by Tanner Helland
'Created: 02/October/13
'Last updated: 03/October/13
'Last update: minor bug-fixes
'
'This form was initially integrated into the main MDI form.  In fall 2014, PhotoDemon left behind the MDI model,
' and all toolbars were moved to their own forms.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Custom tooltip class allows for things like multiline, theming, and multiple monitor support
Dim m_ToolTip As clsToolTip

Private Sub cmdClose_Click()
    Process "Close", True
End Sub

Private Sub cmdOpen_Click()
    Process "Open", True
End Sub

Private Sub cmdRedo_Click()
    Process "Redo", , , UNDO_NOTHING
End Sub

Private Sub cmdSave_Click()
    Process "Save", , , UNDO_NOTHING
End Sub

Private Sub cmdSaveAs_Click()
    Process "Save as", True, , UNDO_NOTHING
End Sub

Private Sub cmdUndo_Click()
    Process "Undo", , , UNDO_NOTHING
End Sub

Private Sub Form_Load()
    
    'Assign the system hand cursor to all relevant objects
    Set m_ToolTip = New clsToolTip
    makeFormPretty Me, m_ToolTip
    
End Sub

'Toolbars can never be unloaded, EXCEPT when the whole program is going down.  Check for the program-wide closing flag prior
' to exiting; if it is not found, cancel the unload and simply hide this form.  (Note that the toggleToolbarVisibility sub
' will also keep this toolbar's Window menu entry in sync with the form's current visibility.)
Private Sub Form_Unload(Cancel As Integer)
    
    If g_ProgramShuttingDown Then
        ReleaseFormTheming Me
        g_WindowManager.unregisterForm Me
    Else
        Cancel = True
        toggleToolbarVisibility FILE_TOOLBOX
    End If
    
End Sub

'External functions can use this to re-theme this form at run-time (important when changing languages, for example)
Public Sub requestMakeFormPretty()
    makeFormPretty Me, m_ToolTip
End Sub
