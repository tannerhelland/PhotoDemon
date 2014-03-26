VERSION 5.00
Begin VB.Form toolbar_Layers 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Layers"
   ClientHeight    =   7245
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3735
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
   ScaleHeight     =   483
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   249
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picLayerButtons 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   249
      TabIndex        =   5
      Top             =   6750
      Width           =   3735
      Begin PhotoDemon.jcbutton cmdLayerAction 
         Height          =   480
         Index           =   0
         Left            =   960
         TabIndex        =   6
         Top             =   15
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   847
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
         PictureNormal   =   "VBP_ToolbarLayers.frx":0000
         DisabledPictureMode=   1
         CaptionEffects  =   0
         TooltipTitle    =   "Open"
      End
      Begin PhotoDemon.jcbutton cmdLayerAction 
         Height          =   480
         Index           =   1
         Left            =   1560
         TabIndex        =   7
         Top             =   15
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   847
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
         PictureNormal   =   "VBP_ToolbarLayers.frx":0D52
         DisabledPictureMode=   1
         CaptionEffects  =   0
         TooltipTitle    =   "Open"
      End
      Begin PhotoDemon.jcbutton cmdLayerAction 
         Height          =   480
         Index           =   2
         Left            =   2160
         TabIndex        =   8
         Top             =   15
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   847
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
         PictureNormal   =   "VBP_ToolbarLayers.frx":1AA4
         DisabledPictureMode=   1
         CaptionEffects  =   0
         TooltipTitle    =   "Open"
      End
   End
   Begin VB.PictureBox picLayers 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4935
      Left            =   120
      ScaleHeight     =   327
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   231
      TabIndex        =   4
      Top             =   1200
      Width           =   3495
   End
   Begin VB.TextBox txtName 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1080
      TabIndex        =   1
      Text            =   "Layer Name"
      Top             =   105
      Width           =   2535
   End
   Begin PhotoDemon.sliderTextCombo sltSelectionFeathering 
      CausesValidation=   0   'False
      Height          =   495
      Left            =   960
      TabIndex        =   3
      Top             =   480
      Width           =   2760
      _ExtentX        =   4868
      _ExtentY        =   873
      Max             =   100
      Value           =   100
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
   Begin VB.Line lnSeparator 
      BorderColor     =   &H8000000D&
      Index           =   0
      X1              =   8
      X2              =   240
      Y1              =   72
      Y2              =   72
   End
   Begin VB.Label lblLayerSettings 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "opacity:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00606060&
      Height          =   240
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   570
      Width           =   675
   End
   Begin VB.Label lblLayerSettings 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "name:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00606060&
      Height          =   240
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   555
   End
End
Attribute VB_Name = "toolbar_Layers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Custom tooltip class allows for things like multiline, theming, and multiple monitor support
Dim m_ToolTip As clsToolTip

Private Sub Form_Load()

    'Assign the system hand cursor to all relevant objects
    Set m_ToolTip = New clsToolTip
    makeFormPretty Me, m_ToolTip

End Sub

Private Sub Form_Resize()

    'When the parent form is resized, resize the layer list (and other items) to properly fill the
    ' available vertical space.
    
    'Start by moving the button box to the bottom of the available area
    picLayerButtons.Top = Me.ScaleHeight - picLayerButtons.Height - fixDPI(8)
    
    'Next, stretch the layer box to fill the available space
    picLayers.Height = (picLayerButtons.Top - picLayers.Top) - fixDPI(8)

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
        toggleToolbarVisibility TOOLS_TOOLBOX
    End If
    
End Sub
