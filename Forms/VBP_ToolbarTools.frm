VERSION 5.00
Begin VB.Form toolbar_Tools 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Tools"
   ClientHeight    =   2355
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   13665
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
   ScaleHeight     =   157
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   911
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picTools 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1575
      Index           =   1
      Left            =   15
      ScaleHeight     =   105
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   918
      TabIndex        =   25
      Top             =   1020
      Visible         =   0   'False
      Width           =   13770
      Begin PhotoDemon.smartCheckBox chkLayerBorder 
         Height          =   480
         Left            =   6480
         TabIndex        =   27
         Top             =   360
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   847
         Caption         =   "show layer borders"
         Value           =   1
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
      Begin PhotoDemon.smartCheckBox chkLayerNodes 
         Height          =   480
         Left            =   6480
         TabIndex        =   28
         Top             =   780
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   847
         Caption         =   "show layer transform nodes"
         Value           =   1
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
      Begin PhotoDemon.smartCheckBox chkAutoActivateLayer 
         Height          =   480
         Left            =   120
         TabIndex        =   30
         Top             =   360
         Width           =   4080
         _ExtentX        =   7197
         _ExtentY        =   847
         Caption         =   "automatically activate layer beneath mouse"
         Value           =   1
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
      Begin PhotoDemon.smartCheckBox chkIgnoreTransparent 
         Height          =   480
         Left            =   120
         TabIndex        =   31
         Top             =   780
         Width           =   4920
         _ExtentX        =   8678
         _ExtentY        =   847
         Caption         =   "ignore transparent pixels when auto-activating layers"
         Value           =   1
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
      Begin VB.Label lblOptions 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "interaction options:"
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
         Left            =   120
         TabIndex        =   29
         Top             =   60
         Width           =   1650
      End
      Begin VB.Label lblOptions 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "display options:"
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
         Left            =   6480
         TabIndex        =   26
         Top             =   60
         Width           =   1335
      End
   End
   Begin PhotoDemon.jcbutton cmdTools 
      Height          =   600
      Index           =   2
      Left            =   1950
      TabIndex        =   16
      Top             =   315
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1058
      ButtonStyle     =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483643
      Caption         =   ""
      Mode            =   1
      HandPointer     =   -1  'True
      PictureNormal   =   "VBP_ToolbarTools.frx":0000
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      ColorScheme     =   3
   End
   Begin VB.PictureBox picTools 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1575
      Index           =   0
      Left            =   15
      ScaleHeight     =   105
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   918
      TabIndex        =   0
      Top             =   1020
      Visible         =   0   'False
      Width           =   13770
      Begin VB.ComboBox cmbSelRender 
         Appearance      =   0  'Flat
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         ItemData        =   "VBP_ToolbarTools.frx":0BE2
         Left            =   120
         List            =   "VBP_ToolbarTools.frx":0BE4
         Style           =   2  'Dropdown List
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   390
         Width           =   2250
      End
      Begin VB.ComboBox cmbSelType 
         Appearance      =   0  'Flat
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         ItemData        =   "VBP_ToolbarTools.frx":0BE6
         Left            =   8340
         List            =   "VBP_ToolbarTools.frx":0BE8
         Style           =   2  'Dropdown List
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "This option controls the selection's area.  You can switch between the three settings without losing the current selection."
         Top             =   390
         Width           =   2445
      End
      Begin VB.ComboBox cmbSelSmoothing 
         Appearance      =   0  'Flat
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         ItemData        =   "VBP_ToolbarTools.frx":0BEA
         Left            =   5640
         List            =   "VBP_ToolbarTools.frx":0BEC
         Style           =   2  'Dropdown List
         TabIndex        =   1
         TabStop         =   0   'False
         ToolTipText     =   "Use this option to change the way selections blend with their surroundings."
         Top             =   390
         Width           =   2445
      End
      Begin PhotoDemon.sliderTextCombo sltCornerRounding 
         CausesValidation=   0   'False
         Height          =   495
         Left            =   10860
         TabIndex        =   3
         Top             =   345
         Width           =   2670
         _ExtentX        =   4710
         _ExtentY        =   873
         Max             =   10000
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
      Begin PhotoDemon.textUpDown tudSel 
         Height          =   405
         Index           =   0
         Left            =   2520
         TabIndex        =   4
         Top             =   390
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   714
         Min             =   -30000
         Max             =   30000
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
      Begin PhotoDemon.textUpDown tudSel 
         Height          =   405
         Index           =   1
         Left            =   2520
         TabIndex        =   5
         Top             =   840
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   714
         Min             =   -30000
         Max             =   30000
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
      Begin PhotoDemon.textUpDown tudSel 
         Height          =   405
         Index           =   2
         Left            =   4080
         TabIndex        =   6
         Top             =   390
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   714
         Min             =   -30000
         Max             =   30000
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
      Begin PhotoDemon.textUpDown tudSel 
         Height          =   405
         Index           =   3
         Left            =   4080
         TabIndex        =   7
         Top             =   840
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   714
         Min             =   -30000
         Max             =   30000
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
      Begin PhotoDemon.sliderTextCombo sltSelectionBorder 
         CausesValidation=   0   'False
         Height          =   495
         Left            =   8220
         TabIndex        =   8
         Top             =   840
         Visible         =   0   'False
         Width           =   2670
         _ExtentX        =   4710
         _ExtentY        =   873
         Min             =   1
         Max             =   10000
         Value           =   1
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
      Begin PhotoDemon.sliderTextCombo sltSelectionFeathering 
         CausesValidation=   0   'False
         Height          =   495
         Left            =   5520
         TabIndex        =   9
         Top             =   840
         Width           =   2670
         _ExtentX        =   4710
         _ExtentY        =   873
         Max             =   100
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
      Begin PhotoDemon.sliderTextCombo sltSelectionLineWidth 
         CausesValidation=   0   'False
         Height          =   495
         Left            =   10860
         TabIndex        =   10
         Top             =   345
         Width           =   2670
         _ExtentX        =   4710
         _ExtentY        =   873
         Min             =   1
         Max             =   10000
         Value           =   10
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
      Begin VB.Label lblSelection 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "appearance:"
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
         Left            =   120
         TabIndex        =   19
         Top             =   60
         Width           =   1080
      End
      Begin VB.Label lblSelection 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "size (w, h)"
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
         Index           =   2
         Left            =   4080
         TabIndex        =   15
         Top             =   60
         Width           =   915
      End
      Begin VB.Label lblSelection 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "position (x, y)"
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
         Left            =   2520
         TabIndex        =   14
         Top             =   60
         Width           =   1170
      End
      Begin VB.Label lblSelection 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "corner rounding"
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
         Index           =   5
         Left            =   10980
         TabIndex        =   13
         Top             =   60
         Width           =   1365
      End
      Begin VB.Label lblSelection 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "area"
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
         Index           =   4
         Left            =   8340
         TabIndex        =   12
         Top             =   60
         Width           =   390
      End
      Begin VB.Label lblSelection 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "smoothing"
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
         Index           =   3
         Left            =   5640
         TabIndex        =   11
         Top             =   60
         Width           =   885
      End
   End
   Begin PhotoDemon.jcbutton cmdTools 
      Height          =   600
      Index           =   3
      Left            =   2670
      TabIndex        =   17
      Top             =   315
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1058
      ButtonStyle     =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483643
      Caption         =   ""
      Mode            =   1
      HandPointer     =   -1  'True
      PictureNormal   =   "VBP_ToolbarTools.frx":0BEE
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      ColorScheme     =   3
   End
   Begin PhotoDemon.jcbutton cmdTools 
      Height          =   600
      Index           =   4
      Left            =   3390
      TabIndex        =   18
      Top             =   315
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1058
      ButtonStyle     =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483643
      Caption         =   ""
      Mode            =   1
      HandPointer     =   -1  'True
      PictureNormal   =   "VBP_ToolbarTools.frx":17D0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      ColorScheme     =   3
   End
   Begin PhotoDemon.jcbutton cmdTools 
      Height          =   600
      Index           =   0
      Left            =   120
      TabIndex        =   23
      Top             =   315
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1058
      ButtonStyle     =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483643
      Caption         =   ""
      Mode            =   1
      HandPointer     =   -1  'True
      PictureNormal   =   "VBP_ToolbarTools.frx":23B2
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      ColorScheme     =   3
   End
   Begin PhotoDemon.jcbutton cmdTools 
      Height          =   600
      Index           =   1
      Left            =   840
      TabIndex        =   24
      Top             =   315
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1058
      ButtonStyle     =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483643
      Caption         =   ""
      Mode            =   1
      HandPointer     =   -1  'True
      PictureNormal   =   "VBP_ToolbarTools.frx":3104
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      ColorScheme     =   3
   End
   Begin VB.Label lblCategory 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "nav"
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
      Left            =   120
      TabIndex        =   22
      Top             =   30
      Width           =   300
   End
   Begin VB.Line lineMain 
      BorderColor     =   &H80000002&
      Index           =   2
      X1              =   0
      X2              =   5000
      Y1              =   67
      Y2              =   67
   End
   Begin VB.Label lblCategory 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "selections"
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
      Left            =   1920
      TabIndex        =   21
      Top             =   30
      Width           =   840
   End
   Begin VB.Line lineMain 
      BorderColor     =   &H80000002&
      Index           =   1
      X1              =   0
      X2              =   5000
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line lineMain 
      BorderColor     =   &H80000002&
      Index           =   0
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   2000
   End
End
Attribute VB_Name = "toolbar_Tools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Tools Toolbox
'Copyright ©2013-2014 by Tanner Helland
'Created: 03/October/13
'Last updated: 02/May/14
'Last update: started adding additional options for the "move/size" tool.
'
'This form was initially integrated into the main MDI form.  In fall 2013, PhotoDemon left behind the MDI model,
' and all toolbars were moved to their own forms.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Used to toggle the command button state of the toolbar buttons
Private Const BM_SETSTATE = &HF3
Private Declare Function SendMessageA Lib "user32" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

'Custom tooltip class allows for things like multiline, theming, and multiple monitor support
Dim m_ToolTip As clsToolTip

'The value of all controls on this form are saved and loaded to file by this class
Private WithEvents lastUsedSettings As pdLastUsedSettings
Attribute lastUsedSettings.VB_VarHelpID = -1

Private Sub chkAutoActivateLayer_Click()
    If CBool(chkAutoActivateLayer) Then
        If Not chkIgnoreTransparent.Enabled Then chkIgnoreTransparent.Enabled = True
    Else
        If chkIgnoreTransparent.Enabled Then chkIgnoreTransparent.Enabled = False
    End If
End Sub

'Show/hide layer borders while using the move tool
Private Sub chkLayerBorder_Click()
    PrepareViewport pdImages(g_CurrentImage), FormMain.mainCanvas(0), "Layer border toggle"
End Sub

'Show/hide layer transform nodes while using the move tool
Private Sub chkLayerNodes_Click()
    PrepareViewport pdImages(g_CurrentImage), FormMain.mainCanvas(0), "Layer nodes toggle"
End Sub

Private Sub cmdTools_Click(Index As Integer)
    g_PreviousTool = g_CurrentTool
    g_CurrentTool = Index
    resetToolButtonStates
End Sub

Private Sub Form_Load()

    Dim i As Long
    
    'Because line controls aren't automatically made DPI-aware by VB, we must manually move this dialog's line
    ' control into place.
    lineMain(2).y1 = picTools(0).Top - fixDPI(2)
    lineMain(2).y2 = lineMain(2).y1
    
    'INITIALIZE ALL TOOLS
    
        'Tool button tooltips
        cmdTools(NAV_DRAG).ToolTip = g_Language.TranslateMessage("Hand (click-and-drag image scrolling)")
        cmdTools(NAV_DRAG).ToolTip = g_Language.TranslateMessage("Hand (click-and-drag image scrolling)")
        cmdTools(SELECT_RECT).ToolTip = g_Language.TranslateMessage("Rectangular Selection")
        cmdTools(SELECT_CIRC).ToolTip = g_Language.TranslateMessage("Elliptical (Oval) Selection")
        cmdTools(SELECT_LINE).ToolTip = g_Language.TranslateMessage("Line Selection")
    
        'Selection visual styles (currently lightbox or highlight)
        toolbar_Tools.cmbSelRender(0).ToolTipText = g_Language.TranslateMessage("Click to change the way selections are rendered onto the image canvas.  This has no bearing on selection contents - only the way they appear while editing.")
        For i = 0 To toolbar_Tools.cmbSelRender.Count - 1
            toolbar_Tools.cmbSelRender(i).AddItem "Lightbox", 0
            toolbar_Tools.cmbSelRender(i).AddItem "Highlight (Blue)", 1
            toolbar_Tools.cmbSelRender(i).AddItem "Highlight (Red)", 2
            toolbar_Tools.cmbSelRender(i).ListIndex = 0
        Next i
        
        'Selection smoothing (currently none, antialiased, fully feathered)
        toolbar_Tools.cmbSelSmoothing(0).ToolTipText = g_Language.TranslateMessage("This option controls how smoothly a selection blends with its surroundings.")
        toolbar_Tools.cmbSelSmoothing(0).AddItem "None", 0
        toolbar_Tools.cmbSelSmoothing(0).AddItem "Antialiased", 1
        
        'Previously, live feathering was disallowed on XP or Vista for performance reasons (GDI+ can't be used to blur
        ' the selection mask, and our own code was too slow).  As of 17 Oct '13, I have reinstated live selection
        ' feathering on these OSes using PD's very fast horizontal and vertical blur.  While not perfect, this should
        ' still provide "good enough" performance for smaller images and/or slight feathering.
        toolbar_Tools.cmbSelSmoothing(0).AddItem "Feathered", 2
        toolbar_Tools.cmbSelSmoothing(0).ListIndex = 1
        
        'Selection types (currently interior, exterior, border)
        toolbar_Tools.cmbSelType(0).ToolTipText = g_Language.TranslateMessage("These options control the area affected by a selection.  The selection can be modified on-canvas while any of these settings are active.  For more advanced selection adjustments, use the Select menu.")
        toolbar_Tools.cmbSelType(0).AddItem "Interior", 0
        toolbar_Tools.cmbSelType(0).AddItem "Exterior", 1
        toolbar_Tools.cmbSelType(0).AddItem "Border", 2
        toolbar_Tools.cmbSelType(0).ListIndex = 0
        
        toolbar_Tools.sltSelectionFeathering.assignTooltip "This feathering slider allows for immediate feathering adjustments.  For performance reasons, it is limited to small radii.  For larger feathering radii, please use the Select -> Feathering menu."
        toolbar_Tools.sltCornerRounding.assignTooltip "This option adjusts the roundness of a rectangular selection's corners."
        toolbar_Tools.sltSelectionLineWidth.assignTooltip "This option adjusts the width of a line selection."
        toolbar_Tools.sltSelectionBorder.assignTooltip "This option adjusts the width of the selection border."
        
        'Load any last-used settings for this form
        Set lastUsedSettings = New pdLastUsedSettings
        lastUsedSettings.setParentForm Me
        lastUsedSettings.loadAllControlValues
        
        'Assign the system hand cursor to all relevant objects
        Set m_ToolTip = New clsToolTip
        makeFormPretty Me, m_ToolTip

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    'Save all last-used settings to file
    lastUsedSettings.saveAllControlValues

End Sub

Private Sub lastUsedSettings_AddCustomPresetData()
    
    'Write the currently selected selection tool to file
    lastUsedSettings.addPresetData "ActiveSelectionTool", g_CurrentTool
    
End Sub

Private Sub lastUsedSettings_ReadCustomPresetData()

    'Restore the last-used selection tool (which will be saved in the main form's preset file, if it exists)
    g_PreviousTool = -1
    If Len(lastUsedSettings.retrievePresetData("ActiveSelectionTool")) > 0 Then
        g_CurrentTool = CLng(lastUsedSettings.retrievePresetData("ActiveSelectionTool"))
    Else
        g_CurrentTool = NAV_DRAG
    End If
    resetToolButtonStates
    
    'Reset the selection coordinate boxes to 0
    Dim i As Long
    For i = 0 To tudSel.Count - 1
        tudSel(i) = 0
    Next i

End Sub

'When the selection type is changed, update the corresponding preference and redraw all selections
Private Sub cmbSelRender_Click(Index As Integer)
            
    If g_OpenImageCount > 0 Then
    
        Dim i As Long
        For i = 0 To g_NumOfImagesLoaded
            If (Not pdImages(i) Is Nothing) Then
                If pdImages(i).IsActive And pdImages(i).selectionActive Then RenderViewport pdImages(i), FormMain.mainCanvas(0)
            End If
        Next i
    
    End If
    
End Sub

'Change selection smoothing (e.g. none, antialiased, fully feathered)
Private Sub cmbSelSmoothing_Click(Index As Integer)
    
    updateSelectionPanelLayout
    
    'If a selection is already active, change its type to match the current selection, then redraw it
    If selectionsAllowed(False) Then
        pdImages(g_CurrentImage).mainSelection.setSmoothingType cmbSelSmoothing(Index).ListIndex
        pdImages(g_CurrentImage).mainSelection.setFeatheringRadius sltSelectionFeathering.Value
        RenderViewport pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    End If
    
End Sub

'Change selection type (e.g. interior, exterior, bordered)
Private Sub cmbSelType_Click(Index As Integer)

    updateSelectionPanelLayout
    
    'If a selection is already active, change its type to match the current selection, then redraw it
    If selectionsAllowed(False) Then
        pdImages(g_CurrentImage).mainSelection.setSelectionType cmbSelType(Index).ListIndex
        pdImages(g_CurrentImage).mainSelection.setBorderSize sltSelectionBorder.Value
        RenderViewport pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    End If
    
End Sub

'External functions can use this to request the selection of a new tool (for example, Select All uses this to set the
' rectangular tool selector as the current tool)
Public Sub selectNewTool(ByVal newToolID As Long)
    g_PreviousTool = g_CurrentTool
    g_CurrentTool = newToolID
    resetToolButtonStates
End Sub

'When a new tool button is selected, we need to raise all the others and display the proper options box
Public Sub resetToolButtonStates()
    
    'Start by depressing the selected button and raising all unselected ones
    Dim catID As Long
    For catID = 0 To cmdTools.Count - 1
        If catID = g_CurrentTool Then
            cmdTools(catID).Value = True
        Else
            cmdTools(catID).Value = False
        End If
    Next catID
    
    Dim i As Long
    
    'Next, we need to display the correct tool options panel.  There is no set pattern to this; some tools share
    ' panels, but show/hide certain controls as necessary.  Other tools require their own unique panel.  I've tried
    ' to strike a balance between "as few panels as possible" without going overboard.
    Dim activeToolPanel As Long
    
    Select Case g_CurrentTool
        
        'Move/size tool
        Case NAV_MOVE
            activeToolPanel = 1
        
        'Rectangular, Elliptical, Line selections
        Case SELECT_RECT, SELECT_CIRC, SELECT_LINE
            activeToolPanel = 0
        
        Case Else
            activeToolPanel = -1
        
    End Select
    
    'If tools share the same panel, they may need to show or hide a few additional controls.  (For example,
    ' "corner rounding", which is needed for rectangular selections but not elliptical ones, despite the two
    ' sharing the same tool panel.)  Do this before showing or hiding the tool panel.
    Select Case g_CurrentTool
    
        'For rectangular selections, show the rounded corners option
        Case SELECT_RECT
            toolbar_Tools.lblSelection(5).Visible = True
            toolbar_Tools.sltCornerRounding.Visible = True
            toolbar_Tools.sltSelectionLineWidth.Visible = False
            
        'For elliptical selections, hide the rounded corners option
        Case SELECT_CIRC
            toolbar_Tools.lblSelection(5).Visible = False
            toolbar_Tools.sltCornerRounding.Visible = False
            toolbar_Tools.sltSelectionLineWidth.Visible = False
            
        'Line selections also show the rounded corners slider, though they repurpose it for line width
        Case SELECT_LINE
            toolbar_Tools.lblSelection(5).Visible = True
            toolbar_Tools.sltCornerRounding.Visible = False
            toolbar_Tools.sltSelectionLineWidth.Visible = True
        
    End Select
    
    'Even if tools share the same panel, they may name controls differently, or use different max/min values.
    ' Check for this, and apply new text and max/min settings as necessary.
    Select Case g_CurrentTool
    
        'Rectangular and elliptical selections use rectangular bounding boxes and potential corner rounding
        Case SELECT_RECT, SELECT_CIRC
            lblSelection(1).Caption = g_Language.TranslateMessage("position (x, y)")
            lblSelection(2).Caption = g_Language.TranslateMessage("size (x, y)")
            lblSelection(5).Caption = g_Language.TranslateMessage("corner rounding")
            
        'Line selections use two points, and the corner rounding slider gets repurposed as line width.
        Case SELECT_LINE
            lblSelection(1).Caption = g_Language.TranslateMessage("1st point (x, y)")
            lblSelection(2).Caption = g_Language.TranslateMessage("2nd point (x, y)")
            lblSelection(5).Caption = g_Language.TranslateMessage("line width")
            
    End Select
    
    'Display the current tool options panel, while hiding all inactive ones
    For i = 0 To picTools.Count - 1
        If i = activeToolPanel Then
            If Not picTools(i).Visible Then
                picTools(i).Visible = True
                setArrowCursor picTools(i)
            End If
        Else
            If picTools(i).Visible Then picTools(i).Visible = False
        End If
    Next i
    
    newToolSelected
        
End Sub

'When a new tool is selected, we may need to initialize certain values
Private Sub newToolSelected()
    
    Select Case g_CurrentTool
    
        'Rectangular, elliptical selections
        Case SELECT_RECT
                
            'If a similar selection is already active, change its shape to match the current tool, then redraw it
            If selectionsAllowed(True) And (Not g_UndoRedoActive) Then
                If (g_PreviousTool = SELECT_CIRC) And (pdImages(g_CurrentImage).mainSelection.getSelectionShape = sCircle) Then
                    pdImages(g_CurrentImage).mainSelection.setSelectionShape sRectangle
                    RenderViewport pdImages(g_CurrentImage), FormMain.mainCanvas(0)
                Else
                    
                    If pdImages(g_CurrentImage).mainSelection.getSelectionShape = sRectangle Then
                        metaToggle tSelectionTransform, True
                    Else
                    
                        'Remove any existing selections
                        If g_OpenImageCount > 0 Then Process "Remove selection", , pdImages(g_CurrentImage).mainSelection.getSelectionParamString, UNDO_SELECTION
                    
                        metaToggle tSelectionTransform, False
                        
                    End If
                End If
            End If
            
        Case SELECT_CIRC
        
            'If a similar selection is already active, change its shape to match the current tool, then redraw it
            If selectionsAllowed(True) And (Not g_UndoRedoActive) Then
                If (g_PreviousTool = SELECT_RECT) And (pdImages(g_CurrentImage).mainSelection.getSelectionShape = sRectangle) Then
                    pdImages(g_CurrentImage).mainSelection.setSelectionShape sCircle
                    RenderViewport pdImages(g_CurrentImage), FormMain.mainCanvas(0)
                Else
                    
                    If pdImages(g_CurrentImage).mainSelection.getSelectionShape = sCircle Then
                        metaToggle tSelectionTransform, True
                    Else
                        
                        'Remove any existing selections
                        If g_OpenImageCount > 0 Then Process "Remove selection", , pdImages(g_CurrentImage).mainSelection.getSelectionParamString, UNDO_SELECTION
                        
                        metaToggle tSelectionTransform, False
                        
                    End If
                End If
            End If
            
        'Line selections
        Case SELECT_LINE
        
            'Deactivate the position text boxes - those shouldn't be accessible unless a line selection is presently active
            If selectionsAllowed(True) Then
                If pdImages(g_CurrentImage).mainSelection.getSelectionShape = sLine Then
                    metaToggle tSelectionTransform, True
                Else
                
                    'Remove any existing selections
                    If g_OpenImageCount > 0 Then Process "Remove selection", , pdImages(g_CurrentImage).mainSelection.getSelectionParamString, UNDO_SELECTION
                
                    metaToggle tSelectionTransform, False
                    
                End If
            Else
                metaToggle tSelectionTransform, False
            End If
            
        Case Else
        
    End Select
    
    'Finally, because tools may do some custom rendering atop the image canvas, now is a good time to redraw the canvas
    RenderViewport pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    
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

Private Sub sltCornerRounding_Change()
    If selectionsAllowed(True) Then
        pdImages(g_CurrentImage).mainSelection.setRoundedCornerAmount sltCornerRounding.Value
        RenderViewport pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    End If
End Sub

Private Sub sltSelectionBorder_Change()
    If selectionsAllowed(False) Then
        pdImages(g_CurrentImage).mainSelection.setBorderSize sltSelectionBorder.Value
        RenderViewport pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    End If
End Sub

Private Sub sltSelectionFeathering_Change()
    If selectionsAllowed(False) Then
        pdImages(g_CurrentImage).mainSelection.setFeatheringRadius sltSelectionFeathering.Value
        RenderViewport pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    End If
End Sub

Private Sub sltSelectionLineWidth_Change()
    If selectionsAllowed(True) Then
        pdImages(g_CurrentImage).mainSelection.setSelectionLineWidth sltSelectionLineWidth.Value
        RenderViewport pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    End If
End Sub

Private Function selectionsAllowed(ByVal transformableMatters As Boolean) As Boolean
    If g_OpenImageCount > 0 Then
        If pdImages(g_CurrentImage).selectionActive And (Not pdImages(g_CurrentImage).mainSelection Is Nothing) And (Not pdImages(g_CurrentImage).mainSelection.rejectRefreshRequests) Then
            
            If transformableMatters Then
                If pdImages(g_CurrentImage).mainSelection.isTransformable Then
                    selectionsAllowed = True
                Else
                    selectionsAllowed = False
                End If
            Else
                selectionsAllowed = True
            End If
            
        Else
            selectionsAllowed = False
        End If
    Else
        selectionsAllowed = False
    End If
End Function

'When certain selection settings are enabled or disabled, corresponding controls are shown or hidden.  To keep the
' panel concise and clean, we move other controls up or down depending on what controls are visible.
Private Sub updateSelectionPanelLayout()

    'Display the feathering slider as necessary
    If cmbSelSmoothing(0).ListIndex = sFullyFeathered Then
        sltSelectionFeathering.Visible = True
    Else
        sltSelectionFeathering.Visible = False
    End If
    
    'Display the border slider as necessary
    If cmbSelType(0).ListIndex = sBorder Then
        sltSelectionBorder.Visible = True
    Else
        sltSelectionBorder.Visible = False
    End If
    
End Sub

'When the selection text boxes are updated, change the scrollbars to match
Private Sub tudSel_Change(Index As Integer)
    updateSelectionsValuesViaText
End Sub

Private Sub updateSelectionsValuesViaText()
    If selectionsAllowed(True) Then
        If Not pdImages(g_CurrentImage).mainSelection.rejectRefreshRequests Then
            pdImages(g_CurrentImage).mainSelection.updateViaTextBox
            RenderViewport pdImages(g_CurrentImage), FormMain.mainCanvas(0)
        End If
    End If
End Sub

'External functions can use this to re-theme this form at run-time (important when changing languages, for example)
Public Sub requestMakeFormPretty()
    makeFormPretty Me, m_ToolTip
End Sub
