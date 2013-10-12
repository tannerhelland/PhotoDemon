VERSION 5.00
Begin VB.Form toolbar_Main 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Main"
   ClientHeight    =   8715
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   1050
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
   Begin VB.ComboBox CmbZoom 
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
      ItemData        =   "VBP_ToolbarPrimary.frx":0000
      Left            =   60
      List            =   "VBP_ToolbarPrimary.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   0
      ToolTipText     =   "Click to adjust image zoom"
      Top             =   4320
      Width           =   960
   End
   Begin PhotoDemon.jcbutton cmdOpen 
      Height          =   615
      Left            =   60
      TabIndex        =   1
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
      PictureNormal   =   "VBP_ToolbarPrimary.frx":0004
      DisabledPictureMode=   1
      CaptionEffects  =   0
      TooltipTitle    =   "Open"
   End
   Begin PhotoDemon.jcbutton cmdSave 
      Height          =   615
      Left            =   60
      TabIndex        =   2
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
      PictureNormal   =   "VBP_ToolbarPrimary.frx":1456
      DisabledPictureMode=   1
      CaptionEffects  =   0
      TooltipTitle    =   "Save"
   End
   Begin PhotoDemon.jcbutton cmdUndo 
      Height          =   615
      Left            =   60
      TabIndex        =   3
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
      PictureNormal   =   "VBP_ToolbarPrimary.frx":26B8
      DisabledPictureMode=   1
      CaptionEffects  =   0
      TooltipTitle    =   "Undo"
      TooltipBackColor=   -2147483643
   End
   Begin PhotoDemon.jcbutton cmdRedo 
      Height          =   615
      Left            =   60
      TabIndex        =   4
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
      PictureNormal   =   "VBP_ToolbarPrimary.frx":370A
      DisabledPictureMode=   1
      CaptionEffects  =   0
      TooltipTitle    =   "Redo"
      TooltipBackColor=   -2147483643
   End
   Begin PhotoDemon.jcbutton cmdClose 
      Height          =   615
      Left            =   60
      TabIndex        =   5
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
      PictureNormal   =   "VBP_ToolbarPrimary.frx":475C
      DisabledPictureMode=   1
      CaptionEffects  =   0
      TooltipTitle    =   "Close"
   End
   Begin PhotoDemon.jcbutton cmdSaveAs 
      Height          =   615
      Left            =   60
      TabIndex        =   6
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
      PictureNormal   =   "VBP_ToolbarPrimary.frx":57AE
      DisabledPictureMode=   1
      CaptionEffects  =   0
      TooltipTitle    =   "Save As"
   End
   Begin PhotoDemon.jcbutton cmdZoomIn 
      Height          =   450
      Left            =   525
      TabIndex        =   7
      Top             =   4800
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   794
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
      PictureNormal   =   "VBP_ToolbarPrimary.frx":6A10
      DisabledPictureMode=   1
      CaptionEffects  =   0
      ToolTip         =   "Use this button to increase image zoom."
      TooltipTitle    =   "Zoom In"
   End
   Begin PhotoDemon.jcbutton cmdZoomOut 
      Height          =   450
      Left            =   45
      TabIndex        =   8
      Top             =   4800
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   794
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
      PictureNormal   =   "VBP_ToolbarPrimary.frx":6E62
      DisabledPictureMode=   1
      CaptionEffects  =   0
      ToolTip         =   "Use this button to decrease image zoom."
      TooltipTitle    =   "Zoom Out"
   End
   Begin VB.Label lblImgSize 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "size:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00D1B499&
      Height          =   675
      Left            =   0
      TabIndex        =   11
      Top             =   5460
      Width           =   990
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblCoordinates 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "(X, Y)"
      ForeColor       =   &H00808080&
      Height          =   195
      Left            =   0
      TabIndex        =   10
      Top             =   6240
      Width           =   990
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
      Height          =   1620
      Left            =   30
      TabIndex        =   9
      Top             =   6840
      Visible         =   0   'False
      Width           =   960
      WordWrap        =   -1  'True
   End
   Begin VB.Line lineMain 
      BorderColor     =   &H80000002&
      Index           =   0
      X1              =   2
      X2              =   68
      Y1              =   279
      Y2              =   279
   End
End
Attribute VB_Name = "toolbar_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Primary Toolbar
'Copyright ©2012-2013 by Tanner Helland
'Created: 02/October/13
'Last updated: 03/October/13
'Last update: minor bug-fixes
'
'This form was initially integrated into the main MDI form.  In fall 2013, PhotoDemon left behind the MDI model,
' and all toolbars were moved to their own forms.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Custom tooltip class allows for things like multiline, theming, and multiple monitor support
Dim m_ToolTip As clsToolTip

'When the zoom combo box is changed, redraw the image using the new zoom value
Private Sub CmbZoom_Click()
    
    'Track the current zoom value
    If g_OpenImageCount > 0 Then
        pdImages(g_CurrentImage).CurrentZoomValue = toolbar_Main.CmbZoom.ListIndex
        If toolbar_Main.CmbZoom.ListIndex = 0 Then
            toolbar_Main.cmdZoomIn.Enabled = False
        Else
            If Not toolbar_Main.cmdZoomIn.Enabled Then toolbar_Main.cmdZoomIn.Enabled = True
        End If
        If toolbar_Main.CmbZoom.ListIndex = toolbar_Main.CmbZoom.ListCount - 1 Then
            toolbar_Main.cmdZoomOut.Enabled = False
        Else
            If Not toolbar_Main.cmdZoomOut.Enabled Then toolbar_Main.cmdZoomOut.Enabled = True
        End If
        PrepareViewport pdImages(g_CurrentImage).containingForm, "zoom changed by user"
    End If
    
End Sub

Private Sub cmdClose_Click()
    Unload pdImages(g_CurrentImage).containingForm
End Sub

Private Sub cmdOpen_Click()
    Process "Open", True
End Sub

Private Sub cmdRedo_Click()
    Process "Redo", , , False
End Sub

Private Sub cmdSave_Click()
    Process "Save", , , False
End Sub

Private Sub cmdSaveAs_Click()
    Process "Save as", True, , False
End Sub

Private Sub cmdUndo_Click()
    Process "Undo", , , False
End Sub

Private Sub cmdZoomIn_Click()
    toolbar_Main.CmbZoom.ListIndex = toolbar_Main.CmbZoom.ListIndex - 1
End Sub

Private Sub cmdZoomOut_Click()
    toolbar_Main.CmbZoom.ListIndex = toolbar_Main.CmbZoom.ListIndex + 1
End Sub

Private Sub Form_Load()
    
    'Assign the system hand cursor to all relevant objects
    Set m_ToolTip = New clsToolTip
    makeFormPretty Me, m_ToolTip
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
    g_WindowManager.unregisterForm Me
End Sub
