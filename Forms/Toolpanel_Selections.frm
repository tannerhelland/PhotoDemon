VERSION 5.00
Begin VB.Form toolpanel_Selections 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   1515
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16650
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   101
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1110
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.PictureBox picSelectionSubcontainer 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1470
      Index           =   1
      Left            =   5340
      ScaleHeight     =   98
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   665
      TabIndex        =   34
      Top             =   0
      Width           =   9975
      Begin PhotoDemon.pdComboBox cboSelArea 
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   35
         Top             =   390
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   635
      End
      Begin PhotoDemon.sliderTextCombo sltSelectionBorder 
         CausesValidation=   0   'False
         Height          =   405
         Index           =   1
         Left            =   0
         TabIndex        =   36
         Top             =   840
         Width           =   2670
         _ExtentX        =   4710
         _ExtentY        =   873
         Min             =   1
         Max             =   10000
         Value           =   1
      End
      Begin PhotoDemon.textUpDown tudSel 
         Height          =   345
         Index           =   4
         Left            =   2820
         TabIndex        =   37
         Top             =   375
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   714
         Min             =   -30000
         Max             =   30000
      End
      Begin PhotoDemon.textUpDown tudSel 
         Height          =   345
         Index           =   5
         Left            =   2820
         TabIndex        =   38
         Top             =   885
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   714
         Min             =   -30000
         Max             =   30000
      End
      Begin PhotoDemon.textUpDown tudSel 
         Height          =   345
         Index           =   6
         Left            =   4380
         TabIndex        =   39
         Top             =   375
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   714
         Min             =   -30000
         Max             =   30000
      End
      Begin PhotoDemon.textUpDown tudSel 
         Height          =   345
         Index           =   7
         Left            =   4380
         TabIndex        =   0
         Top             =   885
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   714
         Min             =   -30000
         Max             =   30000
      End
      Begin PhotoDemon.pdLabel lblSelection 
         Height          =   240
         Index           =   7
         Left            =   120
         Top             =   60
         Width           =   2415
         _ExtentX        =   0
         _ExtentY        =   503
         Caption         =   "area"
      End
      Begin PhotoDemon.pdLabel lblSelection 
         Height          =   240
         Index           =   3
         Left            =   4380
         Top             =   60
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   503
         Caption         =   "size (w, h)"
      End
      Begin PhotoDemon.pdLabel lblSelection 
         Height          =   240
         Index           =   0
         Left            =   2820
         Top             =   60
         Width           =   1395
         _ExtentX        =   0
         _ExtentY        =   503
         Caption         =   "position (x, y)"
      End
   End
   Begin VB.PictureBox picSelectionSubcontainer 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1470
      Index           =   2
      Left            =   5340
      ScaleHeight     =   98
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   665
      TabIndex        =   26
      Top             =   0
      Width           =   9975
      Begin PhotoDemon.pdComboBox cboSelArea 
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   27
         Top             =   390
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   635
      End
      Begin PhotoDemon.sliderTextCombo sltSelectionBorder 
         CausesValidation=   0   'False
         Height          =   405
         Index           =   2
         Left            =   0
         TabIndex        =   28
         Top             =   840
         Width           =   2670
         _ExtentX        =   4710
         _ExtentY        =   873
         Min             =   1
         Max             =   10000
         Value           =   1
      End
      Begin PhotoDemon.textUpDown tudSel 
         Height          =   345
         Index           =   8
         Left            =   2820
         TabIndex        =   29
         Top             =   375
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   714
         Min             =   -30000
         Max             =   30000
      End
      Begin PhotoDemon.textUpDown tudSel 
         Height          =   345
         Index           =   9
         Left            =   2820
         TabIndex        =   30
         Top             =   885
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   714
         Min             =   -30000
         Max             =   30000
      End
      Begin PhotoDemon.textUpDown tudSel 
         Height          =   345
         Index           =   10
         Left            =   4380
         TabIndex        =   31
         Top             =   375
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   714
         Min             =   -30000
         Max             =   30000
      End
      Begin PhotoDemon.textUpDown tudSel 
         Height          =   345
         Index           =   11
         Left            =   4380
         TabIndex        =   32
         Top             =   885
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   714
         Min             =   -30000
         Max             =   30000
      End
      Begin PhotoDemon.sliderTextCombo sltSelectionLineWidth 
         CausesValidation=   0   'False
         Height          =   405
         Left            =   5880
         TabIndex        =   33
         Top             =   360
         Width           =   2670
         _ExtentX        =   4710
         _ExtentY        =   873
         Min             =   1
         Max             =   10000
         Value           =   1
      End
      Begin PhotoDemon.pdLabel lblSelection 
         Height          =   240
         Index           =   18
         Left            =   6000
         Top             =   60
         Width           =   2505
         _ExtentX        =   0
         _ExtentY        =   503
         Caption         =   "line width"
      End
      Begin PhotoDemon.pdLabel lblSelection 
         Height          =   240
         Index           =   11
         Left            =   120
         Top             =   60
         Width           =   2415
         _ExtentX        =   0
         _ExtentY        =   503
         Caption         =   "area"
      End
      Begin PhotoDemon.pdLabel lblSelection 
         Height          =   240
         Index           =   10
         Left            =   4380
         Top             =   60
         Width           =   1425
         _ExtentX        =   0
         _ExtentY        =   503
         Caption         =   "2nd point (x, y)"
      End
      Begin PhotoDemon.pdLabel lblSelection 
         Height          =   240
         Index           =   9
         Left            =   2820
         Top             =   60
         Width           =   1485
         _ExtentX        =   0
         _ExtentY        =   503
         Caption         =   "1st point (x, y)"
      End
   End
   Begin VB.PictureBox picSelectionSubcontainer 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1470
      Index           =   3
      Left            =   5340
      ScaleHeight     =   98
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   665
      TabIndex        =   22
      Top             =   0
      Width           =   9975
      Begin PhotoDemon.pdComboBox cboSelArea 
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   23
         Top             =   390
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   635
      End
      Begin PhotoDemon.sliderTextCombo sltSelectionBorder 
         CausesValidation=   0   'False
         Height          =   405
         Index           =   3
         Left            =   0
         TabIndex        =   24
         Top             =   840
         Width           =   2670
         _ExtentX        =   4710
         _ExtentY        =   873
         Min             =   1
         Max             =   10000
         Value           =   1
      End
      Begin PhotoDemon.sliderTextCombo sltPolygonCurvature 
         CausesValidation=   0   'False
         Height          =   405
         Left            =   2760
         TabIndex        =   25
         Top             =   360
         Width           =   2670
         _ExtentX        =   4710
         _ExtentY        =   873
         Max             =   1
         SigDigits       =   2
      End
      Begin PhotoDemon.pdLabel lblSelection 
         Height          =   240
         Index           =   17
         Left            =   2910
         Top             =   60
         Width           =   2490
         _ExtentX        =   0
         _ExtentY        =   503
         Caption         =   "curvature"
      End
      Begin PhotoDemon.pdLabel lblSelection 
         Height          =   240
         Index           =   15
         Left            =   120
         Top             =   60
         Width           =   2415
         _ExtentX        =   0
         _ExtentY        =   503
         Caption         =   "area"
      End
   End
   Begin VB.PictureBox picSelectionSubcontainer 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1470
      Index           =   4
      Left            =   5340
      ScaleHeight     =   98
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   665
      TabIndex        =   18
      Top             =   0
      Width           =   9975
      Begin PhotoDemon.pdComboBox cboSelArea 
         Height          =   375
         Index           =   4
         Left            =   120
         TabIndex        =   19
         Top             =   390
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   635
      End
      Begin PhotoDemon.sliderTextCombo sltSelectionBorder 
         CausesValidation=   0   'False
         Height          =   405
         Index           =   4
         Left            =   0
         TabIndex        =   20
         Top             =   840
         Width           =   2670
         _ExtentX        =   4710
         _ExtentY        =   873
         Min             =   1
         Max             =   10000
         Value           =   1
      End
      Begin PhotoDemon.sliderTextCombo sltSmoothStroke 
         CausesValidation=   0   'False
         Height          =   405
         Left            =   2760
         TabIndex        =   21
         Top             =   360
         Visible         =   0   'False
         Width           =   2670
         _ExtentX        =   4710
         _ExtentY        =   873
         Max             =   1
         SigDigits       =   2
      End
      Begin PhotoDemon.pdLabel lblSelection 
         Height          =   240
         Index           =   19
         Left            =   2910
         Top             =   60
         Visible         =   0   'False
         Width           =   2550
         _ExtentX        =   0
         _ExtentY        =   503
         Caption         =   "stroke smoothing"
      End
      Begin PhotoDemon.pdLabel lblSelection 
         Height          =   240
         Index           =   12
         Left            =   120
         Top             =   60
         Width           =   2415
         _ExtentX        =   0
         _ExtentY        =   503
         Caption         =   "area"
      End
   End
   Begin VB.PictureBox picSelectionSubcontainer 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1470
      Index           =   5
      Left            =   5340
      ScaleHeight     =   98
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   665
      TabIndex        =   13
      Top             =   0
      Width           =   9975
      Begin PhotoDemon.pdComboBox cboWandCompare 
         Height          =   375
         Left            =   3270
         TabIndex        =   14
         Top             =   855
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   635
      End
      Begin PhotoDemon.buttonStrip btsWandArea 
         Height          =   825
         Left            =   120
         TabIndex        =   15
         Top             =   405
         Width           =   2895
         _ExtentX        =   4366
         _ExtentY        =   1455
      End
      Begin PhotoDemon.sliderTextCombo sltWandTolerance 
         CausesValidation=   0   'False
         Height          =   405
         Left            =   3120
         TabIndex        =   16
         Top             =   360
         Width           =   2670
         _ExtentX        =   4710
         _ExtentY        =   873
         Max             =   255
         SigDigits       =   1
      End
      Begin PhotoDemon.buttonStrip btsWandMerge 
         Height          =   825
         Left            =   6120
         TabIndex        =   17
         Top             =   405
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   1455
      End
      Begin PhotoDemon.pdLabel lblSelection 
         Height          =   240
         Index           =   16
         Left            =   6120
         Top             =   60
         Width           =   2895
         _ExtentX        =   0
         _ExtentY        =   503
         Caption         =   "sampling area"
      End
      Begin PhotoDemon.pdLabel lblSelection 
         Height          =   240
         Index           =   6
         Left            =   3240
         Top             =   60
         Width           =   2595
         _ExtentX        =   0
         _ExtentY        =   503
         Caption         =   "tolerance"
      End
      Begin PhotoDemon.pdLabel lblSelection 
         Height          =   240
         Index           =   14
         Left            =   120
         Top             =   60
         Width           =   2895
         _ExtentX        =   0
         _ExtentY        =   503
         Caption         =   "area"
      End
   End
   Begin VB.PictureBox picSelectionSubcontainer 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1470
      Index           =   0
      Left            =   5340
      ScaleHeight     =   98
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   665
      TabIndex        =   5
      Top             =   0
      Width           =   9975
      Begin PhotoDemon.pdComboBox cboSelArea 
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   390
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   635
      End
      Begin PhotoDemon.sliderTextCombo sltSelectionBorder 
         CausesValidation=   0   'False
         Height          =   405
         Index           =   0
         Left            =   0
         TabIndex        =   7
         Top             =   840
         Width           =   2670
         _ExtentX        =   4710
         _ExtentY        =   873
         Min             =   1
         Max             =   10000
         Value           =   1
      End
      Begin PhotoDemon.textUpDown tudSel 
         Height          =   345
         Index           =   0
         Left            =   2820
         TabIndex        =   8
         Top             =   375
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   714
         Min             =   -30000
         Max             =   30000
      End
      Begin PhotoDemon.textUpDown tudSel 
         Height          =   345
         Index           =   1
         Left            =   2820
         TabIndex        =   9
         Top             =   885
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   714
         Min             =   -30000
         Max             =   30000
      End
      Begin PhotoDemon.textUpDown tudSel 
         Height          =   345
         Index           =   2
         Left            =   4380
         TabIndex        =   10
         Top             =   375
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   714
         Min             =   -30000
         Max             =   30000
      End
      Begin PhotoDemon.textUpDown tudSel 
         Height          =   345
         Index           =   3
         Left            =   4380
         TabIndex        =   11
         Top             =   885
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   714
         Min             =   -30000
         Max             =   30000
      End
      Begin PhotoDemon.sliderTextCombo sltCornerRounding 
         CausesValidation=   0   'False
         Height          =   405
         Left            =   5760
         TabIndex        =   12
         Top             =   345
         Width           =   2670
         _ExtentX        =   4710
         _ExtentY        =   873
         Max             =   1
      End
      Begin PhotoDemon.pdLabel lblSelection 
         Height          =   240
         Index           =   5
         Left            =   5880
         Top             =   60
         Width           =   2445
         _ExtentX        =   0
         _ExtentY        =   503
         Caption         =   "corner rounding"
      End
      Begin PhotoDemon.pdLabel lblSelection 
         Height          =   240
         Index           =   1
         Left            =   2820
         Top             =   60
         Width           =   1395
         _ExtentX        =   0
         _ExtentY        =   503
         Caption         =   "position (x, y)"
      End
      Begin PhotoDemon.pdLabel lblSelection 
         Height          =   240
         Index           =   2
         Left            =   4380
         Top             =   60
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   503
         Caption         =   "size (w, h)"
      End
      Begin PhotoDemon.pdLabel lblSelection 
         Height          =   240
         Index           =   4
         Left            =   120
         Top             =   60
         Width           =   2415
         _ExtentX        =   0
         _ExtentY        =   503
         Caption         =   "area"
      End
   End
   Begin PhotoDemon.pdComboBox cboSelSmoothing 
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   390
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   635
   End
   Begin PhotoDemon.pdComboBox cboSelRender 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   390
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   635
   End
   Begin PhotoDemon.colorSelector csSelectionHighlight 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   2445
      _ExtentX        =   3916
      _ExtentY        =   661
   End
   Begin PhotoDemon.sliderTextCombo sltSelectionFeathering 
      CausesValidation=   0   'False
      Height          =   405
      Left            =   2640
      TabIndex        =   4
      Top             =   840
      Visible         =   0   'False
      Width           =   2670
      _ExtentX        =   4710
      _ExtentY        =   873
      Max             =   100
   End
   Begin PhotoDemon.pdLabel lblSelection 
      Height          =   240
      Index           =   13
      Left            =   2760
      Top             =   60
      Width           =   2445
      _ExtentX        =   0
      _ExtentY        =   503
      Caption         =   "smoothing"
   End
   Begin PhotoDemon.pdLabel lblSelection 
      Height          =   240
      Index           =   8
      Left            =   120
      Top             =   60
      Width           =   2445
      _ExtentX        =   0
      _ExtentY        =   503
      Caption         =   "appearance"
   End
End
Attribute VB_Name = "toolpanel_Selections"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Selection Tool Panel
'Copyright 2013-2015 by Tanner Helland
'Created: 02/Oct/13
'Last updated: 13/May/15
'Last update: finish migrating all relevant controls to this dedicated form
'
'This form includes all user-editable settings for PD's various selection tools.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'The value of all controls on this form are saved and loaded to file by this class
Private WithEvents lastUsedSettings As pdLastUsedSettings
Attribute lastUsedSettings.VB_VarHelpID = -1

Private Sub btsWandArea_Click(ByVal buttonIndex As Long)
    
    'If a selection is already active, change its type to match the current option, then redraw it
    If selectionsAllowed(False) Then
        pdImages(g_CurrentImage).mainSelection.setSelectionProperty SP_WAND_SEARCH_MODE, buttonIndex
        Viewport_Engine.Stage4_CompositeCanvas pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    End If
    
End Sub

Private Sub btsWandMerge_Click(ByVal buttonIndex As Long)

    'If a selection is already active, change its type to match the current option, then redraw it
    If selectionsAllowed(False) Then
        pdImages(g_CurrentImage).mainSelection.setSelectionProperty SP_WAND_SAMPLE_MERGED, buttonIndex
        Viewport_Engine.Stage4_CompositeCanvas pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    End If

End Sub

Private Sub cboSelArea_Click(Index As Integer)

    If cboSelArea(Index).ListIndex = sBorder Then
        sltSelectionBorder(Index).Visible = True
    Else
        sltSelectionBorder(Index).Visible = False
    End If
    
    'If a selection is already active, change its type to match the current selection, then redraw it
    If selectionsAllowed(False) Then
        pdImages(g_CurrentImage).mainSelection.setSelectionProperty SP_AREA, cboSelArea(Index).ListIndex
        pdImages(g_CurrentImage).mainSelection.setSelectionProperty SP_BORDER_WIDTH, sltSelectionBorder(Index).Value
        Viewport_Engine.Stage4_CompositeCanvas pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    End If
    
End Sub

Private Sub cboSelRender_Click()

    'Show or hide the color selector, as appropriate
    If cboSelRender.ListIndex = SELECTION_RENDER_HIGHLIGHT Then
        csSelectionHighlight.Visible = True
    Else
        csSelectionHighlight.Visible = False
    End If
    
    'Redraw the viewport
    If selectionsAllowed(False) Then Viewport_Engine.Stage4_CompositeCanvas pdImages(g_CurrentImage), FormMain.mainCanvas(0)

End Sub

Private Sub cboSelSmoothing_Click()

    updateSelectionPanelLayout
    
    'If a selection is already active, change its type to match the current selection, then redraw it
    If selectionsAllowed(False) Then
        pdImages(g_CurrentImage).mainSelection.setSelectionProperty SP_SMOOTHING, cboSelSmoothing.ListIndex
        pdImages(g_CurrentImage).mainSelection.setSelectionProperty SP_FEATHERING_RADIUS, sltSelectionFeathering.Value
        Viewport_Engine.Stage4_CompositeCanvas pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    End If

End Sub

Private Sub cboWandCompare_Click()
    
    'Limit the accuracy of the tolerance for certain comparison methods.
    If cboWandCompare.ListIndex > 1 Then
        sltWandTolerance.SigDigits = 0
    Else
        sltWandTolerance.SigDigits = 1
    End If
    
    'If a selection is already active, change its type to match the current option, then redraw it
    If selectionsAllowed(False) Then
        pdImages(g_CurrentImage).mainSelection.setSelectionProperty SP_WAND_COMPARE_METHOD, cboWandCompare.ListIndex
        Viewport_Engine.Stage4_CompositeCanvas pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    End If
    
End Sub

Private Sub csSelectionHighlight_ColorChanged()
    
    'Redraw the viewport
    If selectionsAllowed(False) Then Viewport_Engine.Stage4_CompositeCanvas pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    
End Sub

Private Sub Form_Load()
    
    'Initialize various selection tool settings
    
    'Selection visual styles (Highlight, Lightbox, or Outline)
    cboSelRender.Clear
    cboSelRender.AddItem " Highlight", 0
    cboSelRender.AddItem " Lightbox", 1
    cboSelRender.AddItem " Outline", 2
    cboSelRender.ListIndex = 0
    
    csSelectionHighlight.Color = RGB(255, 58, 72)
    csSelectionHighlight.Visible = True
    
    'Selection smoothing (currently none, antialiased, fully feathered)
    cboSelSmoothing.Clear
    cboSelSmoothing.AddItem " None", 0
    cboSelSmoothing.AddItem " Antialiased", 1
    cboSelSmoothing.AddItem " Feathered", 2
    cboSelSmoothing.ListIndex = 1
    
    'Selection types (currently interior, exterior, border)
    Dim i As Long
    For i = 0 To cboSelArea.Count - 1
        cboSelArea(i).AddItem " Interior", 0
        cboSelArea(i).AddItem " Exterior", 1
        cboSelArea(i).AddItem " Border", 2
        cboSelArea(i).ListIndex = 0
    Next i
    
    'Magic wand options
    btsWandMerge.AddItem "image", 0
    btsWandMerge.AddItem "layer", 1
    btsWandMerge.ListIndex = 0
    
    btsWandArea.AddItem "contiguous", 0
    btsWandArea.AddItem "global", 1
    btsWandArea.ListIndex = 0
    
    cboWandCompare.Clear
    cboWandCompare.AddItem " Composite", 0
    cboWandCompare.AddItem " Color", 1
    cboWandCompare.AddItem " Luminance", 2, True
    cboWandCompare.AddItem " Red", 3
    cboWandCompare.AddItem " Green", 4
    cboWandCompare.AddItem " Blue", 5
    cboWandCompare.AddItem " Alpha", 6
    cboWandCompare.ListIndex = 1
    
    'Load any last-used settings for this form
    Set lastUsedSettings = New pdLastUsedSettings
    lastUsedSettings.setParentForm Me
    lastUsedSettings.loadAllControlValues
    
    'Update everything against the current theme.  This will also set tooltips for various controls.
    UpdateAgainstCurrentTheme

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    'Save all last-used settings to file
    lastUsedSettings.saveAllControlValues
    lastUsedSettings.setParentForm Nothing

End Sub

Private Sub lastUsedSettings_ReadCustomPresetData()
    
    'Reset the selection coordinate boxes to 0
    Dim i As Long
    For i = 0 To tudSel.Count - 1
        tudSel(i) = 0
    Next i

End Sub

Private Sub sltCornerRounding_Change()
    If selectionsAllowed(True) Then
        pdImages(g_CurrentImage).mainSelection.setSelectionProperty SP_ROUNDED_CORNER_RADIUS, sltCornerRounding.Value
        Viewport_Engine.Stage4_CompositeCanvas pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    End If
End Sub

Private Sub sltPolygonCurvature_Change()
    If selectionsAllowed(True) Then
        pdImages(g_CurrentImage).mainSelection.setSelectionProperty SP_POLYGON_CURVATURE, sltPolygonCurvature.Value
        Viewport_Engine.Stage4_CompositeCanvas pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    End If
End Sub

Private Sub sltSelectionBorder_Change(Index As Integer)
    If selectionsAllowed(False) Then
        pdImages(g_CurrentImage).mainSelection.setSelectionProperty SP_BORDER_WIDTH, sltSelectionBorder(Index).Value
        Viewport_Engine.Stage4_CompositeCanvas pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    End If
End Sub

Private Sub sltSelectionFeathering_Change()
    If selectionsAllowed(False) Then
        pdImages(g_CurrentImage).mainSelection.setSelectionProperty SP_FEATHERING_RADIUS, sltSelectionFeathering.Value
        Viewport_Engine.Stage4_CompositeCanvas pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    End If
End Sub

Private Sub sltSelectionLineWidth_Change()
    If selectionsAllowed(True) Then
        pdImages(g_CurrentImage).mainSelection.setSelectionProperty SP_LINE_WIDTH, sltSelectionLineWidth.Value
        Viewport_Engine.Stage4_CompositeCanvas pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    End If
End Sub

'When certain selection settings are enabled or disabled, corresponding controls are shown or hidden.  To keep the
' panel concise and clean, we move other controls up or down depending on what controls are visible.
Public Sub updateSelectionPanelLayout()

    'Display the feathering slider as necessary
    If cboSelSmoothing.ListIndex = sFullyFeathered Then
        sltSelectionFeathering.Visible = True
    Else
        sltSelectionFeathering.Visible = False
    End If
    
    'Display the border slider as necessary
    If (Selection_Handler.getSelectionSubPanelFromCurrentTool < cboSelArea.Count - 1) And (Selection_Handler.getSelectionSubPanelFromCurrentTool > 0) Then
        If cboSelArea(Selection_Handler.getSelectionSubPanelFromCurrentTool).ListIndex = sBorder Then
            sltSelectionBorder(Selection_Handler.getSelectionSubPanelFromCurrentTool).Visible = True
        Else
            sltSelectionBorder(Selection_Handler.getSelectionSubPanelFromCurrentTool).Visible = False
        End If
    End If
    
    'Finally, the magic wand selection type is unique because it cannot display an outline.  (This might someday be possible,
    ' but we would need to construct the border region ourselves - and I'm not a huge fan of the work involved.)
    ' As such, when activating that tool, we need to remove the Outline option, and when switching to a different tool, we need
    ' to restore the option.
    If g_CurrentTool = SELECT_WAND Then
    
        'See if the combo box is already modified
        If cboSelRender.ListCount = 3 Then
            
            'Remove the "outline" option
            If cboSelRender.ListIndex = 2 Then cboSelRender.ListIndex = 0
            cboSelRender.RemoveItem 2
            
        End If
    
    Else
    
        'See if the combo box is missing an entry
        If cboSelRender.ListCount = 2 Then
            cboSelRender.AddItem " Outline", 2
        End If
    
    End If
    
End Sub

Private Sub sltSmoothStroke_Change()
    If selectionsAllowed(False) Then
        pdImages(g_CurrentImage).mainSelection.setSelectionProperty SP_SMOOTH_STROKE, sltSmoothStroke.Value
        Viewport_Engine.Stage4_CompositeCanvas pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    End If
End Sub

Private Sub sltWandTolerance_Change()
    If selectionsAllowed(False) Then
        pdImages(g_CurrentImage).mainSelection.setSelectionProperty SP_WAND_TOLERANCE, sltWandTolerance.Value
        Viewport_Engine.Stage4_CompositeCanvas pdImages(g_CurrentImage), FormMain.mainCanvas(0)
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
            Viewport_Engine.Stage4_CompositeCanvas pdImages(g_CurrentImage), FormMain.mainCanvas(0)
        End If
    End If
End Sub

'Updating against the current theme accomplishes a number of things:
' 1) All user-drawn controls are redrawn according to the current g_Themer settings.
' 2) All tooltips and captions are translated according to the current language.
' 3) MakeFormPretty is called, which redraws the form itself according to any theme and/or system settings.
'
'This function is called at least once, at Form_Load, but can be called again if the active language or theme changes.
Public Sub UpdateAgainstCurrentTheme()

    'Start by redrawing the form according to current theme and translation settings.  (This function also takes care of
    ' any common controls that may still exist in the program.)
    MakeFormPretty Me
    
    'Tooltips must be manually re-assigned according to the current language.  This is a necessary evil, if the user switches
    ' between two non-English languages at run-time.
    cboSelRender.AssignTooltip "Click to change the way selections are rendered onto the image canvas.  This has no bearing on selection contents - only the way they appear while editing."
    cboSelSmoothing.AssignTooltip "This option controls how smoothly a selection blends with its surroundings."
        
    Dim i As Long
    For i = 0 To cboSelArea.Count - 1
        cboSelArea(i).AssignTooltip "These options control the area affected by a selection.  The selection can be modified on-canvas while any of these settings are active.  For more advanced selection adjustments, use the Select menu."
        sltSelectionBorder(i).AssignTooltip "This option adjusts the width of the selection border."
    Next i
    
    sltSelectionFeathering.AssignTooltip "This feathering slider allows for immediate feathering adjustments.  For performance reasons, it is limited to small radii.  For larger feathering radii, please use the Select -> Feathering menu."
    sltCornerRounding.AssignTooltip "This option adjusts the roundness of a rectangular selection's corners."
    sltSelectionLineWidth.AssignTooltip "This option adjusts the width of a line selection."
    
    sltPolygonCurvature.AssignTooltip "This option adjusts the curvature, if any, of a polygon selection's sides."
    sltSmoothStroke.AssignTooltip "This option increases the smoothness of a hand-drawn lasso selection."
    sltWandTolerance.AssignTooltip "Tolerance controls how similar two pixels must be before adding them to a magic wand selection."
    
    btsWandMerge.AssignTooltip "The magic wand can operate on the entire image, or just the active layer."
    btsWandArea.AssignTooltip "Normally, the magic wand will spread out from the target pixel, adding neighboring pixels to the selection as it goes.  You can alternatively set it to search the entire image, without regards for continuity."
    
    cboWandCompare.AssignTooltip "This option controls which criteria the magic wand uses to determine whether a pixel should be added to the current selection."

End Sub
