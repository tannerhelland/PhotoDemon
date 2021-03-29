VERSION 5.00
Begin VB.Form toolpanel_Selections 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   1515
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16650
   DrawStyle       =   5  'Transparent
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HasDC           =   0   'False
   Icon            =   "Toolpanel_Selections.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   101
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1110
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin PhotoDemon.pdSpinner spnOpacity 
      Height          =   375
      Index           =   0
      Left            =   1440
      TabIndex        =   28
      Top             =   840
      Width           =   1125
      _ExtentX        =   1931
      _ExtentY        =   661
      DefaultValue    =   50
      Min             =   1
      Max             =   100
      Value           =   50
   End
   Begin PhotoDemon.pdSpinner spnOpacity 
      Height          =   375
      Index           =   1
      Left            =   1440
      TabIndex        =   29
      Top             =   840
      Width           =   1125
      _ExtentX        =   1931
      _ExtentY        =   661
      DefaultValue    =   50
      Min             =   1
      Max             =   100
      Value           =   50
   End
   Begin PhotoDemon.pdDropDown cboSelSmoothing 
      Height          =   735
      Left            =   2760
      TabIndex        =   0
      Top             =   30
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1296
      Caption         =   "smoothing"
      FontSizeCaption =   10
   End
   Begin PhotoDemon.pdDropDown cboSelRender 
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   30
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1296
      Caption         =   "appearance"
      FontSizeCaption =   10
   End
   Begin PhotoDemon.pdColorSelector csSelection 
      Height          =   330
      Index           =   0
      Left            =   225
      TabIndex        =   2
      Top             =   840
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   582
   End
   Begin PhotoDemon.pdSlider sltSelectionFeathering 
      CausesValidation=   0   'False
      Height          =   405
      Left            =   2760
      TabIndex        =   3
      Top             =   840
      Visible         =   0   'False
      Width           =   2550
      _ExtentX        =   4498
      _ExtentY        =   714
      Max             =   100
   End
   Begin PhotoDemon.pdColorSelector csSelection 
      Height          =   330
      Index           =   1
      Left            =   225
      TabIndex        =   30
      Top             =   840
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   582
   End
   Begin PhotoDemon.pdContainer ctlGroupSelectionSubcontainer 
      Height          =   1470
      Index           =   0
      Left            =   5340
      Top             =   0
      Width           =   9975
      _ExtentX        =   0
      _ExtentY        =   0
      Begin PhotoDemon.pdLabel lblColon 
         Height          =   375
         Index           =   0
         Left            =   3960
         Top             =   855
         Visible         =   0   'False
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Alignment       =   2
         Caption         =   ":"
         FontSize        =   12
      End
      Begin PhotoDemon.pdButtonToolbox cmdLock 
         Height          =   360
         Index           =   1
         Left            =   5520
         TabIndex        =   27
         Top             =   855
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   635
         DontHighlightDownState=   -1  'True
         StickyToggle    =   -1  'True
      End
      Begin PhotoDemon.pdDropDown cboSize 
         Height          =   735
         Index           =   0
         Left            =   2760
         TabIndex        =   31
         Top             =   30
         Width           =   3135
         _ExtentX        =   6376
         _ExtentY        =   1296
         Caption         =   "dimensions"
         FontSizeCaption =   10
      End
      Begin PhotoDemon.pdDropDown cboSelArea 
         Height          =   735
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   30
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   1296
         Caption         =   "area"
         FontSizeCaption =   10
      End
      Begin PhotoDemon.pdSlider sltSelectionBorder 
         CausesValidation=   0   'False
         Height          =   405
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   2550
         _ExtentX        =   4498
         _ExtentY        =   714
         Min             =   1
         Max             =   1000
         ScaleStyle      =   2
         Value           =   1
         DefaultValue    =   1
      End
      Begin PhotoDemon.pdSpinner tudSel 
         Height          =   345
         Index           =   0
         Left            =   2865
         TabIndex        =   7
         Top             =   870
         Width           =   1080
         _ExtentX        =   2328
         _ExtentY        =   714
         Min             =   -30000
         Max             =   30000
         ShowResetButton =   0   'False
      End
      Begin PhotoDemon.pdSpinner tudSel 
         Height          =   345
         Index           =   1
         Left            =   4440
         TabIndex        =   8
         Top             =   870
         Width           =   1080
         _ExtentX        =   2328
         _ExtentY        =   714
         Min             =   -30000
         Max             =   30000
         ShowResetButton =   0   'False
      End
      Begin PhotoDemon.pdSlider sltCornerRounding 
         CausesValidation=   0   'False
         Height          =   735
         Left            =   6120
         TabIndex        =   9
         Top             =   30
         Width           =   2670
         _ExtentX        =   4710
         _ExtentY        =   1296
         Caption         =   "corner rounding"
         FontSizeCaption =   10
         Max             =   100
         SigDigits       =   1
      End
      Begin PhotoDemon.pdButtonToolbox cmdLock 
         Height          =   360
         Index           =   0
         Left            =   3960
         TabIndex        =   4
         Top             =   855
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   635
         DontHighlightDownState=   -1  'True
         StickyToggle    =   -1  'True
      End
   End
   Begin PhotoDemon.pdContainer ctlGroupSelectionSubcontainer 
      Height          =   1470
      Index           =   1
      Left            =   5340
      Top             =   0
      Width           =   9975
      _ExtentX        =   0
      _ExtentY        =   0
      Begin PhotoDemon.pdDropDown cboSelArea 
         Height          =   735
         Index           =   1
         Left            =   120
         TabIndex        =   25
         Top             =   30
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   1296
         Caption         =   "area"
         FontSizeCaption =   10
      End
      Begin PhotoDemon.pdSlider sltSelectionBorder 
         CausesValidation=   0   'False
         Height          =   405
         Index           =   1
         Left            =   120
         TabIndex        =   26
         Top             =   840
         Width           =   2550
         _ExtentX        =   4498
         _ExtentY        =   714
         Min             =   1
         Max             =   1000
         ScaleStyle      =   2
         Value           =   1
         DefaultValue    =   1
      End
      Begin PhotoDemon.pdDropDown cboSize 
         Height          =   735
         Index           =   1
         Left            =   2760
         TabIndex        =   10
         Top             =   30
         Width           =   3135
         _ExtentX        =   6376
         _ExtentY        =   1296
         Caption         =   "dimensions"
         FontSizeCaption =   10
      End
      Begin PhotoDemon.pdButtonToolbox cmdLock 
         Height          =   360
         Index           =   3
         Left            =   5520
         TabIndex        =   15
         Top             =   855
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   635
         DontHighlightDownState=   -1  'True
         StickyToggle    =   -1  'True
      End
      Begin PhotoDemon.pdSpinner tudSel 
         Height          =   345
         Index           =   2
         Left            =   2865
         TabIndex        =   19
         Top             =   870
         Width           =   1080
         _ExtentX        =   2328
         _ExtentY        =   714
         Min             =   -30000
         Max             =   30000
         ShowResetButton =   0   'False
      End
      Begin PhotoDemon.pdSpinner tudSel 
         Height          =   345
         Index           =   3
         Left            =   4440
         TabIndex        =   23
         Top             =   870
         Width           =   1080
         _ExtentX        =   2328
         _ExtentY        =   714
         Min             =   -30000
         Max             =   30000
         ShowResetButton =   0   'False
      End
      Begin PhotoDemon.pdButtonToolbox cmdLock 
         Height          =   360
         Index           =   2
         Left            =   3960
         TabIndex        =   24
         Top             =   855
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   635
         DontHighlightDownState=   -1  'True
         StickyToggle    =   -1  'True
      End
      Begin PhotoDemon.pdLabel lblColon 
         Height          =   375
         Index           =   1
         Left            =   3960
         Top             =   855
         Visible         =   0   'False
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Alignment       =   2
         Caption         =   ":"
         FontSize        =   12
      End
   End
   Begin PhotoDemon.pdContainer ctlGroupSelectionSubcontainer 
      Height          =   1470
      Index           =   2
      Left            =   5340
      Top             =   0
      Width           =   9975
      _ExtentX        =   0
      _ExtentY        =   0
      Begin PhotoDemon.pdDropDown cboSelArea 
         Height          =   735
         Index           =   2
         Left            =   120
         TabIndex        =   20
         Top             =   30
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   1296
         Caption         =   "area"
         FontSizeCaption =   10
      End
      Begin PhotoDemon.pdSlider sltSelectionBorder 
         CausesValidation=   0   'False
         Height          =   405
         Index           =   2
         Left            =   120
         TabIndex        =   21
         Top             =   840
         Width           =   2550
         _ExtentX        =   4498
         _ExtentY        =   714
         Min             =   1
         Max             =   1000
         ScaleStyle      =   2
         Value           =   1
         DefaultValue    =   1
      End
      Begin PhotoDemon.pdSlider sltPolygonCurvature 
         CausesValidation=   0   'False
         Height          =   735
         Left            =   2760
         TabIndex        =   22
         Top             =   30
         Width           =   2670
         _ExtentX        =   4710
         _ExtentY        =   1296
         Caption         =   "curvature"
         FontSizeCaption =   10
         Max             =   1
         SigDigits       =   2
      End
   End
   Begin PhotoDemon.pdContainer ctlGroupSelectionSubcontainer 
      Height          =   1470
      Index           =   3
      Left            =   5340
      Top             =   0
      Width           =   9975
      _ExtentX        =   0
      _ExtentY        =   0
      Begin PhotoDemon.pdDropDown cboSelArea 
         Height          =   735
         Index           =   3
         Left            =   120
         TabIndex        =   16
         Top             =   30
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   1296
         Caption         =   "area"
         FontSizeCaption =   10
      End
      Begin PhotoDemon.pdSlider sltSelectionBorder 
         CausesValidation=   0   'False
         Height          =   405
         Index           =   3
         Left            =   120
         TabIndex        =   17
         Top             =   840
         Width           =   2550
         _ExtentX        =   4498
         _ExtentY        =   714
         Min             =   1
         Max             =   1000
         ScaleStyle      =   2
         Value           =   1
         DefaultValue    =   1
      End
      Begin PhotoDemon.pdSlider sltSmoothStroke 
         CausesValidation=   0   'False
         Height          =   735
         Left            =   2760
         TabIndex        =   18
         Top             =   30
         Visible         =   0   'False
         Width           =   2670
         _ExtentX        =   4710
         _ExtentY        =   1296
         Caption         =   "stroke smoothing"
         FontSizeCaption =   10
         Max             =   1
         SigDigits       =   2
      End
   End
   Begin PhotoDemon.pdContainer ctlGroupSelectionSubcontainer 
      Height          =   1470
      Index           =   4
      Left            =   5340
      Top             =   0
      Width           =   9975
      _ExtentX        =   0
      _ExtentY        =   0
      Begin PhotoDemon.pdDropDown cboWandCompare 
         Height          =   375
         Left            =   3300
         TabIndex        =   11
         Top             =   825
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   661
      End
      Begin PhotoDemon.pdButtonStrip btsWandArea 
         Height          =   1185
         Left            =   120
         TabIndex        =   12
         Top             =   30
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   2090
         Caption         =   "area"
         FontSizeCaption =   10
      End
      Begin PhotoDemon.pdSlider sltWandTolerance 
         CausesValidation=   0   'False
         Height          =   675
         Left            =   3240
         TabIndex        =   13
         Top             =   30
         Width           =   2550
         _ExtentX        =   4498
         _ExtentY        =   1191
         Caption         =   "tolerance"
         FontSizeCaption =   10
         Max             =   100
         SigDigits       =   1
         ScaleStyle      =   1
         Value           =   15
         DefaultValue    =   15
      End
      Begin PhotoDemon.pdButtonStrip btsWandMerge 
         Height          =   1185
         Left            =   6120
         TabIndex        =   14
         Top             =   30
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   2090
         Caption         =   "sampling area"
         FontSizeCaption =   10
      End
   End
End
Attribute VB_Name = "toolpanel_Selections"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Selection Tool Panel
'Copyright 2013-2021 by Tanner Helland
'Created: 02/Oct/13
'Last updated: 03/May/18
'Last update: rework UI to support locking width/height/aspect-ratio for certain selection types
'
'This form includes all user-editable settings for PD's various selection tools.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'The value of all controls on this form are saved and loaded to file by this class
Private WithEvents lastUsedSettings As pdLastUsedSettings
Attribute lastUsedSettings.VB_VarHelpID = -1

Private Sub btsWandArea_Click(ByVal buttonIndex As Long)
    
    'If a selection is already active, change its type to match the current option, then redraw it
    If SelectionsAllowed(False) And (g_CurrentTool = Selections.GetRelevantToolFromSelectShape()) Then
        PDImages.GetActiveImage.MainSelection.SetSelectionProperty sp_WandSearchMode, buttonIndex
        Viewport.Stage3_CompositeCanvas PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    End If
    
End Sub

Private Sub btsWandMerge_Click(ByVal buttonIndex As Long)

    'If a selection is already active, change its type to match the current option, then redraw it
    If SelectionsAllowed(False) And (g_CurrentTool = Selections.GetRelevantToolFromSelectShape()) Then
        PDImages.GetActiveImage.MainSelection.SetSelectionProperty sp_WandSampleMerged, buttonIndex
        Viewport.Stage3_CompositeCanvas PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    End If

End Sub

Private Sub cboSelArea_Click(Index As Integer)

    sltSelectionBorder(Index).Visible = (cboSelArea(Index).ListIndex = sa_Border)
    
    'If a selection is already active, change its type to match the current selection, then redraw it
    If SelectionsAllowed(False) And (g_CurrentTool = Selections.GetRelevantToolFromSelectShape()) Then
        PDImages.GetActiveImage.MainSelection.SetSelectionProperty sp_Area, cboSelArea(Index).ListIndex
        PDImages.GetActiveImage.MainSelection.SetSelectionProperty sp_BorderWidth, sltSelectionBorder(Index).Value
        Viewport.Stage3_CompositeCanvas PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    End If
    
End Sub

'The "selection rendering technique" dropdown always receives event processing, even if the current selection
' tool does not match the active selection shape.  (Other tool changes are typically restricted by selection type.)
Private Sub cboSelRender_Click()
    
    'Show or hide the color selector, as appropriate
    csSelection(0).Visible = (cboSelRender.ListIndex = PDSR_Highlight)
    spnOpacity(0).Visible = (cboSelRender.ListIndex = PDSR_Highlight)
    csSelection(1).Visible = (cboSelRender.ListIndex = PDSR_Lightbox)
    spnOpacity(1).Visible = (cboSelRender.ListIndex = PDSR_Lightbox)
    
    'Redraw the viewport
    Selections.NotifySelectionRenderChange pdsr_RenderMode, cboSelRender.ListIndex
    If SelectionsAllowed(False) Then Viewport.Stage3_CompositeCanvas PDImages.GetActiveImage(), FormMain.MainCanvas(0)

End Sub

'Selection smoothing is handled universally, even if the current selection shape does not match the active
' selection tool.  (This is done because selection smoothing is universally supported across all shapes.)
Private Sub cboSelSmoothing_Click()

    UpdateSelectionPanelLayout
    
    'If a selection is already active, change its type to match the current selection, then redraw it
    If SelectionsAllowed(False) Then
        PDImages.GetActiveImage.MainSelection.SetSelectionProperty sp_Smoothing, cboSelSmoothing.ListIndex
        PDImages.GetActiveImage.MainSelection.SetSelectionProperty sp_FeatheringRadius, sltSelectionFeathering.Value
        Viewport.Stage3_CompositeCanvas PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    End If

End Sub

Private Sub cboSize_Click(Index As Integer)
    Selections.SyncTextToCurrentSelection PDImages.GetActiveImageID()
End Sub

Private Sub cboWandCompare_Click()
    
    'Limit the accuracy of the tolerance for certain comparison methods.
    If (cboWandCompare.ListIndex > 1) Then sltWandTolerance.SigDigits = 0 Else sltWandTolerance.SigDigits = 1
    
    'If a selection is already active, change its type to match the current option, then redraw it
    If SelectionsAllowed(False) And (g_CurrentTool = Selections.GetRelevantToolFromSelectShape()) Then
        PDImages.GetActiveImage.MainSelection.SetSelectionProperty sp_WandCompareMethod, cboWandCompare.ListIndex
        Viewport.Stage3_CompositeCanvas PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    End If
    
End Sub

Private Sub cmdLock_Click(Index As Integer, ByVal Shift As ShiftConstants)
    
    'Ignore lock actions unless a selection is active, *and* the current selection tool matches the currently
    ' active selection.
    If SelectionsAllowed(False) And (g_CurrentTool = Selections.GetRelevantToolFromSelectShape()) Then
        
        Dim lockedValue As Variant
        
        'Because of the way the cmdLock buttons are structured (with *two* instances per button, one for the
        ' rectangular selection tool and another for the elliptical selection tool), we have to perform some
        ' manual remapping of indices based on the active tool and the active selection attribute
        ' (position/size/aspect ratio).
        Dim relevantIndex As PD_SelectionLockable
        If (g_CurrentTool = SELECT_RECT) Then
            If (cboSize(0).ListIndex = 1) Then
                relevantIndex = Index
                lockedValue = tudSel(Index).Value
            Else
                relevantIndex = pdsl_AspectRatio
                If (tudSel(1).Value <> 0) Then lockedValue = tudSel(0).Value / tudSel(1).Value
            End If
        Else
            If (cboSize(1).ListIndex = 1) Then
                relevantIndex = Index - 2
                lockedValue = tudSel(Index).Value
            Else
                relevantIndex = pdsl_AspectRatio
                If (tudSel(3).Value <> 0) Then lockedValue = tudSel(2).Value / tudSel(3).Value
            End If
        End If
        
        'In the case of aspect ratio vs width/height locks, we don't see both controls at the same time so we
        ' don't have to manually synchronize any UI elements.  Width and height changes are different, however,
        ' because locking one necessarily unlocks the other.
        If cmdLock(Index).Value Then
            If (relevantIndex = pdsl_Width) Then
                cmdLock(Index + 1).Value = False
            ElseIf (relevantIndex = pdsl_Height) Then
                cmdLock(Index - 1).Value = False
            End If
        End If
        
        If cmdLock(Index).Value Then
            PDImages.GetActiveImage.MainSelection.LockProperty relevantIndex, lockedValue
        Else
            PDImages.GetActiveImage.MainSelection.UnlockProperty relevantIndex
        End If
        
    End If
    
End Sub

Private Sub csSelection_ColorChanged(Index As Integer)
    
    If (Index = 0) Then
        Selections.NotifySelectionRenderChange pdsr_HighlightColor, csSelection(Index).Color
    ElseIf (Index = 1) Then
        Selections.NotifySelectionRenderChange pdsr_LightboxColor, csSelection(Index).Color
    End If
    
    If SelectionsAllowed(False) Then Viewport.Stage3_CompositeCanvas PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    
End Sub

Private Sub Form_Load()
    
    'Suspend any visual updates while the form is being loaded
    Viewport.DisableRendering
    
    Dim suspendActive As Boolean
    If PDImages.IsImageActive() Then
        suspendActive = True
        PDImages.GetActiveImage.MainSelection.SuspendAutoRefresh True
    End If
    
    'Initialize various selection tool settings
    
    'Selection visual styles (Highlight, Lightbox, or Outline)
    cboSelRender.SetAutomaticRedraws False
    cboSelRender.Clear
    cboSelRender.AddItem "highlight", 0
    cboSelRender.AddItem "lightbox", 1
    cboSelRender.AddItem "ants", 2
    cboSelRender.AddItem "outline", 3
    cboSelRender.ListIndex = 2
    cboSelRender.SetAutomaticRedraws True
    
    csSelection(0).Color = RGB(255, 58, 72)
    csSelection(0).Visible = True
    spnOpacity(0).Value = 50
    spnOpacity(0).Visible = True
    
    csSelection(1).Color = 0
    csSelection(1).Visible = False
    spnOpacity(1).Value = 50
    spnOpacity(1).Visible = False
    
    'Selection smoothing (currently none, antialiased, fully feathered)
    cboSelSmoothing.SetAutomaticRedraws False
    cboSelSmoothing.Clear
    cboSelSmoothing.AddItem "none", 0
    cboSelSmoothing.AddItem "antialiased", 1
    cboSelSmoothing.AddItem "feathered", 2
    cboSelSmoothing.SetAutomaticRedraws True
    cboSelSmoothing.ListIndex = 1
    
    'Selection types (currently interior, exterior, border)
    Dim i As Long
    For i = 0 To cboSelArea.Count - 1
        cboSelArea(i).SetAutomaticRedraws False
        cboSelArea(i).AddItem "interior", 0
        cboSelArea(i).AddItem "exterior", 1
        cboSelArea(i).AddItem "border", 2
        cboSelArea(i).ListIndex = 0
        cboSelArea(i).SetAutomaticRedraws True
    Next i
    
    'Rectangular and elliptical selections support different sizing modes
    For i = 0 To cboSize.Count - 1
        cboSize(i).SetAutomaticRedraws False
        cboSize(i).AddItem "position (x, y)", 0
        cboSize(i).AddItem "size (w, h)", 1
        cboSize(i).AddItem "aspect ratio", 2
        cboSize(i).ListIndex = 0
        cboSize(i).SetAutomaticRedraws True
    Next i
    
    'Magic wand options
    btsWandMerge.AddItem "image", 0
    btsWandMerge.AddItem "layer", 1
    btsWandMerge.ListIndex = 0
    
    btsWandArea.AddItem "contiguous", 0
    btsWandArea.AddItem "global", 1
    btsWandArea.ListIndex = 0
    
    Interface.PopulateFloodFillTypes cboWandCompare
    
    'Load any last-used settings for this form
    Set lastUsedSettings = New pdLastUsedSettings
    lastUsedSettings.SetParentForm Me
    lastUsedSettings.LoadAllControlValues
    
    If suspendActive Then PDImages.GetActiveImage.MainSelection.SuspendAutoRefresh False
    
    'If a selection is already active, synchronize all UI elements to match
    If suspendActive Then
        If PDImages.GetActiveImage.IsSelectionActive Then Selections.SyncTextToCurrentSelection PDImages.GetActiveImageID()
    End If
    
    Viewport.EnableRendering
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    'Save all last-used settings to file
    If (Not lastUsedSettings Is Nothing) Then
        lastUsedSettings.SaveAllControlValues
        lastUsedSettings.SetParentForm Nothing
    End If

End Sub

Private Sub lastUsedSettings_ReadCustomPresetData()
    
    'Reset the selection coordinate boxes to 0
    Dim i As Long
    For i = 0 To tudSel.Count - 1
        tudSel(i).Value = 0
    Next i
    
    'Selection properties always default to *unlocked*
    cmdLock(0).Value = False
    
    'Pull certain universal selection settings from PD's main preferences file
    If UserPrefs.IsReady Then
        cboSelRender.ListIndex = Selections.GetSelectionRenderMode()
        csSelection(0).Color = Selections.GetSelectionColor_Highlight()
        spnOpacity(0).Value = Selections.GetSelectionOpacity_Highlight()
        csSelection(1).Color = Selections.GetSelectionColor_Lightbox()
        spnOpacity(1).Value = Selections.GetSelectionOpacity_Lightbox()
    End If
    
End Sub

Private Sub sltCornerRounding_Change()
    If SelectionsAllowed(True) And (g_CurrentTool = Selections.GetRelevantToolFromSelectShape()) Then
        PDImages.GetActiveImage.MainSelection.SetSelectionProperty sp_RoundedCornerRadius, sltCornerRounding.Value
        Viewport.Stage3_CompositeCanvas PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    End If
End Sub

Private Sub sltPolygonCurvature_Change()
    If SelectionsAllowed(True) And (g_CurrentTool = Selections.GetRelevantToolFromSelectShape()) Then
        PDImages.GetActiveImage.MainSelection.SetSelectionProperty sp_PolygonCurvature, sltPolygonCurvature.Value
        Viewport.Stage3_CompositeCanvas PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    End If
End Sub

Private Sub sltSelectionBorder_Change(Index As Integer)
    If SelectionsAllowed(False) And (g_CurrentTool = Selections.GetRelevantToolFromSelectShape()) Then
        PDImages.GetActiveImage.MainSelection.SetSelectionProperty sp_BorderWidth, sltSelectionBorder(Index).Value
        Viewport.Stage3_CompositeCanvas PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    End If
End Sub

Private Sub sltSelectionFeathering_Change()
    If SelectionsAllowed(False) Then
        PDImages.GetActiveImage.MainSelection.SetSelectionProperty sp_FeatheringRadius, sltSelectionFeathering.Value
        Viewport.Stage3_CompositeCanvas PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    End If
End Sub

'When certain selection settings are enabled or disabled, corresponding controls are shown or hidden.  To keep the
' panel concise and clean, we move other controls up or down depending on what controls are visible.
Public Sub UpdateSelectionPanelLayout()

    'Display the feathering slider as necessary
    sltSelectionFeathering.Visible = (cboSelSmoothing.ListIndex = es_FullyFeathered)
    
    'Display the border slider as necessary
    If (Selections.GetSelectionSubPanelFromCurrentTool < cboSelArea.Count - 1) And (Selections.GetSelectionSubPanelFromCurrentTool > 0) Then
        sltSelectionBorder(Selections.GetSelectionSubPanelFromCurrentTool).Visible = (cboSelArea(Selections.GetSelectionSubPanelFromCurrentTool).ListIndex = sa_Border)
    End If
    
End Sub

Private Sub sltSmoothStroke_Change()
    If SelectionsAllowed(False) And (g_CurrentTool = Selections.GetRelevantToolFromSelectShape()) Then
        PDImages.GetActiveImage.MainSelection.SetSelectionProperty sp_SmoothStroke, sltSmoothStroke.Value
        Viewport.Stage3_CompositeCanvas PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    End If
End Sub

Private Sub sltWandTolerance_Change()
    If SelectionsAllowed(False) And (g_CurrentTool = Selections.GetRelevantToolFromSelectShape()) Then
        PDImages.GetActiveImage.MainSelection.SetSelectionProperty sp_WandTolerance, sltWandTolerance.Value
        Viewport.Stage3_CompositeCanvas PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    End If
End Sub

Private Sub spnOpacity_Change(Index As Integer)

    If (Index = 0) Then
        Selections.NotifySelectionRenderChange pdsr_HighlightOpacity, spnOpacity(Index).Value
    ElseIf (Index = 1) Then
        Selections.NotifySelectionRenderChange pdsr_LightboxOpacity, spnOpacity(Index).Value
    End If
    
    If SelectionsAllowed(False) Then Viewport.Stage3_CompositeCanvas PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    
End Sub

'When the selection text boxes are updated, change the scrollbars to match
Private Sub tudSel_Change(Index As Integer)
    UpdateSelectionsValuesViaText Index
End Sub

'All text boxes wrap this function.  Note that text box changes are not relayed unless the current selection shape
' matches the current selection tool.
Private Sub UpdateSelectionsValuesViaText(ByVal Index As Integer)
    If SelectionsAllowed(True) Then
        If (Not PDImages.GetActiveImage.MainSelection.GetAutoRefreshSuspend) And (g_CurrentTool = Selections.GetRelevantToolFromSelectShape()) Then
            PDImages.GetActiveImage.MainSelection.UpdateViaTextBox Index
            Viewport.Stage3_CompositeCanvas PDImages.GetActiveImage(), FormMain.MainCanvas(0)
        End If
    End If
End Sub

'Updating against the current theme accomplishes a number of things:
' 1) All user-drawn controls are redrawn according to the current g_Themer settings.
' 2) All tooltips and captions are translated according to the current language.
' 3) ApplyThemeAndTranslations is called, which redraws the form itself according to any theme and/or system settings.
'
'This function is called at least once, at Form_Load, but can be called again if the active language or theme changes.
Public Sub UpdateAgainstCurrentTheme()
    
    Dim buttonImageSize As Long
    buttonImageSize = Interface.FixDPI(16)
    
    Dim i As Long
    For i = 0 To cmdLock.Count - 1
        cmdLock(i).AssignImage "generic_unlock", , buttonImageSize, buttonImageSize
        cmdLock(i).AssignImage_Pressed "generic_lock", , buttonImageSize, buttonImageSize
        cmdLock(i).AssignTooltip "Lock this value.  (Only one value can be locked at a time.  If you lock a new value, previously locked values will unlock.)"
    Next i
    
    'Redrawing the form according to current theme and translation settings.
    ApplyThemeAndTranslations Me
    
    'Tooltips must be manually re-assigned according to the current language.  This is a necessary evil, if the user switches
    ' between two non-English languages at run-time.
    cboSelRender.AssignTooltip "Click to change the way selections are rendered onto the image canvas.  This has no bearing on selection contents - only the way they appear while editing."
    cboSelSmoothing.AssignTooltip "This option controls how smoothly a selection blends with its surroundings."
    
    For i = 0 To cboSelArea.Count - 1
        cboSelArea(i).AssignTooltip "These options control the area affected by a selection.  The selection can be modified on-canvas while any of these settings are active.  For more advanced selection adjustments, use the Select menu."
        sltSelectionBorder(i).AssignTooltip "This option adjusts the width of the selection border."
    Next i
    
    sltSelectionFeathering.AssignTooltip "This feathering slider allows for immediate feathering adjustments.  For performance reasons, it is limited to small radii.  For larger feathering radii, please use the Select -> Feathering menu."
    sltCornerRounding.AssignTooltip "This option adjusts the roundness of a rectangular selection's corners."
    
    sltPolygonCurvature.AssignTooltip "This option adjusts the curvature, if any, of a polygon selection's sides."
    sltSmoothStroke.AssignTooltip "This option increases the smoothness of a hand-drawn lasso selection."
    sltWandTolerance.AssignTooltip "Tolerance controls how similar two pixels must be before adding them to a magic wand selection."
    
    btsWandMerge.AssignTooltip "The magic wand can operate on the entire image, or just the active layer."
    btsWandArea.AssignTooltip "Normally, the magic wand will spread out from the target pixel, adding neighboring pixels to the selection as it goes.  You can alternatively set it to search the entire image, without regards for continuity."
    
    cboWandCompare.AssignTooltip "This option controls which criteria the magic wand uses to determine whether a pixel should be added to the current selection."
    
End Sub
