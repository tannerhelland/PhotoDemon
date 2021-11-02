VERSION 5.00
Begin VB.Form toolpanel_MoveSize 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   4005
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16650
   ControlBox      =   0   'False
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
   Icon            =   "Toolpanel_MoveSize.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   267
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1110
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin PhotoDemon.pdContainer cntrPopOut 
      Height          =   2175
      Left            =   720
      Top             =   1680
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   3836
      Begin PhotoDemon.pdCheckBox chkPopoutTest 
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   1080
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   661
         Caption         =   ""
      End
      Begin PhotoDemon.pdSlider sldPopoutTest 
         Height          =   735
         Left            =   120
         TabIndex        =   17
         Top             =   120
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   1296
      End
   End
   Begin PhotoDemon.pdButtonStripVertical btsMoveOptions 
      Height          =   1320
      Left            =   120
      TabIndex        =   13
      Top             =   60
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   2328
   End
   Begin PhotoDemon.pdContainer ctlMoveContainer 
      Height          =   1455
      Index           =   0
      Left            =   2520
      Top             =   0
      Width           =   14055
      _ExtentX        =   0
      _ExtentY        =   0
      Begin PhotoDemon.pdTitle ttlTest 
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   960
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   661
         Caption         =   "test flyout"
         Value           =   0   'False
      End
      Begin PhotoDemon.pdCheckBox chkAspectRatio 
         Height          =   375
         Left            =   3960
         TabIndex        =   16
         Top             =   810
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   661
         Caption         =   "lock aspect ratio"
         Value           =   0   'False
      End
      Begin PhotoDemon.pdDropDown cboLayerResizeQuality 
         Height          =   690
         Left            =   3960
         TabIndex        =   2
         Top             =   15
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   1217
         Caption         =   "transform quality"
         FontSizeCaption =   10
      End
      Begin PhotoDemon.pdSpinner tudLayerMove 
         Height          =   345
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   420
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   609
      End
      Begin PhotoDemon.pdLabel lblOptions 
         Height          =   240
         Index           =   9
         Left            =   135
         Top             =   30
         Width           =   1650
         _ExtentX        =   2910
         _ExtentY        =   423
         Caption         =   "position (x, y)"
      End
      Begin PhotoDemon.pdLabel lblOptions 
         Height          =   240
         Index           =   10
         Left            =   2040
         Top             =   30
         Width           =   1650
         _ExtentX        =   2910
         _ExtentY        =   423
         Caption         =   "size (w, h)"
      End
      Begin PhotoDemon.pdSpinner tudLayerMove 
         Height          =   345
         Index           =   1
         Left            =   240
         TabIndex        =   4
         Top             =   840
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   609
      End
      Begin PhotoDemon.pdSpinner tudLayerMove 
         Height          =   345
         Index           =   2
         Left            =   2160
         TabIndex        =   5
         Top             =   420
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   609
      End
      Begin PhotoDemon.pdSpinner tudLayerMove 
         Height          =   345
         Index           =   3
         Left            =   2160
         TabIndex        =   6
         Top             =   840
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   609
      End
      Begin PhotoDemon.pdButtonToolbox cmdLayerMove 
         Height          =   570
         Index           =   0
         Left            =   6960
         TabIndex        =   7
         Top             =   360
         Width           =   660
         _ExtentX        =   1164
         _ExtentY        =   1005
         AutoToggle      =   -1  'True
      End
      Begin PhotoDemon.pdLabel lblOptions 
         Height          =   240
         Index           =   12
         Left            =   6840
         Top             =   30
         Width           =   3360
         _ExtentX        =   5927
         _ExtentY        =   503
         Caption         =   "other options"
      End
   End
   Begin PhotoDemon.pdContainer ctlMoveContainer 
      Height          =   1455
      Index           =   1
      Left            =   2520
      Top             =   0
      Width           =   14055
      _ExtentX        =   0
      _ExtentY        =   0
      Begin PhotoDemon.pdSlider sltLayerAngle 
         Height          =   765
         Left            =   120
         TabIndex        =   14
         Top             =   60
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   1349
         Caption         =   "layer angle"
         FontSizeCaption =   10
         Min             =   -360
         Max             =   360
         SigDigits       =   2
      End
      Begin PhotoDemon.pdSlider sltLayerShearX 
         Height          =   765
         Left            =   4080
         TabIndex        =   1
         Top             =   60
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   1349
         Caption         =   "layer shear (x, y)"
         FontSizeCaption =   10
         Min             =   -5
         Max             =   5
         SigDigits       =   2
      End
      Begin PhotoDemon.pdButtonToolbox cmdLayerAffinePermanent 
         Height          =   570
         Left            =   8040
         TabIndex        =   8
         Top             =   360
         Width           =   660
         _ExtentX        =   1164
         _ExtentY        =   1005
         AutoToggle      =   -1  'True
      End
      Begin PhotoDemon.pdSlider sltLayerShearY 
         Height          =   405
         Left            =   4080
         TabIndex        =   12
         Top             =   840
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   714
         Min             =   -5
         Max             =   5
         SigDigits       =   2
      End
      Begin PhotoDemon.pdLabel lblOptions 
         Height          =   240
         Index           =   4
         Left            =   8040
         Top             =   60
         Width           =   3120
         _ExtentX        =   5503
         _ExtentY        =   503
         Caption         =   "other options"
      End
   End
   Begin PhotoDemon.pdContainer ctlMoveContainer 
      Height          =   1455
      Index           =   2
      Left            =   2520
      Top             =   0
      Width           =   14055
      _ExtentX        =   0
      _ExtentY        =   0
      Begin PhotoDemon.pdLabel lblOptions 
         Height          =   240
         Index           =   0
         Left            =   135
         Top             =   75
         Width           =   5370
         _ExtentX        =   9472
         _ExtentY        =   503
         Caption         =   "interaction options"
      End
      Begin PhotoDemon.pdCheckBox chkAutoActivateLayer 
         Height          =   330
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   5250
         _ExtentX        =   9260
         _ExtentY        =   582
         Caption         =   "automatically activate layer beneath mouse"
      End
      Begin PhotoDemon.pdCheckBox chkIgnoreTransparent 
         Height          =   330
         Left            =   240
         TabIndex        =   10
         Top             =   720
         Width           =   5250
         _ExtentX        =   9260
         _ExtentY        =   582
         Caption         =   "ignore transparent pixels when auto-activating layers"
      End
      Begin PhotoDemon.pdCheckBox chkLayerBorder 
         Height          =   330
         Left            =   5760
         TabIndex        =   11
         Top             =   360
         Width           =   5250
         _ExtentX        =   9260
         _ExtentY        =   582
         Caption         =   "show layer borders"
      End
      Begin PhotoDemon.pdCheckBox chkLayerNodes 
         Height          =   330
         Left            =   5760
         TabIndex        =   0
         Top             =   720
         Width           =   5250
         _ExtentX        =   9260
         _ExtentY        =   582
         Caption         =   "show resize nodes"
      End
      Begin PhotoDemon.pdLabel lblOptions 
         Height          =   240
         Index           =   1
         Left            =   5655
         Top             =   75
         Width           =   5370
         _ExtentX        =   9472
         _ExtentY        =   503
         Caption         =   "display options"
      End
      Begin PhotoDemon.pdCheckBox chkRotateNode 
         Height          =   330
         Left            =   5760
         TabIndex        =   15
         Top             =   1080
         Width           =   5250
         _ExtentX        =   9260
         _ExtentY        =   582
         Caption         =   "show rotate nodes"
      End
   End
End
Attribute VB_Name = "toolpanel_MoveSize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Move/Size Tool Panel
'Copyright 2013-2021 by Tanner Helland
'Created: 02/Oct/13
'Last updated: 09/November/20
'Last update: add a dedicated lock for layer aspect ratio (see https://github.com/tannerhelland/PhotoDemon/issues/342)
'
'This form includes all user-editable settings for the Move/Size canvas tool.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************


Option Explicit

'Flyout manager
Private m_Flyout As pdFlyout

'The value of all controls on this form are saved and loaded to file by this class
Private WithEvents m_lastUsedSettings As pdLastUsedSettings
Attribute m_lastUsedSettings.VB_VarHelpID = -1

'Two sub-panels are available on the "move options" panel
Private Sub btsMoveOptions_Click(ByVal buttonIndex As Long)
    UpdateSubpanel
End Sub

Private Sub UpdateSubpanel()
    Dim i As Long
    For i = 0 To ctlMoveContainer.UBound
        ctlMoveContainer(i).Visible = (i = btsMoveOptions.ListIndex)
    Next i
End Sub

Private Sub btsMoveOptions_SetCustomTabTarget(ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If (btsMoveOptions.ListIndex = 0) Then
        If shiftTabWasPressed Then
            If cmdLayerMove(0).Enabled Then newTargetHwnd = cmdLayerMove(0).hWnd Else newTargetHwnd = cboLayerResizeQuality.hWnd
        Else
            If tudLayerMove(0).Enabled Then newTargetHwnd = tudLayerMove(0).hWnd Else newTargetHwnd = cboLayerResizeQuality.hWnd
        End If
    ElseIf (btsMoveOptions.ListIndex = 1) Then
        If shiftTabWasPressed Then
            If cmdLayerAffinePermanent.Enabled Then newTargetHwnd = cmdLayerAffinePermanent.hWnd Else newTargetHwnd = sltLayerShearY.hWndSpinner
        End If
    End If
End Sub

Private Sub cboLayerResizeQuality_Click()
    
    'If tool changes are not allowed, exit.
    ' NOTE: this will also check tool busy status, via Tools.getToolBusyState
    If (Not Tools.CanvasToolsAllowed) Then Exit Sub
    
    'Mark the tool engine as busy
    Tools.SetToolBusyState True
    
    'Apply the new quality mode
    PDImages.GetActiveImage.GetActiveLayer.SetLayerResizeQuality cboLayerResizeQuality.ListIndex
    
    'Free the tool engine
    Tools.SetToolBusyState False
    
    'Redraw the viewport
    Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    
End Sub

Private Sub cboLayerResizeQuality_GotFocusAPI()
    If (Not PDImages.IsImageActive()) Then Exit Sub
    Processor.FlagInitialNDFXState_Generic pgp_ResizeQuality, cboLayerResizeQuality.ListIndex, PDImages.GetActiveImage.GetActiveLayerID
End Sub

Private Sub cboLayerResizeQuality_LostFocusAPI()
    Processor.FlagFinalNDFXState_Generic pgp_ResizeQuality, cboLayerResizeQuality.ListIndex
End Sub

Private Sub cboLayerResizeQuality_SetCustomTabTarget(ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If shiftTabWasPressed Then
        newTargetHwnd = tudLayerMove(3).hWnd
    Else
        If cmdLayerMove(0).Enabled Then newTargetHwnd = cmdLayerMove(0).hWnd Else newTargetHwnd = btsMoveOptions.hWnd
    End If
End Sub

Private Sub chkAspectRatio_Click()
    
    'If tool changes are not allowed, exit.
    ' NOTE: this will also check tool busy status, via Tools.getToolBusyState
    If (Not Tools.CanvasToolsAllowed) Then Exit Sub
    
    'Mark the tool engine as busy
    Tools.SetToolBusyState True
    
    'When clicked, lock the Y aspect ratio to the X aspect ratio
    If chkAspectRatio.Value Then PDImages.GetActiveImage.GetActiveLayer.SetLayerCanvasYModifier PDImages.GetActiveImage.GetActiveLayer.GetLayerCanvasXModifier()
    
    'Free the tool engine
    Tools.SetToolBusyState False
    
    'Redraw the viewport
    Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    
    'Also, activate the "make transforms permanent" button(s) as necessary
    If (cmdLayerAffinePermanent.Enabled <> PDImages.GetActiveImage.GetActiveLayer.AffineTransformsActive(True)) Then cmdLayerAffinePermanent.Enabled = PDImages.GetActiveImage.GetActiveLayer.AffineTransformsActive(True)
    If (cmdLayerMove(0).Enabled <> PDImages.GetActiveImage.GetActiveLayer.AffineTransformsActive(True)) Then cmdLayerMove(0).Enabled = PDImages.GetActiveImage.GetActiveLayer.AffineTransformsActive(True)

End Sub

Private Sub chkAspectRatio_GotFocusAPI()
    If (Not Tools.CanvasToolsAllowed) Then Exit Sub
    Processor.FlagInitialNDFXState_Generic pgp_CanvasYModifier, tudLayerMove(3).Value / PDImages.GetActiveImage.GetActiveLayer.GetLayerHeight(False), PDImages.GetActiveImage.GetActiveLayerID
End Sub

Private Sub chkAspectRatio_LostFocusAPI()
    If (Not Tools.CanvasToolsAllowed) Then Exit Sub
    Processor.FlagFinalNDFXState_Generic pgp_CanvasYModifier, tudLayerMove(3).Value / PDImages.GetActiveImage.GetActiveLayer.GetLayerHeight(False)
End Sub

'En/disable the "ignore transparent layer bits on click activations" setting if the auto-activate clicked layer setting changes
Private Sub chkAutoActivateLayer_Click()
    chkIgnoreTransparent.Enabled = chkAutoActivateLayer
End Sub

Private Sub chkAutoActivateLayer_SetCustomTabTarget(ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If (Not shiftTabWasPressed) Then newTargetHwnd = chkIgnoreTransparent.hWnd
End Sub

Private Sub chkIgnoreTransparent_SetCustomTabTarget(ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If shiftTabWasPressed Then
        newTargetHwnd = chkAutoActivateLayer.hWnd
    Else
        newTargetHwnd = chkLayerBorder.hWnd
    End If
End Sub

'Show/hide layer borders while using the move tool
Private Sub chkLayerBorder_Click()
    Tools_Move.SetDrawLayerBorders chkLayerBorder.Value
    Viewport.Stage4_FlipBufferAndDrawUI PDImages.GetActiveImage(), FormMain.MainCanvas(0)
End Sub

Private Sub chkLayerBorder_SetCustomTabTarget(ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If shiftTabWasPressed Then
        newTargetHwnd = chkIgnoreTransparent.hWnd
    Else
        newTargetHwnd = chkLayerNodes.hWnd
    End If
End Sub

'Show/hide layer transform nodes while using the move tool
Private Sub chkLayerNodes_Click()
    Tools_Move.SetDrawLayerCornerNodes chkLayerNodes.Value
    Viewport.Stage4_FlipBufferAndDrawUI PDImages.GetActiveImage(), FormMain.MainCanvas(0)
End Sub

Private Sub chkLayerNodes_SetCustomTabTarget(ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If shiftTabWasPressed Then
        newTargetHwnd = chkLayerBorder.hWnd
    Else
        newTargetHwnd = chkRotateNode.hWnd
    End If
End Sub

Private Sub chkRotateNode_Click()
    Tools_Move.SetDrawLayerRotateNodes chkRotateNode.Value
    Viewport.Stage4_FlipBufferAndDrawUI PDImages.GetActiveImage(), FormMain.MainCanvas(0)
End Sub

Private Sub chkRotateNode_SetCustomTabTarget(ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If shiftTabWasPressed Then newTargetHwnd = chkLayerNodes.hWnd
End Sub

Private Sub cmdLayerAffinePermanent_Click(ByVal Shift As ShiftConstants)
    If (Not PDImages.IsImageActive()) Then Exit Sub
    Process "Make layer changes permanent", , BuildParamList("layerindex", PDImages.GetActiveImage.GetActiveLayerIndex), UNDO_Layer
End Sub

Private Sub cmdLayerAffinePermanent_SetCustomTabTarget(ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If shiftTabWasPressed Then
        newTargetHwnd = sltLayerShearY.hWndSpinner
    Else
        newTargetHwnd = btsMoveOptions.hWnd
    End If
End Sub

Private Sub cmdLayerMove_Click(Index As Integer, ByVal Shift As ShiftConstants)
    
    If (Not PDImages.IsImageActive()) Then Exit Sub
    
    Select Case Index
    
        'Make non-destructive resize permanent
        Case 0
            Process "Make layer changes permanent", , BuildParamList("layerindex", PDImages.GetActiveImage.GetActiveLayerIndex), UNDO_Layer
    
    End Select
    
End Sub

Private Sub cmdLayerMove_SetCustomTabTarget(Index As Integer, ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If (Index = 0) And (Not shiftTabWasPressed) Then newTargetHwnd = btsMoveOptions.hWnd
End Sub

Private Sub Form_Load()
    
    Tools.SetToolBusyState True
    
    'Initialize move tool panels
    btsMoveOptions.AddItem "position and size", 0
    btsMoveOptions.AddItem "angle and shear", 1
    btsMoveOptions.AddItem "tool settings", 2
    btsMoveOptions.ListIndex = 0
    UpdateSubpanel
    
    'Several reset/apply buttons on this form have very similar purposes
    cmdLayerMove(0).AssignTooltip "Make current layer transforms (size, angle, and shear) permanent.  This action is never required, but if viewport rendering is sluggish, it may improve performance."
    cmdLayerAffinePermanent.AssignTooltip "Make current layer transforms (size, angle, and shear) permanent.  This action is never required, but if viewport rendering is sluggish, it may improve performance."
    
    cboLayerResizeQuality.SetAutomaticRedraws False
    cboLayerResizeQuality.Clear
    cboLayerResizeQuality.AddItem "nearest-neighbor", 0
    cboLayerResizeQuality.AddItem "bilinear", 1
    cboLayerResizeQuality.AddItem "bicubic", 2
    cboLayerResizeQuality.SetAutomaticRedraws True
    cboLayerResizeQuality.ListIndex = 1
    
    'Ensure our corresponding tool manager is synchronized with default layer rendering styles
    Tools_Move.SetDrawLayerBorders chkLayerBorder.Value
    Tools_Move.SetDrawLayerCornerNodes chkLayerNodes.Value
    Tools_Move.SetDrawLayerRotateNodes chkRotateNode.Value
    
    'Load any last-used settings for this form
    'NOTE: this is currently disabled, as all settings on this form are synched to the active layer
    Set m_lastUsedSettings = New pdLastUsedSettings
    m_lastUsedSettings.SetParentForm Me
    m_lastUsedSettings.LoadAllControlValues True
    
    Tools.SetToolBusyState False
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    'Save all last-used settings to file
    If (Not m_lastUsedSettings Is Nothing) Then
        m_lastUsedSettings.SaveAllControlValues True
        m_lastUsedSettings.SetParentForm Nothing
    End If
    
    'Failsafe only
    If (Not m_Flyout Is Nothing) Then m_Flyout.HideFlyout
    Set m_Flyout = Nothing
    
End Sub

'Because this tool synchronizes the vast majority of its properties to the active layer
' in the current image, we do *not* automatically save all settings for this tool.
' Instead, we custom-save only the settings we absolutely want.
Private Sub m_LastUsedSettings_AddCustomPresetData()

    With m_lastUsedSettings
        .AddPresetData "move-size-auto-activate", chkAutoActivateLayer.Value
        .AddPresetData "move-size-ignore-transparent", chkIgnoreTransparent.Value
        .AddPresetData "move-size-show-layer-borders", chkLayerBorder.Value
        .AddPresetData "move-size-show-resize-nodes", chkLayerNodes.Value
        .AddPresetData "move-size-show-rotate-nodes", chkRotateNode.Value
        .AddPresetData "move-size-lock-aspect-ratio", chkAspectRatio.Value
    End With

End Sub

Private Sub m_LastUsedSettings_ReadCustomPresetData()

    With m_lastUsedSettings
        chkAutoActivateLayer.Value = .RetrievePresetData("move-size-auto-activate", True)
        chkIgnoreTransparent.Value = .RetrievePresetData("move-size-ignore-transparent", True)
        chkLayerBorder.Value = .RetrievePresetData("move-size-show-layer-borders", True)
        chkLayerNodes.Value = .RetrievePresetData("move-size-show-resize-nodes", True)
        chkRotateNode.Value = .RetrievePresetData("move-size-show-rotate-nodes", True)
        
        'The "lock aspect ratio" control is tricky, because we don't want to set this value
        ' if the current image has variable aspect ratio enabled; otherwise, it will forcibly
        ' modify the active layer's size value(s).
        Dim okToLoad As Boolean
        okToLoad = True
        
        If Tools.CanvasToolsAllowed(False) Then
            okToLoad = (PDImages.GetActiveImage.GetActiveLayer.GetLayerCanvasXModifier = 1#) And (PDImages.GetActiveImage.GetActiveLayer.GetLayerCanvasYModifier = 1#)
        End If
        
        If okToLoad Then
            chkAspectRatio.Value = .RetrievePresetData("move-size-lock-aspect-ratio", False)
        Else
            chkAspectRatio.Value = False
        End If
        
    End With
    
End Sub

Private Sub sltLayerAngle_Change()
    
    'If tool changes are not allowed, exit.
    ' NOTE: this will also check tool busy status, via Tools.getToolBusyState
    If (Not Tools.CanvasToolsAllowed) Then Exit Sub
    
    'Mark the tool engine as busy
    Tools.SetToolBusyState True
    
    'Notify the layer of the setting change
    PDImages.GetActiveImage.GetActiveLayer.SetLayerAngle sltLayerAngle.Value
    
    'Free the tool engine
    Tools.SetToolBusyState False
    
    'Redraw the viewport
    Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    
    'Also, activate the "make transforms permanent" button(s) as necessary
    If (cmdLayerAffinePermanent.Enabled <> PDImages.GetActiveImage.GetActiveLayer.AffineTransformsActive(True)) Then cmdLayerAffinePermanent.Enabled = PDImages.GetActiveImage.GetActiveLayer.AffineTransformsActive(True)
    If (cmdLayerMove(0).Enabled <> PDImages.GetActiveImage.GetActiveLayer.AffineTransformsActive(True)) Then cmdLayerMove(0).Enabled = PDImages.GetActiveImage.GetActiveLayer.AffineTransformsActive(True)
    
End Sub

Private Sub sltLayerAngle_FinalChange()
    FinalChangeHandler
End Sub

Private Sub sltLayerAngle_GotFocusAPI()
    If (Not PDImages.IsImageActive()) Then Exit Sub
    Processor.FlagInitialNDFXState_Generic pgp_Angle, sltLayerAngle.Value, PDImages.GetActiveImage.GetActiveLayerID
End Sub

Private Sub sltLayerAngle_LostFocusAPI()
    Processor.FlagFinalNDFXState_Generic pgp_Angle, sltLayerAngle.Value
End Sub

Private Sub sltLayerShearX_Change()
    
    'If tool changes are not allowed, exit.
    ' NOTE: this will also check tool busy status, via Tools.getToolBusyState
    If (Not Tools.CanvasToolsAllowed) Then Exit Sub
    
    'Mark the tool engine as busy
    Tools.SetToolBusyState True
    
    'Notify the layer of the setting change
    PDImages.GetActiveImage.GetActiveLayer.SetLayerShearX sltLayerShearX.Value
    
    'Free the tool engine
    Tools.SetToolBusyState False
    
    'Redraw the viewport
    Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    
    'Also, activate the "make transforms permanent" button(s) as necessary
    If (cmdLayerAffinePermanent.Enabled <> PDImages.GetActiveImage.GetActiveLayer.AffineTransformsActive(True)) Then cmdLayerAffinePermanent.Enabled = PDImages.GetActiveImage.GetActiveLayer.AffineTransformsActive(True)
    If (cmdLayerMove(0).Enabled <> PDImages.GetActiveImage.GetActiveLayer.AffineTransformsActive(True)) Then cmdLayerMove(0).Enabled = PDImages.GetActiveImage.GetActiveLayer.AffineTransformsActive(True)
    
End Sub

Private Sub sltLayerShearX_FinalChange()
    FinalChangeHandler
End Sub

Private Sub sltLayerShearX_GotFocusAPI()
    If (Not PDImages.IsImageActive()) Then Exit Sub
    Processor.FlagInitialNDFXState_Generic pgp_ShearX, sltLayerShearX.Value, PDImages.GetActiveImage.GetActiveLayerID
End Sub

Private Sub sltLayerShearX_LostFocusAPI()
    Processor.FlagFinalNDFXState_Generic pgp_ShearX, sltLayerShearX.Value
End Sub

Private Sub sltLayerShearX_SetCustomTabTarget(ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If (Not shiftTabWasPressed) Then newTargetHwnd = sltLayerShearY.hWndSlider
End Sub

Private Sub sltLayerShearY_Change()
    
    'If tool changes are not allowed, exit.
    ' NOTE: this will also check tool busy status, via Tools.getToolBusyState
    If (Not Tools.CanvasToolsAllowed) Then Exit Sub
    
    'Mark the tool engine as busy
    Tools.SetToolBusyState True
    
    'Notify the layer of the setting change
    PDImages.GetActiveImage.GetActiveLayer.SetLayerShearY sltLayerShearY.Value
    
    'Free the tool engine
    Tools.SetToolBusyState False
    
    'Redraw the viewport
    Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    
    'Also, activate the "make transforms permanent" button(s) as necessary
    If (cmdLayerAffinePermanent.Enabled <> PDImages.GetActiveImage.GetActiveLayer.AffineTransformsActive(True)) Then cmdLayerAffinePermanent.Enabled = PDImages.GetActiveImage.GetActiveLayer.AffineTransformsActive(True)
    If (cmdLayerMove(0).Enabled <> PDImages.GetActiveImage.GetActiveLayer.AffineTransformsActive(True)) Then cmdLayerMove(0).Enabled = PDImages.GetActiveImage.GetActiveLayer.AffineTransformsActive(True)
    
End Sub

Private Sub sltLayerShearY_FinalChange()
    FinalChangeHandler
End Sub

Private Sub sltLayerShearY_GotFocusAPI()
    If (Not PDImages.IsImageActive()) Then Exit Sub
    Processor.FlagInitialNDFXState_Generic pgp_ShearY, sltLayerShearY.Value, PDImages.GetActiveImage.GetActiveLayerID
End Sub

Private Sub sltLayerShearY_LostFocusAPI()
    Processor.FlagFinalNDFXState_Generic pgp_ShearY, sltLayerShearY.Value
End Sub

Private Sub sltLayerShearY_SetCustomTabTarget(ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If shiftTabWasPressed Then
        newTargetHwnd = sltLayerShearX.hWndSpinner
    Else
        If cmdLayerAffinePermanent.Enabled Then newTargetHwnd = cmdLayerAffinePermanent.hWnd Else newTargetHwnd = btsMoveOptions.hWnd
    End If
End Sub

Private Sub ttlTest_Click(ByVal newState As Boolean)
    
    If newState Then
        If (m_Flyout Is Nothing) Then Set m_Flyout = New pdFlyout
        m_Flyout.ShowFlyout Me, ttlTest, cntrPopOut
    Else
        If (Not m_Flyout Is Nothing) Then m_Flyout.HideFlyout
    End If

End Sub

Private Sub tudLayerMove_Change(Index As Integer)
    
    'If tool changes are not allowed, exit.
    ' NOTE: this will also check tool busy status, via Tools.getToolBusyState
    If (Not Tools.CanvasToolsAllowed) Then Exit Sub
    
    'Mark the tool engine as busy
    Tools.SetToolBusyState True
    
    Select Case Index
    
        'Layer position (x)
        Case 0
            PDImages.GetActiveImage.GetActiveLayer.SetLayerOffsetX tudLayerMove(Index).Value
        
        'Layer position (y)
        Case 1
            PDImages.GetActiveImage.GetActiveLayer.SetLayerOffsetY tudLayerMove(Index).Value
        
        'Layer width
        Case 2
            PDImages.GetActiveImage.GetActiveLayer.SetLayerCanvasXModifier tudLayerMove(Index).Value / PDImages.GetActiveImage.GetActiveLayer.GetLayerWidth(False)
            If chkAspectRatio.Value Then
                PDImages.GetActiveImage.GetActiveLayer.SetLayerCanvasYModifier PDImages.GetActiveImage.GetActiveLayer.GetLayerCanvasXModifier()
                toolpanel_MoveSize.tudLayerMove(3).Value = PDImages.GetActiveImage.GetActiveLayer.GetLayerHeight(True)
            End If
            
        'Layer height
        Case 3
            PDImages.GetActiveImage.GetActiveLayer.SetLayerCanvasYModifier tudLayerMove(Index).Value / PDImages.GetActiveImage.GetActiveLayer.GetLayerHeight(False)
            If chkAspectRatio.Value Then
                PDImages.GetActiveImage.GetActiveLayer.SetLayerCanvasXModifier PDImages.GetActiveImage.GetActiveLayer.GetLayerCanvasYModifier()
                toolpanel_MoveSize.tudLayerMove(2).Value = PDImages.GetActiveImage.GetActiveLayer.GetLayerWidth(True)
            End If
        
    End Select
    
    'Free the tool engine
    Tools.SetToolBusyState False
    
    'Redraw the viewport
    Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    
    'Also, activate the "make transforms permanent" button(s) as necessary
    If (cmdLayerAffinePermanent.Enabled <> PDImages.GetActiveImage.GetActiveLayer.AffineTransformsActive(True)) Then cmdLayerAffinePermanent.Enabled = PDImages.GetActiveImage.GetActiveLayer.AffineTransformsActive(True)
    If (cmdLayerMove(0).Enabled <> PDImages.GetActiveImage.GetActiveLayer.AffineTransformsActive(True)) Then cmdLayerMove(0).Enabled = PDImages.GetActiveImage.GetActiveLayer.AffineTransformsActive(True)

End Sub

'Non-destructive resizing requires the synchronization of several menus, as well.  Because it's time-consuming to invoke
' SyncInterfaceToCurrentImage, we only call it when the user releases the mouse.
Private Sub tudLayerMove_FinalChange(Index As Integer)
    FinalChangeHandler True
End Sub

Private Sub tudLayerMove_GotFocusAPI(Index As Integer)
    If (Not PDImages.IsImageActive()) Then Exit Sub
    If (Index = 0) Then
        Processor.FlagInitialNDFXState_Generic pgp_OffsetX, tudLayerMove(Index).Value, PDImages.GetActiveImage.GetActiveLayerID
    ElseIf (Index = 1) Then
        Processor.FlagInitialNDFXState_Generic pgp_OffsetY, tudLayerMove(Index).Value, PDImages.GetActiveImage.GetActiveLayerID
    ElseIf (Index = 2) Then
        Processor.FlagInitialNDFXState_Generic pgp_CanvasXModifier, tudLayerMove(Index).Value / PDImages.GetActiveImage.GetActiveLayer.GetLayerWidth(False), PDImages.GetActiveImage.GetActiveLayerID
    ElseIf (Index = 3) Then
        Processor.FlagInitialNDFXState_Generic pgp_CanvasYModifier, tudLayerMove(Index).Value / PDImages.GetActiveImage.GetActiveLayer.GetLayerHeight(False), PDImages.GetActiveImage.GetActiveLayerID
    End If
End Sub

Private Sub tudLayerMove_LostFocusAPI(Index As Integer)
    If (Not PDImages.IsImageActive()) Then Exit Sub
    If (Index = 0) Then
        Processor.FlagFinalNDFXState_Generic pgp_OffsetX, tudLayerMove(Index).Value
    ElseIf (Index = 1) Then
        Processor.FlagFinalNDFXState_Generic pgp_OffsetY, tudLayerMove(Index).Value
    ElseIf (Index = 2) Then
        Processor.FlagFinalNDFXState_Generic pgp_CanvasXModifier, tudLayerMove(Index).Value / PDImages.GetActiveImage.GetActiveLayer.GetLayerWidth(False)
    ElseIf (Index = 3) Then
        Processor.FlagFinalNDFXState_Generic pgp_CanvasYModifier, tudLayerMove(Index).Value / PDImages.GetActiveImage.GetActiveLayer.GetLayerHeight(False)
    End If
End Sub

'When a layer property control (either a slider or spinner) has a "FinalChange" event, use this control to update any
' necessary UI elements that may need to be adjusted due to non-destructive changes.
Private Sub FinalChangeHandler(Optional ByVal performFullUISync As Boolean = False)
    Processor.NDFXUiUpdate
    If performFullUISync Then Interface.SyncInterfaceToCurrentImage
End Sub

'Updating against the current theme accomplishes a number of things:
' 1) All user-drawn controls are redrawn according to the current g_Themer settings.
' 2) All tooltips and captions are translated according to the current language.
' 3) ApplyThemeAndTranslations is called, which redraws the form itself according to any theme and/or system settings.
'
'This function is called at least once, at Form_Load, but can be called again if the active language or theme changes.
Public Sub UpdateAgainstCurrentTheme()
    
    'UI images must be updated against theme-specific colors
    Dim buttonSize As Long
    buttonSize = Interface.FixDPI(32)
    cmdLayerMove(0).AssignImage "generic_commit", , buttonSize, buttonSize
    cmdLayerAffinePermanent.AssignImage "generic_commit", , buttonSize, buttonSize
    
    Interface.ApplyThemeAndTranslations Me
    
End Sub

'Because these controls are laid out in a non-standard pattern, we want to manually specify tab and
' shift+tab focus targets.
Private Sub tudLayerMove_SetCustomTabTarget(Index As Integer, ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    
    If shiftTabWasPressed Then
        If (Index > 0) Then
            newTargetHwnd = tudLayerMove(Index - 1).hWnd
        Else
            newTargetHwnd = btsMoveOptions.hWnd
        End If
    Else
        If (Index < 3) Then
            newTargetHwnd = tudLayerMove(Index + 1).hWnd
        Else
            newTargetHwnd = cboLayerResizeQuality.hWnd
        End If
    End If

End Sub
