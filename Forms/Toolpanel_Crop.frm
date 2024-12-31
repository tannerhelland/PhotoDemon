VERSION 5.00
Begin VB.Form toolpanel_Crop 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   6315
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11910
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
   Icon            =   "Toolpanel_Crop.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   421
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   794
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin PhotoDemon.pdContainer cntrPopOut 
      Height          =   975
      Index           =   0
      Left            =   0
      Top             =   960
      Visible         =   0   'False
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   3625
      Begin PhotoDemon.pdLabel lblOptions 
         Height          =   240
         Index           =   10
         Left            =   0
         Top             =   0
         Width           =   3450
         _ExtentX        =   6085
         _ExtentY        =   423
         Caption         =   "size (w, h)"
      End
      Begin PhotoDemon.pdSpinner tudLayerMove 
         Height          =   345
         Index           =   2
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   609
      End
      Begin PhotoDemon.pdSpinner tudLayerMove 
         Height          =   345
         Index           =   3
         Left            =   1680
         TabIndex        =   6
         Top             =   390
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   609
      End
      Begin PhotoDemon.pdButtonToolbox cmdFlyoutLock 
         Height          =   390
         Index           =   0
         Left            =   3480
         TabIndex        =   10
         Top             =   480
         Width           =   390
         _ExtentX        =   1111
         _ExtentY        =   1111
         StickyToggle    =   -1  'True
      End
   End
   Begin PhotoDemon.pdContainer cntrPopOut 
      Height          =   4935
      Index           =   2
      Left            =   8160
      Top             =   960
      Width           =   3705
      _ExtentX        =   6535
      _ExtentY        =   8705
      Begin PhotoDemon.pdCheckBox chkDistances 
         Height          =   330
         Left            =   240
         TabIndex        =   1
         Top             =   2790
         Width           =   3330
         _ExtentX        =   5874
         _ExtentY        =   582
         Caption         =   "show distances"
         Value           =   0   'False
      End
      Begin PhotoDemon.pdCheckBox chkLayerNodes 
         Height          =   330
         Left            =   240
         TabIndex        =   0
         Top             =   3180
         Width           =   3330
         _ExtentX        =   5874
         _ExtentY        =   582
         Caption         =   "show resize nodes"
      End
      Begin PhotoDemon.pdLabel lblOptions 
         Height          =   240
         Index           =   1
         Left            =   120
         Top             =   2460
         Width           =   3360
         _ExtentX        =   5927
         _ExtentY        =   423
         Caption         =   "display options"
      End
      Begin PhotoDemon.pdCheckBox chkRotateNode 
         Height          =   330
         Left            =   240
         TabIndex        =   2
         Top             =   3570
         Width           =   3330
         _ExtentX        =   5874
         _ExtentY        =   582
         Caption         =   "show rotate nodes"
      End
      Begin PhotoDemon.pdButtonToolbox cmdFlyoutLock 
         Height          =   390
         Index           =   2
         Left            =   3210
         TabIndex        =   12
         Top             =   4440
         Width           =   390
         _ExtentX        =   1111
         _ExtentY        =   1111
         StickyToggle    =   -1  'True
      End
   End
   Begin PhotoDemon.pdTitle ttlPanel 
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   3510
      _ExtentX        =   6191
      _ExtentY        =   661
      Caption         =   "position (x, y)"
      Value           =   0   'False
   End
   Begin PhotoDemon.pdSpinner tudLayerMove 
      Height          =   345
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   450
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   609
   End
   Begin PhotoDemon.pdSpinner tudLayerMove 
      Height          =   345
      Index           =   1
      Left            =   1680
      TabIndex        =   3
      Top             =   450
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   609
   End
   Begin PhotoDemon.pdTitle ttlPanel 
      Height          =   375
      Index           =   1
      Left            =   3720
      TabIndex        =   8
      Top             =   0
      Width           =   3510
      _ExtentX        =   6191
      _ExtentY        =   661
      Caption         =   "angle"
      Value           =   0   'False
   End
   Begin PhotoDemon.pdContainer cntrPopOut 
      Height          =   1335
      Index           =   1
      Left            =   3960
      Top             =   960
      Visible         =   0   'False
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   2355
      Begin PhotoDemon.pdButtonToolbox cmdFlyoutLock 
         Height          =   390
         Index           =   1
         Left            =   3600
         TabIndex        =   11
         Top             =   842
         Width           =   390
         _ExtentX        =   1111
         _ExtentY        =   1111
         StickyToggle    =   -1  'True
      End
   End
   Begin PhotoDemon.pdTitle ttlPanel 
      Height          =   375
      Index           =   2
      Left            =   7440
      TabIndex        =   9
      Top             =   0
      Width           =   3060
      _ExtentX        =   5398
      _ExtentY        =   661
      Caption         =   "other options"
      Value           =   0   'False
   End
End
Attribute VB_Name = "toolpanel_Crop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Crop Tool Panel
'Copyright 2024-2024 by Tanner Helland
'Created: 11/Nov/24
'Last updated: 11/Nov/24
'Last update: initial build
'
'This form includes all user-editable settings for the Crop on-canvas tool.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Flyout manager
Private WithEvents m_Flyout As pdFlyout
Attribute m_Flyout.VB_VarHelpID = -1

'The value of all controls on this form are saved and loaded to file by this class
Private WithEvents m_lastUsedSettings As pdLastUsedSettings
Attribute m_lastUsedSettings.VB_VarHelpID = -1

'Show/hide layer borders while using the move tool
Private Sub chkDistances_Click()
    Tools_Move.SetDrawDistances chkDistances.Value
    Viewport.Stage4_FlipBufferAndDrawUI PDImages.GetActiveImage(), FormMain.MainCanvas(0)
End Sub

Private Sub chkDistances_GotFocusAPI()
    UpdateFlyout 2, True
End Sub

Private Sub chkDistances_SetCustomTabTarget(ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If shiftTabWasPressed Then
        'newTargetHwnd = Me.btsCopyCut.hWnd
    Else
        newTargetHwnd = Me.chkLayerNodes.hWnd
    End If
End Sub

'Show/hide layer transform nodes while using the move tool
Private Sub chkLayerNodes_Click()
    Tools_Move.SetDrawLayerCornerNodes chkLayerNodes.Value
    Viewport.Stage4_FlipBufferAndDrawUI PDImages.GetActiveImage(), FormMain.MainCanvas(0)
End Sub

Private Sub chkLayerNodes_GotFocusAPI()
    UpdateFlyout 2, True
End Sub

Private Sub chkLayerNodes_SetCustomTabTarget(ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If shiftTabWasPressed Then
        newTargetHwnd = chkDistances.hWnd
    Else
        newTargetHwnd = chkRotateNode.hWnd
    End If
End Sub

Private Sub chkRotateNode_Click()
    Tools_Move.SetDrawLayerRotateNodes chkRotateNode.Value
    Viewport.Stage4_FlipBufferAndDrawUI PDImages.GetActiveImage(), FormMain.MainCanvas(0)
End Sub

Private Sub chkRotateNode_GotFocusAPI()
    UpdateFlyout 2, True
End Sub

Private Sub chkRotateNode_SetCustomTabTarget(ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If shiftTabWasPressed Then
        newTargetHwnd = chkLayerNodes.hWnd
    Else
        newTargetHwnd = Me.cmdFlyoutLock(2).hWnd
    End If
End Sub

Private Sub cmdFlyoutLock_Click(Index As Integer, ByVal Shift As ShiftConstants)
    If (Not m_Flyout Is Nothing) Then m_Flyout.UpdateLockStatus Me.cntrPopOut(Index).hWnd, cmdFlyoutLock(Index).Value, cmdFlyoutLock(Index)
End Sub

Private Sub cmdFlyoutLock_GotFocusAPI(Index As Integer)
    UpdateFlyout Index, True
End Sub

Private Sub cmdFlyoutLock_SetCustomTabTarget(Index As Integer, ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    Select Case Index
        Case 0
            'If shiftTabWasPressed Then newTargetHwnd = Me.cboLayerResizeQuality.hWnd Else newTargetHwnd = Me.ttlPanel(1).hWnd
        Case 1
            'If shiftTabWasPressed Then newTargetHwnd = Me.sltLayerShearY.hWndSpinner Else newTargetHwnd = Me.ttlPanel(2).hWnd
        Case 2
            If shiftTabWasPressed Then
                newTargetHwnd = Me.chkRotateNode.hWnd
            Else
                newTargetHwnd = Me.ttlPanel(0).hWnd
            End If
    End Select
End Sub

Private Sub Form_Load()
    
    Tools.SetToolBusyState True
    
    'TODO:
    'Ensure our corresponding tool manager is synchronized with default layer rendering styles
    'Tools_Move.SetDrawDistances chkDistances.Value
    'Tools_Move.SetDrawLayerCornerNodes chkLayerNodes.Value
    'Tools_Move.SetDrawLayerRotateNodes chkRotateNode.Value
    'Tools_Move.SetMoveSelectedPixels_SampleMerged (btsSampleMerged.ListIndex = 0)
    'Tools_Move.SetMoveSelectedPixels_DefaultCut (btsCopyCut.ListIndex = 1)
    
    'Load any last-used settings for this form
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

'Whenever an active flyout panel is closed, we need to reset the matching titlebar to "closed" state
Private Sub m_Flyout_FlyoutClosed(origTriggerObject As Control)
    If (Not origTriggerObject Is Nothing) Then origTriggerObject.Value = False
End Sub

'Because this tool synchronizes the vast majority of its properties to the active layer
' in the current image, we do *not* automatically save all settings for this tool.
' Instead, we custom-save only the settings we absolutely want.
Private Sub m_LastUsedSettings_AddCustomPresetData()

    'TODO:
    With m_lastUsedSettings
        '.AddPresetData "move-size-auto-activate", chkAutoActivateLayer.Value
        '.AddPresetData "move-size-ignore-transparent", chkIgnoreTransparent.Value
        '.AddPresetData "move-size-selection-sample-merged", btsSampleMerged.ListIndex
        '.AddPresetData "move-size-selection-default-cut", btsCopyCut.ListIndex
        '.AddPresetData "move-size-show-distances", chkDistances.Value
        '.AddPresetData "move-size-show-resize-nodes", chkLayerNodes.Value
        '.AddPresetData "move-size-show-rotate-nodes", chkRotateNode.Value
        '.AddPresetData "move-size-lock-aspect-ratio", chkAspectRatio.Value
    End With

End Sub

Private Sub m_LastUsedSettings_ReadCustomPresetData()

'TODO:
    With m_lastUsedSettings
'        chkAutoActivateLayer.Value = .RetrievePresetData("move-size-auto-activate", True)
'        chkIgnoreTransparent.Value = .RetrievePresetData("move-size-ignore-transparent", True)
'
'        btsSampleMerged.ListIndex = CLng(.RetrievePresetData("move-size-selection-sample-merged", "1"))
'        btsCopyCut.ListIndex = CLng(.RetrievePresetData("move-size-selection-default-cut", "1"))
'
'        chkDistances.Value = .RetrievePresetData("move-size-show-distances", False)
'        chkLayerNodes.Value = .RetrievePresetData("move-size-show-resize-nodes", True)
'        chkRotateNode.Value = .RetrievePresetData("move-size-show-rotate-nodes", True)
'
'        'The "lock aspect ratio" control is tricky, because we don't want to set this value
'        ' if the current image has variable aspect ratio enabled; otherwise, it will forcibly
'        ' modify the active layer's size value(s).
'        Dim okToLoad As Boolean
'        okToLoad = True
'
'        If Tools.CanvasToolsAllowed(False) Then
'            okToLoad = (PDImages.GetActiveImage.GetActiveLayer.GetLayerCanvasXModifier() = 1#) And (PDImages.GetActiveImage.GetActiveLayer.GetLayerCanvasYModifier() = 1#)
'        End If
'
'        If okToLoad Then
'            chkAspectRatio.Value = .RetrievePresetData("move-size-lock-aspect-ratio", False)
'        Else
'            chkAspectRatio.Value = False
'        End If
'
    End With
    
End Sub

Private Sub ttlPanel_Click(Index As Integer, ByVal newState As Boolean)
    UpdateFlyout Index, newState
End Sub

Private Sub ttlPanel_SetCustomTabTarget(Index As Integer, ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If shiftTabWasPressed Then
        Select Case Index
            Case 0
                newTargetHwnd = Me.cmdFlyoutLock(2).hWnd
            Case 1
                newTargetHwnd = Me.cmdFlyoutLock(0).hWnd
            Case 2
                newTargetHwnd = Me.cmdFlyoutLock(1).hWnd
        End Select
    Else
        Select Case Index
            Case 0
                newTargetHwnd = Me.tudLayerMove(0).hWnd
            Case 1
                'newTargetHwnd = Me.sltLayerAngle.hWndSlider
            Case 2
                'newTargetHwnd = Me.chkAutoActivateLayer.hWnd
        End Select
    End If
End Sub

Private Sub tudLayerMove_Change(Index As Integer)
    
    'If tool changes are not allowed, exit.
    ' NOTE: this will also check tool busy status, via Tools.getToolBusyState
    If (Not Tools.CanvasToolsAllowed) Then Exit Sub
    
    'Mark the tool engine as busy
    Tools.SetToolBusyState True
    
'    Select Case Index
'
'        'Layer position (x)
'        Case 0
'            PDImages.GetActiveImage.GetActiveLayer.SetLayerOffsetX tudLayerMove(Index).Value
'
'        'Layer position (y)
'        Case 1
'            PDImages.GetActiveImage.GetActiveLayer.SetLayerOffsetY tudLayerMove(Index).Value
'
'        'Layer width
'        Case 2
'            PDImages.GetActiveImage.GetActiveLayer.SetLayerCanvasXModifier tudLayerMove(Index).Value / PDImages.GetActiveImage.GetActiveLayer.GetLayerWidth(False)
'            If chkAspectRatio.Value Then
'                PDImages.GetActiveImage.GetActiveLayer.SetLayerCanvasYModifier PDImages.GetActiveImage.GetActiveLayer.GetLayerCanvasXModifier()
'                toolpanel_MoveSize.tudLayerMove(3).Value = PDImages.GetActiveImage.GetActiveLayer.GetLayerHeight(True)
'            End If
'
'        'Layer height
'        Case 3
'            PDImages.GetActiveImage.GetActiveLayer.SetLayerCanvasYModifier tudLayerMove(Index).Value / PDImages.GetActiveImage.GetActiveLayer.GetLayerHeight(False)
'            If chkAspectRatio.Value Then
'                PDImages.GetActiveImage.GetActiveLayer.SetLayerCanvasXModifier PDImages.GetActiveImage.GetActiveLayer.GetLayerCanvasYModifier()
'                toolpanel_MoveSize.tudLayerMove(2).Value = PDImages.GetActiveImage.GetActiveLayer.GetLayerWidth(True)
'            End If
'
'    End Select
    
    'Free the tool engine
    Tools.SetToolBusyState False
    
    'Redraw the viewport
    Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    
End Sub

Private Sub tudLayerMove_GotFocusAPI(Index As Integer)
    UpdateFlyout 0, True
End Sub

Private Sub tudLayerMove_LostFocusAPI(Index As Integer)
    If (Not PDImages.IsImageActive()) Then Exit Sub
End Sub

'Because these controls are laid out in a non-standard pattern, we want to manually specify tab and
' shift+tab focus targets.
Private Sub tudLayerMove_SetCustomTabTarget(Index As Integer, ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    
    If shiftTabWasPressed Then
        If (Index > 0) Then
            newTargetHwnd = tudLayerMove(Index - 1).hWnd
        Else
            newTargetHwnd = Me.ttlPanel(0).hWnd
        End If
    Else
        If (Index < 3) Then
            newTargetHwnd = tudLayerMove(Index + 1).hWnd
        Else
            'newTargetHwnd = chkAspectRatio.hWnd
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
    
    'TODO:
    
    'UI images must be updated against theme-specific colors
    Dim buttonSize As Long
    buttonSize = Interface.FixDPI(32)
    'cmdLayerAffinePermanent.AssignImage "generic_commit", , buttonSize, buttonSize
    
    'Tooltips must be localized
    'cmdLayerAffinePermanent.AssignTooltip "Make current layer transforms (size, angle, and shear) permanent.  This action is never required, but if viewport rendering is sluggish, it may improve performance."
    
    'Flyout lock controls use the same behavior across all instances
    UserControls.ThemeFlyoutControls cmdFlyoutLock
    
    Interface.ApplyThemeAndTranslations Me
    
End Sub

'Update the actively displayed flyout (if any).  Note that the flyout manager will automatically
' hide any other open flyouts, as necessary.
Private Sub UpdateFlyout(ByVal flyoutIndex As Long, Optional ByVal newState As Boolean = True)
    
    'Ensure we have a flyout manager
    If (m_Flyout Is Nothing) Then Set m_Flyout = New pdFlyout
    
    'Exit if we're already in the process of synchronizing
    If m_Flyout.GetFlyoutSyncState() Then Exit Sub
    m_Flyout.SetFlyoutSyncState True
    
    'Ensure we have a flyout manager, then raise the corresponding panel
    If newState Then
        If (flyoutIndex <> m_Flyout.GetFlyoutTrackerID()) Then m_Flyout.ShowFlyout Me, ttlPanel(flyoutIndex), cntrPopOut(flyoutIndex), flyoutIndex, IIf(flyoutIndex > 0, Interface.FixDPI(-8), 0)
    Else
        If (flyoutIndex = m_Flyout.GetFlyoutTrackerID()) Then m_Flyout.HideFlyout
    End If
    
    'Update titlebar state(s) to match
    Dim i As Long
    For i = ttlPanel.lBound To ttlPanel.UBound
        If (i = m_Flyout.GetFlyoutTrackerID()) Then
            If (Not ttlPanel(i).Value) Then ttlPanel(i).Value = True
        Else
            If ttlPanel(i).Value Then ttlPanel(i).Value = False
        End If
    Next i
    
    'Clear the synchronization flag before exiting
    m_Flyout.SetFlyoutSyncState False
    
End Sub
